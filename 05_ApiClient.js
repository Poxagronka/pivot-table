let BUNDLE_ID_CACHE = new Map();
let BUNDLE_ID_CACHE_LOADED = false;
let BUNDLE_ID_CACHE_TIME = null;
const BUNDLE_ID_CACHE_DURATION = 3600000;

let APPS_DB_CACHE = null;
let APPS_DB_CACHE_TIME = null;

function fetchCampaignData(dateRange) {
  const config = getCurrentConfig();
  const query = getGraphQLQuery();
  
  const variables = {
    dateFilters: [{
      key: config.API_CONFIG.DATE_DIMENSION,
      operator: "BETWEEN",
      values: [dateRange.from, dateRange.to]
    }],
    filters: [
      {
        key: "USER",
        operator: "IN",
        values: config.API_CONFIG.FILTERS.USER
      },
      {
        key: "ATTRIBUTION_NETWORK_HID",
        operator: "IN",
        values: config.API_CONFIG.FILTERS.ATTRIBUTION_NETWORK_HID
      }
    ],
    groupBy: config.API_CONFIG.GROUP_BY,
    measures: config.API_CONFIG.MEASURES,
    havingFilters: []
  };

  if (config.API_CONFIG.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH) {
    variables.filters.push({
      key: "ATTRIBUTION_CAMPAIGN_NAME",
      operator: config.API_CONFIG.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH.startsWith('!') ? 'DOES_NOT_MATCH_REGEX' : 'MATCHES_REGEX',
      values: [config.API_CONFIG.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH.replace(/^!/, '')]
    });
  }

  const projectName = CURRENT_PROJECT;
  const startTime = Date.now();
  
  const payload = {
    query: query,
    variables: variables
  };

  const maxRetries = 3;
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = UrlFetchApp.fetch('https://api.appodeal.com/graphql', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${config.BEARER_TOKEN}`
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      
      const apiTime = Date.now() - startTime;
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode === 200) {
        try {
          const data = JSON.parse(responseText);
          
          if (data.errors && data.errors.length > 0) {
            console.error(`${projectName}: GraphQL errors:`, data.errors);
            throw new Error(`GraphQL errors: ${data.errors.map(e => e.message).join(', ')}`);
          }
          
          if (!data.data || !data.data.analytics || !data.data.analytics.richStats) {
            throw new Error('Invalid API response structure');
          }
          
          const recordCount = data.data.analytics.richStats.stats ? data.data.analytics.richStats.stats.length : 0;
          logDebugTiming({ api: apiTime });
          
          return data;
        } catch (parseError) {
          console.error(`${projectName}: Error parsing JSON response:`, parseError);
          throw new Error(`JSON parsing failed: ${parseError.message}`);
        }
      }
      
      if (responseCode === 400) {
        console.error(`${projectName}: Bad request (400):`, responseText.substring(0, 1000));
        throw new Error(`Bad request: ${responseText.substring(0, 200)}`);
      }
      
      if (responseCode === 401) {
        console.error(`${projectName}: Unauthorized (401) - check Bearer token`);
        throw new Error('Unauthorized - check Bearer token');
      }
      
      if (responseCode === 403) {
        console.error(`${projectName}: Forbidden (403) - insufficient permissions`);
        throw new Error('Forbidden - insufficient permissions');
      }
      
      if (responseCode === 429) {
        console.error(`${projectName}: Rate limited (429)`);
        if (attempt < maxRetries) {
          const delay = 60000;
          Utilities.sleep(delay);
          continue;
        } else {
          throw new Error('Too many requests');
        }
      } else if (responseCode >= 400 && responseCode < 500) {
        if (responseText.includes('Rate limit exceeded')) {
          console.error(`${projectName}: Rate limit detected in response body`);
          if (attempt < maxRetries) {
            const delay = 60000;
            Utilities.sleep(delay);
            continue;
          } else {
            throw new Error('Too many requests');
          }
        } else {
          throw new Error(`Client error ${responseCode}: ${responseText.substring(0, 200)}`);
        }
      }
      
      if (responseCode >= 500) {
        const errorMsg = `Server error ${responseCode}`;
        console.error(`${projectName}: ${errorMsg}, attempt ${attempt}/${maxRetries}`);
        
        lastError = new Error(`${errorMsg}: Server returned HTML error page`);
        
        if (attempt < maxRetries) {
          const delay = Math.min(1000 * Math.pow(2, attempt - 1), 10000);
          Utilities.sleep(delay);
          continue;
        }
      }
      
      lastError = new Error(`Unexpected response code ${responseCode}: ${responseText.substring(0, 200)}`);
      
    } catch (e) {
      console.error(`${projectName}: Request attempt ${attempt} failed:`, e);
      lastError = e;
      
      if (attempt < maxRetries) {
        const delay = Math.min(2000 * attempt, 10000);
        Utilities.sleep(delay);
      }
    }
  }
  
  console.error(`${projectName}: All API request attempts failed`);
  throw lastError || new Error('API request failed after all retries');
}

function getGraphQLQuery() {
  return `query RichStats($dateFilters: [DateFilterInput!]!, $filters: [FilterInput!]!, $groupBy: [GroupByInput!]!, $measures: [RichMeasureInput!]!, $havingFilters: [HavingFilterInput!]!) {
    analytics {
      richStats(
        dateFilters: $dateFilters
        filters: $filters
        groupBy: $groupBy
        measures: $measures
        havingFilters: $havingFilters
      ) {
        stats
      }
    }
  }`;
}

function extractBundleIdFromCampaign(campaignName) {
  if (!campaignName) return null;
  
  const patterns = [
    /id(\d+)/i,
    /bundle[_\-]?id[_\-]?(\d+)/i,
    /app[_\-]?id[_\-]?(\d+)/i,
    /(\d{9,})/
  ];
  
  for (const pattern of patterns) {
    const match = campaignName.match(pattern);
    if (match) {
      return match[1];
    }
  }
  
  return null;
}

function ensureBundleIdCacheLoaded() {
  const now = new Date().getTime();
  
  if (BUNDLE_ID_CACHE_LOADED && BUNDLE_ID_CACHE_TIME && (now - BUNDLE_ID_CACHE_TIME) < BUNDLE_ID_CACHE_DURATION) {
    return;
  }
  
  try {
    const config = getProjectConfig('TRICKY');
    const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
    const cacheSheet = spreadsheet.getSheetByName('BundleIdCache');
    
    if (!cacheSheet) {
      BUNDLE_ID_CACHE.clear();
      BUNDLE_ID_CACHE_LOADED = true;
      BUNDLE_ID_CACHE_TIME = now;
      return;
    }
    
    const range = `BundleIdCache!A:C`;
    const response = Sheets.Spreadsheets.Values.get(config.SHEET_ID, range);
    const data = response.values || [];
    
    BUNDLE_ID_CACHE.clear();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i] && data[i][0] && data[i][1]) {
        BUNDLE_ID_CACHE.set(data[i][0], {
          bundleId: data[i][1],
          lastUsed: data[i][2] || new Date().toISOString()
        });
      }
    }
    
    BUNDLE_ID_CACHE_LOADED = true;
    BUNDLE_ID_CACHE_TIME = now;
    
  } catch (e) {
    console.error('Error loading Bundle ID Cache:', e);
    BUNDLE_ID_CACHE.clear();
    BUNDLE_ID_CACHE_LOADED = true;
    BUNDLE_ID_CACHE_TIME = now;
  }
}

function saveBundleIdCache(newEntries) {
  if (newEntries.size === 0) return;
  
  try {
    const config = getProjectConfig('TRICKY');
    const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
    let cacheSheet = spreadsheet.getSheetByName('BundleIdCache');
    
    if (!cacheSheet) {
      cacheSheet = spreadsheet.insertSheet('BundleIdCache');
      cacheSheet.hideSheet();
      cacheSheet.getRange('A1:C1').setValues([['Campaign Name', 'Bundle ID', 'Last Used']]);
    }
    
    const existingData = cacheSheet.getDataRange().getValues();
    const existingMap = new Map();
    
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0]) {
        existingMap.set(existingData[i][0], i + 1);
      }
    }
    
    const updates = [];
    const appends = [];
    
    newEntries.forEach((data, campaignName) => {
      const row = [campaignName, data.bundleId, new Date().toISOString()];
      
      if (existingMap.has(campaignName)) {
        const rowIndex = existingMap.get(campaignName);
        updates.push({ range: `A${rowIndex}:C${rowIndex}`, values: [row] });
      } else {
        appends.push(row);
      }
    });
    
    if (updates.length > 0) {
      updates.forEach(update => {
        cacheSheet.getRange(update.range).setValues(update.values);
      });
    }
    
    if (appends.length > 0) {
      const startRow = cacheSheet.getLastRow() + 1;
      cacheSheet.getRange(startRow, 1, appends.length, 3).setValues(appends);
    }
    
    newEntries.forEach((data, campaignName) => {
      BUNDLE_ID_CACHE.set(campaignName, data);
    });
    
  } catch (e) {
    console.error('Error saving Bundle ID Cache:', e);
  }
}

function getOptimizedAppsDbForTricky() {
  const now = new Date().getTime();
  
  if (APPS_DB_CACHE && APPS_DB_CACHE_TIME && (now - APPS_DB_CACHE_TIME) < BUNDLE_ID_CACHE_DURATION) {
    return APPS_DB_CACHE;
  }
  
  try {
    const appsDb = new AppsDatabase('TRICKY');
    appsDb.ensureCacheUpToDate();
    const cache = appsDb.loadFromCache();
    
    APPS_DB_CACHE = cache;
    APPS_DB_CACHE_TIME = now;
    
    return cache;
  } catch (e) {
    console.error('Error loading Apps Database:', e);
    return {};
  }
}

function getCachedBundleId(campaignName, campaignId = '') {
  if (BUNDLE_ID_CACHE.has(campaignName)) {
    return BUNDLE_ID_CACHE.get(campaignName).bundleId;
  }
  
  const bundleId = extractBundleIdFromCampaign(campaignName);
  return bundleId;
}

function getOptimizedSourceAppDisplayName(bundleId, appsDbCache) {
  if (!bundleId || CURRENT_PROJECT !== 'TRICKY' || !appsDbCache) {
    return bundleId || 'Unknown';
  }
  
  const appInfo = appsDbCache[bundleId];
  if (appInfo && appInfo.publisher !== bundleId) {
    const publisher = appInfo.publisher || '';
    const appName = appInfo.appName || '';
    
    if (publisher && appName && publisher !== appName) {
      return `${publisher} ${appName}`;
    } else if (publisher) {
      return publisher;
    } else if (appName) {
      return appName;
    }
  }
  
  return bundleId;
}

function processApiData(rawData, includeLastWeek = null) {
  const processingStartTime = Date.now();
  const stats = rawData.data.analytics.richStats.stats;
  const appData = {};

  const today = new Date();
  const currentWeekStart = formatDateForAPI(getMondayOfWeek(today));
  const lastWeekStart = formatDateForAPI(getMondayOfWeek(new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000)));

  const dayOfWeek = today.getDay();
  const shouldIncludeLastWeek = includeLastWeek !== null ? includeLastWeek : (dayOfWeek >= 2 || dayOfWeek === 0);

  let processedCount = 0;

  if (CURRENT_PROJECT === 'TRICKY') {
    const result = processTrickyDataOptimized(stats, currentWeekStart, lastWeekStart, shouldIncludeLastWeek);
    const processingTime = Date.now() - processingStartTime;
    logDebugTiming({ processing: processingTime });
    return result;
  }

  const BATCH_SIZE = 500;
  for (let batchStart = 0; batchStart < stats.length; batchStart += BATCH_SIZE) {
    const batchEnd = Math.min(batchStart + BATCH_SIZE, stats.length);
    const batch = stats.slice(batchStart, batchEnd);
    
    batch.forEach((row, index) => {
      try {
        const date = row[0].value;
        const monday = getMondayOfWeek(new Date(date));
        const weekKey = formatDateForAPI(monday);

        if (weekKey >= currentWeekStart) {
          return;
        }
        
        if (!shouldIncludeLastWeek && weekKey >= lastWeekStart) {
          return;
        }

        let campaign, app, network, metricsStartIndex;
        
        if (CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'INCENT_TRAFFIC') {
          campaign = null;
          network = row[1];
          app = row[2];
          metricsStartIndex = 3;
        } else {
          campaign = row[1];
          app = row[2];
          network = null;
          metricsStartIndex = 3;
        }
        
        const metrics = {
          cpi: parseFloat(row[metricsStartIndex].value) || 0,
          installs: parseInt(row[metricsStartIndex + 1].value) || 0,
          ipm: parseFloat(row[metricsStartIndex + 2].value) || 0,
          spend: parseFloat(row[metricsStartIndex + 3].value) || 0,
          rrD1: parseFloat(row[metricsStartIndex + 4].value) || 0,
          roasD1: parseFloat(row[metricsStartIndex + 5].value) || 0,
          roasD3: parseFloat(row[metricsStartIndex + 6].value) || 0,
          rrD7: parseFloat(row[metricsStartIndex + 7].value) || 0,
          roasD7: parseFloat(row[metricsStartIndex + 8].value) || 0,
          roasD30: parseFloat(row[metricsStartIndex + 9].value) || 0,
          eArpuForecast: parseFloat(row[metricsStartIndex + 10].value) || 0,
          eRoasForecast: parseFloat(row[metricsStartIndex + 11].value) || 0,
          eProfitForecast: parseFloat(row[metricsStartIndex + 12].value) || 0,
          eRoasForecastD730: parseFloat(row[metricsStartIndex + 13].value) || 0
        };

        const sunday = new Date(monday);
        sunday.setDate(monday.getDate() + 6);

        if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
          const networkKey = `${network.networkName}_${network.networkId}`;
          if (!appData[networkKey]) {
            appData[networkKey] = {
              networkId: network.networkId,
              networkName: network.networkName,
              weeks: {}
            };
          }

          if (!appData[networkKey].weeks[weekKey]) {
            appData[networkKey].weeks[weekKey] = {
              weekStart: formatDateForAPI(monday),
              weekEnd: formatDateForAPI(sunday),
              apps: {}
            };
          }

          if (!appData[networkKey].weeks[weekKey].apps[app.appId]) {
            appData[networkKey].weeks[weekKey].apps[app.appId] = {
              appId: app.appId,
              appName: app.appName,
              platform: app.platform,
              bundleId: app.bundleId,
              campaigns: []
            };
          }

          if (campaign) {
            const campaignData = {
              campaignId: campaign.campaignId,
              campaignName: campaign.campaignName,
              ...metrics
            };
            appData[networkKey].weeks[weekKey].apps[app.appId].campaigns.push(campaignData);
          }

        } else if (CURRENT_PROJECT === 'OVERALL') {
          const appKey = `${app.appName}_${app.appId}`;
          if (!appData[appKey]) {
            appData[appKey] = {
              appId: app.appId,
              appName: app.appName,
              platform: app.platform,
              bundleId: app.bundleId,
              weeks: {}
            };
          }

          if (!appData[appKey].weeks[weekKey]) {
            appData[appKey].weeks[weekKey] = {
              weekStart: formatDateForAPI(monday),
              weekEnd: formatDateForAPI(sunday),
              networks: {}
            };
          }

          const networkKey = `${network.networkName}_${network.networkId}`;
          if (!appData[appKey].weeks[weekKey].networks[networkKey]) {
            appData[appKey].weeks[weekKey].networks[networkKey] = {
              networkId: network.networkId,
              networkName: network.networkName,
              campaigns: []
            };
          }

          const campaignData = {
            campaignId: 'aggregated',
            campaignName: 'Aggregated',
            ...metrics
          };
          appData[appKey].weeks[weekKey].networks[networkKey].campaigns.push(campaignData);

        } else {
          const appKey = `${app.appName}_${app.appId}`;
          if (!appData[appKey]) {
            appData[appKey] = {
              appId: app.appId,
              appName: app.appName,
              platform: app.platform,
              bundleId: app.bundleId,
              weeks: {}
            };
          }

          if (!appData[appKey].weeks[weekKey]) {
            appData[appKey].weeks[weekKey] = {
              weekStart: formatDateForAPI(monday),
              weekEnd: formatDateForAPI(sunday),
              campaigns: []
            };
          }

          const campaignData = {
            campaignId: campaign.campaignId,
            campaignName: campaign.campaignName,
            ...metrics
          };
          appData[appKey].weeks[weekKey].campaigns.push(campaignData);
        }

        processedCount++;

      } catch (error) {
        console.error(`Error processing row ${batchStart + index}:`, error);
      }
    });
  }

  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    Object.values(appData).forEach(network => {
      Object.values(network.weeks).forEach(week => {
        Object.values(week.apps).forEach(app => {
          if (app.campaigns) {
            app.campaigns.sort((a, b) => b.spend - a.spend);
          }
        });
      });
    });
    
    const processingTime = Date.now() - processingStartTime;
    logDebugTiming({ processing: processingTime });
    return appData;
  }

  const processingTime = Date.now() - processingStartTime;
  logDebugTiming({ processing: processingTime });
  return appData;
}

function processTrickyDataOptimized(stats, currentWeekStart, lastWeekStart, shouldIncludeLastWeek) {
  const processingStartTime = Date.now();
  
  ensureBundleIdCacheLoaded();
  const appsDbCache = getOptimizedAppsDbForTricky();
  
  const appData = {};
  const newBundleIds = new Map();
  const bundleIdToDisplayName = new Map();
  let processedCount = 0;

  const BATCH_SIZE = 1000;
  for (let batchStart = 0; batchStart < stats.length; batchStart += BATCH_SIZE) {
    const batchEnd = Math.min(batchStart + BATCH_SIZE, stats.length);
    const batch = stats.slice(batchStart, batchEnd);
    
    batch.forEach((row, index) => {
      try {
        const date = row[0].value;
        const monday = getMondayOfWeek(new Date(date));
        const weekKey = formatDateForAPI(monday);

        if (weekKey >= currentWeekStart) {
          return;
        }
        
        if (!shouldIncludeLastWeek && weekKey >= lastWeekStart) {
          return;
        }

        const campaign = row[1];
        const app = row[2];
        const metricsStartIndex = 3;
        
        const metrics = {
          cpi: parseFloat(row[metricsStartIndex].value) || 0,
          installs: parseInt(row[metricsStartIndex + 1].value) || 0,
          ipm: parseFloat(row[metricsStartIndex + 2].value) || 0,
          spend: parseFloat(row[metricsStartIndex + 3].value) || 0,
          rrD1: parseFloat(row[metricsStartIndex + 4].value) || 0,
          roasD1: parseFloat(row[metricsStartIndex + 5].value) || 0,
          roasD3: parseFloat(row[metricsStartIndex + 6].value) || 0,
          rrD7: parseFloat(row[metricsStartIndex + 7].value) || 0,
          roasD7: parseFloat(row[metricsStartIndex + 8].value) || 0,
          roasD30: parseFloat(row[metricsStartIndex + 9].value) || 0,
          eArpuForecast: parseFloat(row[metricsStartIndex + 10].value) || 0,
          eRoasForecast: parseFloat(row[metricsStartIndex + 11].value) || 0,
          eProfitForecast: parseFloat(row[metricsStartIndex + 12].value) || 0,
          eRoasForecastD730: parseFloat(row[metricsStartIndex + 13].value) || 0
        };

        let bundleId = getCachedBundleId(campaign.campaignName);
        
        if (!bundleId) {
          bundleId = extractBundleIdFromCampaign(campaign.campaignName);
          if (bundleId) {
            newBundleIds.set(campaign.campaignName, {
              bundleId: bundleId,
              lastUsed: new Date().toISOString()
            });
          }
        }

        let sourceAppDisplayName;
        if (bundleIdToDisplayName.has(bundleId)) {
          sourceAppDisplayName = bundleIdToDisplayName.get(bundleId);
        } else {
          sourceAppDisplayName = getOptimizedSourceAppDisplayName(bundleId, appsDbCache);
          if (bundleId) {
            bundleIdToDisplayName.set(bundleId, sourceAppDisplayName);
          }
        }

        const campaignData = {
          campaignId: campaign.campaignId,
          campaignName: campaign.campaignName,
          ...metrics
        };

        const sunday = new Date(monday);
        sunday.setDate(monday.getDate() + 6);

        const appKey = `${app.appName}_${app.appId}`;
        if (!appData[appKey]) {
          appData[appKey] = {
            appId: app.appId,
            appName: app.appName,
            platform: app.platform,
            bundleId: app.bundleId,
            weeks: {}
          };
        }

        if (!appData[appKey].weeks[weekKey]) {
          appData[appKey].weeks[weekKey] = {
            weekStart: formatDateForAPI(monday),
            weekEnd: formatDateForAPI(sunday),
            sourceApps: {}
          };
        }

        const finalBundleId = bundleId || 'unknown';
        if (!appData[appKey].weeks[weekKey].sourceApps[finalBundleId]) {
          appData[appKey].weeks[weekKey].sourceApps[finalBundleId] = {
            sourceAppId: finalBundleId,
            sourceAppName: sourceAppDisplayName,
            campaigns: []
          };
        }
        
        appData[appKey].weeks[weekKey].sourceApps[finalBundleId].campaigns.push(campaignData);
        processedCount++;

      } catch (error) {
        console.error(`Error processing TRICKY row ${batchStart + index}:`, error);
      }
    });
  }

  Object.values(appData).forEach(app => {
    Object.values(app.weeks).forEach(week => {
      Object.values(week.sourceApps).forEach(sourceApp => {
        sourceApp.campaigns.sort((a, b) => b.spend - a.spend);
      });
    });
  });

  if (newBundleIds.size > 0) {
    try {
      saveBundleIdCache(newBundleIds);
    } catch (e) {
      console.error('Error saving Bundle ID cache:', e);
    }
  }

  const processingTime = Date.now() - processingStartTime;
  logDebugTiming({ processing: processingTime });
  return appData;
}

function processProjectApiData(projectName, rawData, includeLastWeek = null) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    const result = processApiData(rawData, includeLastWeek);
    return result;
  } finally {
    setCurrentProject(originalProject);
  }
}

function extractGeoFromCampaign(campaignName) {
  if (!campaignName) return 'OTHER';
  
  if (CURRENT_PROJECT === 'TRICKY' || CURRENT_PROJECT === 'REGULAR') {
    const geoMap = {
      '| USA |': 'USA', '| MEX |': 'MEX', '| AUS |': 'AUS', '| DEU |': 'DEU',
      '| JPN |': 'JPN', '| KOR |': 'KOR', '| BRA |': 'BRA', '| CAN |': 'CAN', '| GBR |': 'GBR',
      '| FRA |': 'FRA', '| ITA |': 'ITA', '| ESP |': 'ESP', '| NLD |': 'NLD', '| POL |': 'POL',
      '| TUR |': 'TUR', '| RUS |': 'RUS', '| IND |': 'IND', '| THA |': 'THA', '| IDN |': 'IDN',
      '| VNM |': 'VNM', '| MYS |': 'MYS', '| PHL |': 'PHL', '| SGP |': 'SGP', '| ARE |': 'ARE',
      '| SAU |': 'SAU', '| EGY |': 'EGY', '| ZAF |': 'ZAF', '| NGA |': 'NGA', '| KEN |': 'KEN',
      '| ETH |': 'ETH', '| GHA |': 'GHA', '| UGA |': 'UGA', '| TZA |': 'TZA', '| RWA |': 'RWA',
      '| SEN |': 'SEN', '| MDG |': 'MDG', '| MOZ |': 'MOZ', '| ZWE |': 'ZWE', '| ZMB |': 'ZMB',
      '| BWA |': 'BWA', '| NAM |': 'NAM', '| AGO |': 'AGO', '| CMR |': 'CMR', '| CIV |': 'CIV',
      '| MLI |': 'MLI', '| BFA |': 'BFA', '| NER |': 'NER', '| TCD |': 'TCD', '| SDN |': 'SDN',
      '| LBY |': 'LBY', '| TUN |': 'TUN', '| DZA |': 'DZA', '| MAR |': 'MAR', '| ARG |': 'ARG',
      '| CHL |': 'CHL', '| COL |': 'COL', '| PER |': 'PER', '| ECU |': 'ECU', '| BOL |': 'BOL',
      '| URY |': 'URY', '| PRY |': 'PRY', '| VEN |': 'VEN', '| GUY |': 'GUY', '| SUR |': 'SUR',
      '| GUF |': 'GUF', '| PAN |': 'PAN', '| CRI |': 'CRI', '| NIC |': 'NIC', '| HND |': 'HND',
      '| GTM |': 'GTM', '| BLZ |': 'BLZ', '| SLV |': 'SLV', '| DOM |': 'DOM', '| HTI |': 'HTI',
      '| JAM |': 'JAM', '| CUB |': 'CUB', '| PRI |': 'PRI', '| TTO |': 'TTO', '| GRD |': 'GRD',
      '| LCA |': 'LCA', '| VCT |': 'VCT', '| DMA |': 'DMA', '| ATG |': 'ATG', '| KNA |': 'KNA',
      '| BRB |': 'BRB', '| ABW |': 'ABW', '| CUW |': 'CUW', '| SXM |': 'SXM', '| BES |': 'BES',
      '| CHN |': 'CHN', '| TWN |': 'TWN', '| HKG |': 'HKG', '| MAC |': 'MAC'
    };
    
    for (const [pattern, country] of Object.entries(geoMap)) {
      if (campaignName.includes(pattern)) {
        return country;
      }
    }
    
    const geoPatterns = ['USA', 'MEX', 'AUS', 'DEU', 'JPN', 'KOR', 'BRA', 'CAN', 'GBR', 'FRA', 'ITA', 'ESP', 'NLD', 'POL', 'TUR', 'RUS', 'IND', 'THA', 'IDN', 'VNM', 'MYS', 'PHL', 'SGP', 'ARE', 'SAU', 'EGY', 'ZAF', 'NGA', 'KEN', 'ETH', 'GHA', 'UGA', 'TZA', 'RWA', 'SEN', 'MDG', 'MOZ', 'ZWE', 'ZMB', 'BWA', 'NAM', 'AGO', 'CMR', 'CIV', 'MLI', 'BFA', 'NER', 'TCD', 'SDN', 'LBY', 'TUN', 'DZA', 'MAR', 'ARG', 'CHL', 'COL', 'PER', 'ECU', 'BOL', 'URY', 'PRY', 'VEN', 'GUY', 'SUR', 'GUF', 'PAN', 'CRI', 'NIC', 'HND', 'GTM', 'BLZ', 'SLV', 'DOM', 'HTI', 'JAM', 'CUB', 'PRI', 'TTO', 'GRD', 'LCA', 'VCT', 'DMA', 'ATG', 'KNA', 'BRB', 'ABW', 'CUW', 'SXM', 'BES', 'CHN', 'TWN', 'HKG', 'MAC'];
    
    const upperCampaignName = campaignName.toUpperCase();
    
    for (const pattern of geoPatterns) {
      const upperPattern = pattern.toUpperCase();
      
      if (upperCampaignName.includes('_' + upperPattern + '_') || upperCampaignName.includes('-' + upperPattern + '-') ||
          upperCampaignName.includes('_' + upperPattern) || upperCampaignName.includes('-' + upperPattern) ||
          upperCampaignName.includes(upperPattern + '_') || upperCampaignName.includes(upperPattern + '-') ||
          upperCampaignName === upperPattern) {
        return pattern;
      }
    }
    
    return 'OTHER';
  }
  
  if (CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    return 'ALL';
  }
  
  if (CURRENT_PROJECT === 'GOOGLE_ADS') {
    const geoPatterns = [
      { pattern: 'LatAm', geo: 'LatAm' },
      { pattern: 'UK,GE', geo: 'UK,GE' },
      { pattern: 'BR (PT)', geo: 'BR' },
      { pattern: 'US ', geo: 'US' },
      { pattern: ' US ', geo: 'US' },
      { pattern: 'WW ', geo: 'WW' },
      { pattern: ' WW ', geo: 'WW' },
      { pattern: 'UK', geo: 'UK' },
      { pattern: 'GE', geo: 'GE' },
      { pattern: 'BR', geo: 'BR' }
    ];
    
    for (const {pattern, geo} of geoPatterns) {
      if (campaignName.includes(pattern)) {
        return geo;
      }
    }
    return 'OTHER';
  }
  
  const geoPatterns = ['WW_ru', 'WW_es', 'WW_de', 'WW_pt', 'Asia T1', 'T2-ES', 'T1-EN', 'LatAm', 'TopGeo', 'Europe', 'US', 'RU', 'UK', 'GE', 'FR', 'PT', 'ES', 'DE', 'T1', 'WW'];
  const upperCampaignName = campaignName.toUpperCase();
  
  for (const pattern of geoPatterns) {
    const upperPattern = pattern.toUpperCase();
    
    if (upperCampaignName.includes('_' + upperPattern + '_') || upperCampaignName.includes('-' + upperPattern + '-') ||
        upperCampaignName.includes('_' + upperPattern) || upperCampaignName.includes('-' + upperPattern) ||
        upperCampaignName.includes(upperPattern + '_') || upperCampaignName.includes(upperPattern + '-') ||
        upperCampaignName === upperPattern) {
      return pattern;
    }
  }
  
  return 'OTHER';
}

function extractSourceApp(campaignName) {
  try {
    if (CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'INCENT_TRAFFIC') {
      return campaignName;
    }
    
    if (campaignName.startsWith('APD_')) {
      return campaignName;
    }
    
    if (CURRENT_PROJECT === 'REGULAR' || CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN' || CURRENT_PROJECT === 'MINTEGRAL' || CURRENT_PROJECT === 'INCENT') {
      return campaignName;
    }
    
    const eq = campaignName.indexOf('=');
    if (eq !== -1) {
      let t = campaignName.substring(eq + 1).trim();
      const subs = [];
      let idx = t.indexOf('subj');
      while (idx !== -1) {
        subs.push(idx);
        idx = t.indexOf('subj', idx + 1);
      }
      if (subs.length >= 2) t = t.substring(0, subs[1]).trim();
      else if (subs.length === 1 && subs[0] > 10) t = t.substring(0, subs[0]).trim();
      t = t.replace(/autobudget$/, '').trim();
      if (t) return t;
    }
    const lp = campaignName.lastIndexOf('|');
    if (lp !== -1) return campaignName.substring(lp + 1).trim();
    return 'Unknown';
  } catch (e) {
    return 'Unknown';
  }
}

function extractProjectSourceApp(projectName, campaignName) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    return extractSourceApp(campaignName);
  } finally {
    setCurrentProject(originalProject);
  }
}

function extractProjectGeoFromCampaign(projectName, campaignName) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    return extractGeoFromCampaign(campaignName);
  } finally {
    setCurrentProject(originalProject);
  }
}

function clearTrickyCaches() {
  BUNDLE_ID_CACHE.clear();
  BUNDLE_ID_CACHE_LOADED = false;
  BUNDLE_ID_CACHE_TIME = null;
  APPS_DB_CACHE = null;
  APPS_DB_CACHE_TIME = null;
  console.log('TRICKY caches cleared');
}