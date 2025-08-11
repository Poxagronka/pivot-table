var BUNDLE_ID_CACHE = new Map();
var BUNDLE_ID_CACHE_LOADED = false;
var BUNDLE_ID_CACHE_TIME = null;
var BUNDLE_ID_CACHE_DURATION = 21600000;
var APPS_DB_CACHE = null;
var APPS_DB_CACHE_TIME = null;
var BUNDLE_ID_CACHE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1Z5pJgtg--9EACJL8PVZgJsmeUemv6PKhSsyx9ArChrM/edit?gid=754371211#gid=754371211';
var BUNDLE_ID_CACHE_SHEET_ID = '1Z5pJgtg--9EACJL8PVZgJsmeUemv6PKhSsyx9ArChrM';

function fetchCampaignData(dateRange) {
  const startTime = Date.now();
  const config = getCurrentConfig();
  const apiConfig = getCurrentApiConfig();
  
  const filters = [
    { dimension: "USER", values: apiConfig.FILTERS.USER, include: true },
    { dimension: "ATTRIBUTION_PARTNER", values: apiConfig.FILTERS.ATTRIBUTION_PARTNER, include: true }
  ];
  
  if (apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID && apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.length > 0) {
    filters.push({ dimension: "ATTRIBUTION_NETWORK_HID", values: apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID, include: true });
  }
  
  if (apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH) {
    const searchPattern = apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH;
    
    if (searchPattern.startsWith('!')) {
      const excludePattern = searchPattern.substring(1);
      filters.push({
        dimension: "ATTRIBUTION_CAMPAIGN_HID", 
        values: [], 
        include: false,
        searchByString: excludePattern
      });
    } else {
      filters.push({
        dimension: "ATTRIBUTION_CAMPAIGN_HID", 
        values: [], 
        include: true, 
        searchByString: searchPattern
      });
    }
  }
  
  const dateDimension = (CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN' || CURRENT_PROJECT === 'INCENT' || CURRENT_PROJECT === 'INCENT_TRAFFIC' || CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'APPLOVIN_NEW') ? 'DATE' : 'INSTALL_DATE';
  
  const payload = {
    operationName: apiConfig.OPERATION_NAME,
    variables: {
      dateFilters: [{
        dimension: dateDimension,
        from: dateRange.from,
        to: dateRange.to,
        include: true
      }],
      filters: filters,
      groupBy: apiConfig.GROUP_BY,
      measures: apiConfig.MEASURES,
      havingFilters: [{ measure: { id: "spend", day: null }, operator: "MORE", value: 0 }],
      anonymizationMode: "OFF",
      topFilter: null,
      revenuePredictionVersion: "",
      isMultiMediation: true
    },
    query: getGraphQLQuery()
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Accept: 'application/json, text/plain, */*',
      'Accept-Language': 'en-US,en;q=0.9',
      Authorization: `Bearer ${config.BEARER_TOKEN}`,
      Connection: 'keep-alive',
      DNT: '1',
      Origin: 'https://app.appodeal.com',
      Referer: 'https://app.appodeal.com/analytics/reports?reloadTime=' + Date.now(),
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36',
      'x-requested-with': 'XMLHttpRequest',
      'Trace-Id': Utilities.getUuid()
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  return executeApiRequestWithRetry(config.API_URL, options, CURRENT_PROJECT, startTime);
}

function fetchProjectCampaignData(projectName, dateRange) {
  const startTime = Date.now();
  const config = getProjectConfig(projectName);
  const apiConfig = getProjectApiConfig(projectName);
  
  if (!config.BEARER_TOKEN) {
    throw new Error(`${projectName} project is not configured: missing BEARER_TOKEN`);
  }
  
  if (!apiConfig.FILTERS.USER || apiConfig.FILTERS.USER.length === 0) {
    throw new Error(`${projectName} project is not configured: missing USER filters`);
  }
  
  const filters = [
    { dimension: "USER", values: apiConfig.FILTERS.USER, include: true },
    { dimension: "ATTRIBUTION_PARTNER", values: apiConfig.FILTERS.ATTRIBUTION_PARTNER, include: true }
  ];
  
  if (apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID && apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.length > 0) {
    filters.push({ dimension: "ATTRIBUTION_NETWORK_HID", values: apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID, include: true });
  }
  
  if (apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH) {
    const searchPattern = apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH;
    
    if (searchPattern.startsWith('!')) {
      const excludePattern = searchPattern.substring(1);
      filters.push({
        dimension: "ATTRIBUTION_CAMPAIGN_HID", 
        values: [], 
        include: false,
        searchByString: excludePattern
      });
    } else {
      filters.push({
        dimension: "ATTRIBUTION_CAMPAIGN_HID", 
        values: [], 
        include: true, 
        searchByString: searchPattern
      });
    }
  }
  
  const dateDimension = (projectName === 'GOOGLE_ADS' || projectName === 'APPLOVIN' || projectName === 'INCENT' || projectName === 'INCENT_TRAFFIC' || projectName === 'OVERALL' || projectName === 'APPLOVIN_NEW') ? 'DATE' : 'INSTALL_DATE';
  
  const payload = {
    operationName: apiConfig.OPERATION_NAME,
    variables: {
      dateFilters: [{
        dimension: dateDimension,
        from: dateRange.from,
        to: dateRange.to,
        include: true
      }],
      filters: filters,
      groupBy: apiConfig.GROUP_BY,
      measures: apiConfig.MEASURES,
      havingFilters: [{ measure: { id: "spend", day: null }, operator: "MORE", value: 0 }],
      anonymizationMode: "OFF",
      topFilter: null,
      revenuePredictionVersion: "",
      isMultiMediation: true
    },
    query: getGraphQLQuery()
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Accept: 'application/json, text/plain, */*',
      'Accept-Language': 'en-US,en;q=0.9',
      Authorization: `Bearer ${config.BEARER_TOKEN}`,
      Connection: 'keep-alive',
      DNT: '1',
      Origin: 'https://app.appodeal.com',
      Referer: 'https://app.appodeal.com/analytics/reports?reloadTime=' + Date.now(),
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36',
      'x-requested-with': 'XMLHttpRequest',
      'Trace-Id': Utilities.getUuid()
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  return executeApiRequestWithRetry(config.API_URL, options, projectName, startTime);
}

function executeApiRequestWithRetry(url, options, projectName, startTime, maxRetries = 3) {
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const responseCode = resp.getResponseCode();
      const responseText = resp.getContentText();
      
      if (responseCode === 200) {
        try {
          const parsedResponse = JSON.parse(responseText);
          
          if (parsedResponse.errors && parsedResponse.errors.length > 0) {
            console.error(`${projectName}: GraphQL errors:`, parsedResponse.errors);
            throw new Error(`GraphQL errors: ${JSON.stringify(parsedResponse.errors)}`);
          }
          
          const endTime = Date.now();
          console.log(`${projectName}: API request completed in ${(endTime - startTime) / 1000}s`);
          return parsedResponse;
          
        } catch (parseError) {
          console.error(`${projectName}: JSON parse error:`, parseError);
          throw new Error(`JSON parse error: ${parseError.toString()}`);
        }
      }
      
      if (responseCode >= 400 && responseCode < 500) {
        console.error(`${projectName}: Client error ${responseCode}`);
        
        if (responseCode === 401) {
          throw new Error('Unauthorized: Bearer token may be expired or invalid');
        } else if (responseCode === 403) {
          throw new Error('Forbidden: Insufficient permissions');
        } else if (responseCode === 429) {
          throw new Error('Rate limited: Too many requests');
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
  return `query RichStats($dateFilters: [DateFilterInput!]!, $filters: [FilterInput!]!, $groupBy: [GroupByInput!]!, $measures: [RichMeasureInput!]!, $havingFilters: [HavingFilterInput!], $anonymizationMode: DataAnonymizationMode, $revenuePredictionVersion: String!, $topFilter: TopFilterInput, $funnelFilter: FunnelAttributes, $isMultiMediation: Boolean) {
    analytics(anonymizationMode: $anonymizationMode) {
      richStats(
        funnelFilter: $funnelFilter
        dateFilters: $dateFilters
        filters: $filters
        groupBy: $groupBy
        measures: $measures
        havingFilters: $havingFilters
        revenuePredictionVersion: $revenuePredictionVersion
        topFilter: $topFilter
        isMultiMediation: $isMultiMediation
      ) {
        stats {
          id
          ... on RetentionStatsValue { value cohortSize __typename }
          ... on ForecastStatsItem { value uncertainForecast __typename }
          ... on AppInfo { name platform bundleId __typename }
          ... on LineItemInfo { value appId __typename }
          ... on StatsValue { value __typename }
          ... on SegmentInfo { name description __typename }
          ... on WaterfallConfigurationStats { value appId __typename }
          ... on CountryInfo { code value __typename }
          ... on UaAdSet {
            hid accountId adSetId appId budget budgetPeriod name cpc createdAt lastBidChangedAt
            network recommendedTargetCpa targetCpa targetDayN updatedAt isBeingUpdated isAutomated
            status url type permissions { canUpdateBid canUpdateAutoBid canUpdateBudget canUpdateStatus __typename }
            __typename
          }
          ... on UaCampaign {
            hid accountId campaignId appId budget budgetPeriod campaignName cpc createdAt
            lastBidChangedAt network recommendedTargetCpa targetCpa targetDayN updatedAt
            isBeingUpdated isAutomated autoBidsIgnored status url type permissions {
              canUpdateBid canUpdateAutoBid canUpdateBudget canUpdateStatus __typename
            }
            __typename
          }
          ... on UaCampaignCountry { code bid isBeingUpdated recommendedBid budget country countryId status permissions { canUpdateBid canUpdateAutoBid canUpdateBudget canUpdateStatus __typename } __typename }
          ... on UaCampaignCountrySourceApp { bid iconUrl isBeingUpdated name recommendedBid sourceApp status storeUrl permissions { canUpdateBid canUpdateAutoBid canUpdateBudget canUpdateStatus __typename } __typename }
          ... on SourceAppInfo { name iconUrl storeUrl __typename }
          __typename
        }
        totals {
          day measure value {
            id
            ... on StatsValue { value __typename }
            ... on WaterfallConfigurationStats { value __typename }
            ... on RetentionStatsValue { value cohortSize __typename }
            ... on ForecastStatsItem { value uncertainForecast __typename }
            __typename
          }
          __typename
        }
        anonDict {
          id
          from { id ... on StatsValue { value __typename } __typename }
          to {
            id
            ... on RetentionStatsValue { value cohortSize __typename }
            ... on ForecastStatsItem { value uncertainForecast __typename }
            ... on AppInfo { name __typename }
            ... on StatsValue { value __typename }
            ... on SegmentInfo { name description __typename }
            ... on UaAdSet { name __typename }
            ... on UaCampaign { campaignName __typename }
            __typename
          }
          __typename
        }
        __typename
      }
      __typename
    }
  }`;
}

function ensureBundleIdCacheLoaded() {
  const now = new Date().getTime();
  
  if (BUNDLE_ID_CACHE_LOADED && BUNDLE_ID_CACHE_TIME && (now - BUNDLE_ID_CACHE_TIME) < BUNDLE_ID_CACHE_DURATION) {
    return;
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(BUNDLE_ID_CACHE_SHEET_ID);
    const sheet = spreadsheet.getSheetByName('Bundle ID Cache');
    
    if (!sheet) {
      createBundleIdCacheSheet();
      BUNDLE_ID_CACHE_LOADED = true;
      BUNDLE_ID_CACHE_TIME = now;
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    BUNDLE_ID_CACHE.clear();
    
    for (let i = 1; i < data.length; i++) {
      const [campaignName, campaignId, bundleId, lastUpdated] = data[i];
      if (campaignName && bundleId) {
        BUNDLE_ID_CACHE.set(campaignName, { campaignId, bundleId, lastUpdated });
      }
    }
    
    BUNDLE_ID_CACHE_LOADED = true;
    BUNDLE_ID_CACHE_TIME = now;
  } catch (e) {
    console.error('Error loading Bundle ID Cache:', e);
    BUNDLE_ID_CACHE_LOADED = true;
    BUNDLE_ID_CACHE_TIME = now;
  }
}

function createBundleIdCacheSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(BUNDLE_ID_CACHE_SHEET_ID);
    const sheet = spreadsheet.insertSheet('Bundle ID Cache');
    
    sheet.getRange(1, 1, 1, 4).setValues([['Campaign Name', 'Campaign ID', 'Bundle ID', 'Last Updated']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
    
    sheet.setColumnWidth(1, 300);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 150);
  } catch (e) {
    console.error('Error creating Bundle ID Cache sheet:', e);
  }
}

function saveBundleIdCache(newCache) {
  if (newCache.size === 0) return;
  
  try {
    const spreadsheet = SpreadsheetApp.openById(BUNDLE_ID_CACHE_SHEET_ID);
    const sheet = spreadsheet.getSheetByName('Bundle ID Cache');
    
    if (!sheet) {
      console.error('Bundle ID Cache sheet not found');
      return;
    }
    
    const now = new Date();
    const newEntries = [];
    
    newCache.forEach((value, campaignName) => {
      if (!BUNDLE_ID_CACHE.has(campaignName)) {
        newEntries.push([campaignName, value.campaignId || '', value.bundleId, now]);
        BUNDLE_ID_CACHE.set(campaignName, value);
      }
    });
    
    if (newEntries.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newEntries.length, 4).setValues(newEntries);
    }
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
  if (CURRENT_PROJECT === 'APPLOVIN_NEW') {
    return processApplovinNewApiData(rawData, includeLastWeek);
  }
  
  const startTime = Date.now();
  const stats = rawData.data.analytics.richStats.stats;
  const appData = {};

  const today = new Date();
  const currentWeekStart = formatDateForAPI(getMondayOfWeek(today));
  const lastWeekStart = formatDateForAPI(getMondayOfWeek(new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000)));

  const dayOfWeek = today.getDay();
  const shouldIncludeLastWeek = includeLastWeek !== null ? includeLastWeek : (dayOfWeek >= 2 || dayOfWeek === 0);

  console.log(`Processing ${stats.length} records for ${CURRENT_PROJECT}...`);

  let processedCount = 0;

  if (CURRENT_PROJECT === 'TRICKY') {
    return processTrickyDataOptimized(stats, currentWeekStart, lastWeekStart, shouldIncludeLastWeek);
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

        const sunday = getSundayOfWeek(new Date(date));
        const appKey = app.id;
        
        if (CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'INCENT_TRAFFIC') {
          if (!appData[appKey]) {
            appData[appKey] = {
              appId: app.id,
              appName: app.name,
              platform: app.platform,
              bundleId: app.bundleId,
              weeks: new Map()
            };
          }

          if (!appData[appKey].weeks.has(weekKey)) {
            appData[appKey].weeks.set(weekKey, {
              weekStart: formatDateForAPI(monday),
              weekEnd: formatDateForAPI(sunday),
              networks: new Map()
            });
          }
          
          const networkId = network?.id || 'unknown';
          const networkName = network?.value || 'Unknown Network';
          
          if (!appData[appKey].weeks.get(weekKey).networks.has(networkId)) {
            appData[appKey].weeks.get(weekKey).networks.set(networkId, {
              networkId: networkId,
              networkName: networkName,
              campaigns: []
            });
          }
          
          const networkCampaignData = {
            date: date,
            campaignId: `network_${networkId}_${app.id}_${weekKey}`,
            campaignName: networkName,
            ...metrics,
            status: 'Active',
            type: 'Network',
            geo: 'ALL',
            sourceApp: networkName,
            isAutomated: false
          };
          
          appData[appKey].weeks.get(weekKey).networks.get(networkId).campaigns.push(networkCampaignData);
          processedCount++;
        } else {
          let campaignName = 'Unknown';
          let campaignId = 'Unknown';
          
          if (campaign) {
            if (campaign.campaignName) {
              campaignName = campaign.campaignName;
              campaignId = campaign.campaignId || campaign.id || 'Unknown';
            } else if (campaign.value) {
              campaignName = campaign.value;
              campaignId = campaign.id || 'Unknown';
            }
          }

          const geo = extractGeoFromCampaign(campaignName);
          const sourceApp = extractSourceApp(campaignName);

          const campaignData = {
            date: date,
            campaignId: campaignId,
            campaignName: campaignName,
            ...metrics,
            status: campaign?.status || 'Unknown',
            type: campaign?.type || 'Unknown',
            geo,
            sourceApp: sourceApp,
            isAutomated: campaign?.isAutomated || false
          };

          if (!appData[appKey]) {
            appData[appKey] = {
              appId: app.id,
              appName: app.name,
              platform: app.platform,
              bundleId: app.bundleId,
              weeks: new Map()
            };
          }

          if (!appData[appKey].weeks.has(weekKey)) {
            appData[appKey].weeks.set(weekKey, {
              weekStart: formatDateForAPI(monday),
              weekEnd: formatDateForAPI(sunday),
              campaigns: []
            });
          }
          
          appData[appKey].weeks.get(weekKey).campaigns.push(campaignData);
          processedCount++;
        }

      } catch (error) {
        console.error(`Error processing row ${batchStart + index}:`, error);
      }
    });
  }

  Object.values(appData).forEach(app => {
    const weekMap = app.weeks;
    app.weeks = {};
    weekMap.forEach((value, key) => {
      if (value.networks) {
        const networkMap = value.networks;
        value.networks = {};
        networkMap.forEach((netValue, netKey) => {
          value.networks[netKey] = netValue;
        });
      }
      app.weeks[key] = value;
    });
  });

  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    const networkData = {};
    
    Object.values(appData).forEach(app => {
      Object.values(app.weeks).forEach(week => {
        if (week.networks) {
          Object.values(week.networks).forEach(network => {
            const networkKey = network.networkId;
            
            if (!networkData[networkKey]) {
              networkData[networkKey] = {
                networkId: network.networkId,
                networkName: network.networkName,
                weeks: {}
              };
            }
            
            if (!networkData[networkKey].weeks[week.weekStart]) {
              networkData[networkKey].weeks[week.weekStart] = {
                weekStart: week.weekStart,
                weekEnd: week.weekEnd,
                apps: {}
              };
            }
            
            networkData[networkKey].weeks[week.weekStart].apps[app.appId] = {
              appId: app.appId,
              appName: app.appName,
              platform: app.platform,
              bundleId: app.bundleId,
              campaigns: network.campaigns
            };
          });
        }
      });
    });
    
    const endTime = Date.now();
    console.log(`${CURRENT_PROJECT}: Processing completed in ${(endTime - startTime) / 1000}s - ${processedCount} records processed`);
    return networkData;
  }

  const endTime = Date.now();
  console.log(`${CURRENT_PROJECT}: Processing completed in ${(endTime - startTime) / 1000}s - ${processedCount} records processed`);
  return appData;
}

function processApplovinNewApiData(rawData, includeLastWeek = null) {
  const startTime = Date.now();
  const stats = rawData.data.analytics.richStats.stats;
  const appData = {};

  const today = new Date();
  const currentWeekStart = formatDateForAPI(getMondayOfWeek(today));

  console.log(`Processing ${stats.length} records for APPLOVIN_NEW (including current week)...`);

  let processedCount = 0;

  const BATCH_SIZE = 500;
  for (let batchStart = 0; batchStart < stats.length; batchStart += BATCH_SIZE) {
    const batchEnd = Math.min(batchStart + BATCH_SIZE, stats.length);
    const batch = stats.slice(batchStart, batchEnd);
    
    batch.forEach((row, index) => {
      try {
        const date = row[0].value;
        const campaign = row[1];
        const app = row[2];
        
        const metrics = {
          cpi: parseFloat(row[3].value) || 0,
          installs: parseInt(row[4].value) || 0,
          spend: parseFloat(row[5].value) || 0,
          rrD1: parseFloat(row[6].value) || 0,
          roasD1: parseFloat(row[7].value) || 0,
          roasD3: parseFloat(row[8].value) || 0,
          rrD7: parseFloat(row[9].value) || 0,
          roasD7: parseFloat(row[10].value) || 0,
          roasD14: parseFloat(row[11].value) || 0,
          roasD30: parseFloat(row[12].value) || 0,
          eRoasForecast: parseFloat(row[13].value) || 0,
          eProfitForecast: parseFloat(row[14].value) || 0
        };

        let campaignName = 'Unknown';
        let campaignId = 'Unknown';
        
        if (campaign) {
          if (campaign.campaignName) {
            campaignName = campaign.campaignName;
            campaignId = campaign.campaignId || campaign.id || 'Unknown';
          } else if (campaign.value) {
            campaignName = campaign.value;
            campaignId = campaign.id || 'Unknown';
          }
        }

        const monday = getMondayOfWeek(new Date(date));
        const sunday = getSundayOfWeek(new Date(date));
        const weekKey = formatDateForAPI(monday);
        const geo = extractGeoFromCampaign(campaignName);
        const sourceApp = extractSourceApp(campaignName);
        const appKey = app.id;

        const campaignData = {
          date: date,
          campaignId: campaignId,
          campaignName: campaignName,
          ...metrics,
          status: campaign?.status || 'Unknown',
          type: campaign?.type || 'Unknown',
          geo,
          sourceApp: sourceApp,
          isAutomated: campaign?.isAutomated || false,
          roasRatio3to1: metrics.roasD1 > 0 ? metrics.roasD3 / metrics.roasD1 : 0,
          roasRatio7to3: metrics.roasD3 > 0 ? metrics.roasD7 / metrics.roasD3 : 0,
          roasRatio14to7: metrics.roasD7 > 0 ? metrics.roasD14 / metrics.roasD7 : 0,
          roasRatio30to7: metrics.roasD7 > 0 ? metrics.roasD30 / metrics.roasD7 : 0
        };

        if (!appData[appKey]) {
          appData[appKey] = {
            appId: app.id,
            appName: app.name,
            platform: app.platform,
            bundleId: app.bundleId,
            campaigns: {}
          };
        }

        if (!appData[appKey].campaigns[campaignId]) {
          appData[appKey].campaigns[campaignId] = {
            campaignId: campaignId,
            campaignName: campaignName,
            geo: geo,
            sourceApp: sourceApp,
            weeks: {}
          };
        }

        if (!appData[appKey].campaigns[campaignId].weeks[weekKey]) {
          appData[appKey].campaigns[campaignId].weeks[weekKey] = {
            weekStart: formatDateForAPI(monday),
            weekEnd: formatDateForAPI(sunday),
            days: {}
          };
        }
        
        appData[appKey].campaigns[campaignId].weeks[weekKey].days[date] = campaignData;
        processedCount++;

      } catch (error) {
        console.error(`Error processing APPLOVIN_NEW row ${batchStart + index}:`, error);
      }
    });
  }

  const endTime = Date.now();
  console.log(`APPLOVIN_NEW: Processing completed in ${(endTime - startTime) / 1000}s - ${processedCount} records processed`);
  return appData;
}

function processTrickyDataOptimized(stats, currentWeekStart, lastWeekStart, shouldIncludeLastWeek) {
  const startTime = Date.now();
  
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

        let campaignName = 'Unknown';
        let campaignId = 'Unknown';
        
        if (campaign) {
          if (campaign.campaignName) {
            campaignName = campaign.campaignName;
            campaignId = campaign.campaignId || campaign.id || 'Unknown';
          } else if (campaign.value) {
            campaignName = campaign.value;
            campaignId = campaign.id || 'Unknown';
          }
        }

        const bundleId = getCachedBundleId(campaignName, campaignId);
        if (bundleId && !BUNDLE_ID_CACHE.has(campaignName)) {
          newBundleIds.set(campaignName, { campaignId, bundleId });
        }

        let sourceAppDisplayName;
        if (bundleIdToDisplayName.has(bundleId)) {
          sourceAppDisplayName = bundleIdToDisplayName.get(bundleId);
        } else {
          sourceAppDisplayName = getOptimizedSourceAppDisplayName(bundleId, appsDbCache);
          bundleIdToDisplayName.set(bundleId, sourceAppDisplayName);
        }

        const geo = extractGeoFromCampaign(campaignName);
        const sourceApp = extractSourceApp(campaignName);
        const sunday = getSundayOfWeek(new Date(date));
        const appKey = app.id;

        const campaignData = {
          date: date,
          campaignId: campaignId,
          campaignName: campaignName,
          ...metrics,
          status: campaign?.status || 'Unknown',
          type: campaign?.type || 'Unknown',
          geo,
          sourceApp: sourceApp,
          isAutomated: campaign?.isAutomated || false
        };

        if (!appData[appKey]) {
          appData[appKey] = {
            appId: app.id,
            appName: app.name,
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

  const endTime = Date.now();
  console.log(`TRICKY: Processing completed in ${(endTime - startTime) / 1000}s - ${processedCount} records processed`);
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
    '| FRA |': 'FRA', '| ITA |': 'ITA', '| ESP |': 'ESP', '| RUS |': 'RUS', '| CHN |': 'CHN',
    '| IND |': 'IND', '| TUR |': 'TUR', '| POL |': 'POL', '| NLD |': 'NLD', '| SWE |': 'SWE',
    '| NOR |': 'NOR', '| DNK |': 'DNK', '| FIN |': 'FIN', '| CHE |': 'CHE', '| AUT |': 'AUT',
    '| BEL |': 'BEL', '| PRT |': 'PRT', '| GRC |': 'GRC', '| CZE |': 'CZE', '| HUN |': 'HUN',
    '| ROU |': 'ROU', '| BGR |': 'BGR', '| HRV |': 'HRV', '| SVK |': 'SVK', '| SVN |': 'SVN',
    '| LTU |': 'LTU', '| LVA |': 'LVA', '| EST |': 'EST', '| UKR |': 'UKR', '| BLR |': 'BLR',
    '| ISR |': 'ISR', '| SAU |': 'SAU', '| ARE |': 'ARE', '| QAT |': 'QAT', '| KWT |': 'KWT',
    '| EGY |': 'EGY', '| ZAF |': 'ZAF', '| NGA |': 'NGA', '| KEN |': 'KEN', '| MAR |': 'MAR',
    '| THA |': 'THA', '| VNM |': 'VNM', '| IDN |': 'IDN', '| MYS |': 'MYS', '| SGP |': 'SGP',
    '| PHL |': 'PHL', '| TWN |': 'TWN', '| HKG |': 'HKG', '| ARG |': 'ARG', '| CHL |': 'CHL',
    '| COL |': 'COL', '| PER |': 'PER', '| VEN |': 'VEN', '| URY |': 'URY', '| ECU |': 'ECU',
    '| BOL |': 'BOL', '| PRY |': 'PRY', '| CRI |': 'CRI', '| GTM |': 'GTM', '| DOM |': 'DOM',
    '| PAN |': 'PAN', '| NZL |': 'NZL'
  };

    for (const [pattern, geo] of Object.entries(geoMap)) {
      if (campaignName.includes(pattern)) {
        return geo;
      }
    }
    return 'OTHER';
  }
  
  if (CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    return 'ALL';
  }
  
  if (CURRENT_PROJECT === 'GOOGLE_ADS') {
    const geoPatterns = [
      { pattern: 'LatAm', geo: 'LatAm' }, { pattern: 'UK,GE', geo: 'UK,GE' }, { pattern: 'BR (PT)', geo: 'BR' },
      { pattern: 'US ', geo: 'US' }, { pattern: ' US ', geo: 'US' }, { pattern: 'WW ', geo: 'WW' },
      { pattern: ' WW ', geo: 'WW' }, { pattern: 'UK', geo: 'UK' }, { pattern: 'GE', geo: 'GE' }, { pattern: 'BR', geo: 'BR' }
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
    
    if (CURRENT_PROJECT === 'REGULAR' || CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN' || CURRENT_PROJECT === 'APPLOVIN_NEW' || CURRENT_PROJECT === 'MINTEGRAL' || CURRENT_PROJECT === 'INCENT') {
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