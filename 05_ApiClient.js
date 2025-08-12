// Кеши для Bundle ID и Apps Database
var BUNDLE_ID_CACHE = new Map();
var BUNDLE_ID_CACHE_LOADED = false;
var BUNDLE_ID_CACHE_TIME = null;
var BUNDLE_ID_CACHE_DURATION = 21600000;
var APPS_DB_CACHE = null;
var APPS_DB_CACHE_TIME = null;
var BUNDLE_ID_CACHE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1Z5pJgtg--9EACJL8PVZgJsmeUemv6PKhSsyx9ArChrM/edit?gid=754371211#gid=754371211';
var BUNDLE_ID_CACHE_SHEET_ID = '1Z5pJgtg--9EACJL8PVZgJsmeUemv6PKhSsyx9ArChrM';

// GEO паттерны для всех проектов
const GEO_PATTERNS = {
  TRICKY: {
    type: 'pipe',
    map: {
      '| USA |': 'USA', '| MEX |': 'MEX', '| AUS |': 'AUS', '| DEU |': 'DEU', '| JPN |': 'JPN',
      '| KOR |': 'KOR', '| BRA |': 'BRA', '| CAN |': 'CAN', '| GBR |': 'GBR', '| FRA |': 'FRA',
      '| ITA |': 'ITA', '| ESP |': 'ESP', '| RUS |': 'RUS', '| CHN |': 'CHN', '| IND |': 'IND',
      '| TUR |': 'TUR', '| POL |': 'POL', '| NLD |': 'NLD', '| SWE |': 'SWE', '| NOR |': 'NOR',
      '| DNK |': 'DNK', '| FIN |': 'FIN', '| CHE |': 'CHE', '| AUT |': 'AUT', '| BEL |': 'BEL',
      '| PRT |': 'PRT', '| GRC |': 'GRC', '| CZE |': 'CZE', '| HUN |': 'HUN', '| ROU |': 'ROU',
      '| BGR |': 'BGR', '| HRV |': 'HRV', '| SVK |': 'SVK', '| SVN |': 'SVN', '| LTU |': 'LTU',
      '| LVA |': 'LVA', '| EST |': 'EST', '| UKR |': 'UKR', '| BLR |': 'BLR', '| ISR |': 'ISR',
      '| SAU |': 'SAU', '| ARE |': 'ARE', '| QAT |': 'QAT', '| KWT |': 'KWT', '| EGY |': 'EGY',
      '| ZAF |': 'ZAF', '| NGA |': 'NGA', '| KEN |': 'KEN', '| MAR |': 'MAR', '| THA |': 'THA',
      '| VNM |': 'VNM', '| IDN |': 'IDN', '| MYS |': 'MYS', '| SGP |': 'SGP', '| PHL |': 'PHL',
      '| TWN |': 'TWN', '| HKG |': 'HKG', '| ARG |': 'ARG', '| CHL |': 'CHL', '| COL |': 'COL',
      '| PER |': 'PER', '| VEN |': 'VEN', '| URY |': 'URY', '| ECU |': 'ECU', '| BOL |': 'BOL',
      '| PRY |': 'PRY', '| CRI |': 'CRI', '| GTM |': 'GTM', '| DOM |': 'DOM', '| PAN |': 'PAN', '| NZL |': 'NZL'
    }
  },
  GOOGLE_ADS: {
    type: 'keywords',
    patterns: [
      { pattern: 'LatAm', geo: 'LatAm' }, { pattern: 'UK,GE', geo: 'UK,GE' }, 
      { pattern: 'BR (PT)', geo: 'BR' }, { pattern: 'US ', geo: 'US' },
      { pattern: ' US ', geo: 'US' }, { pattern: 'WW ', geo: 'WW' },
      { pattern: ' WW ', geo: 'WW' }, { pattern: 'UK', geo: 'UK' },
      { pattern: 'GE', geo: 'GE' }, { pattern: 'BR', geo: 'BR' }
    ]
  },
  DEFAULT: {
    type: 'standard',
    patterns: ['WW_ru', 'WW_es', 'WW_de', 'WW_pt', 'Asia T1', 'T2-ES', 'T1-EN', 'LatAm', 
               'TopGeo', 'Europe', 'US', 'RU', 'UK', 'GE', 'FR', 'PT', 'ES', 'DE', 'T1', 'WW']
  }
};

// Основные API функции
function fetchCampaignData(dateRange, projectName = null) {
  const effectiveProject = projectName || CURRENT_PROJECT;
  const config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
  const apiConfig = projectName ? getProjectApiConfig(projectName) : getCurrentApiConfig();
  
  if (!config.BEARER_TOKEN) throw new Error(`${effectiveProject} project is not configured: missing BEARER_TOKEN`);
  if (!apiConfig.FILTERS.USER?.length) throw new Error(`${effectiveProject} project is not configured: missing USER filters`);
  
  return buildAndExecuteApiRequest(config, apiConfig, dateRange, effectiveProject, Date.now());
}

function fetchProjectCampaignData(projectName, dateRange) {
  return fetchCampaignData(dateRange, projectName);
}

function buildAndExecuteApiRequest(config, apiConfig, dateRange, projectName, startTime) {
  const filters = buildFilters(apiConfig);
  const dateDimension = getDateDimension(projectName);
  
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
    headers: buildHeaders(config.BEARER_TOKEN),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  return executeApiRequestWithRetry(config.API_URL, options, projectName, startTime);
}

// Вспомогательные функции для API
function buildFilters(apiConfig) {
  const filters = [
    { dimension: "USER", values: apiConfig.FILTERS.USER, include: true },
    { dimension: "ATTRIBUTION_PARTNER", values: apiConfig.FILTERS.ATTRIBUTION_PARTNER, include: true }
  ];
  
  if (apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID?.length > 0) {
    filters.push({ dimension: "ATTRIBUTION_NETWORK_HID", values: apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID, include: true });
  }
  
  if (apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH) {
    const search = apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH;
    const isExclude = search.startsWith('!');
    filters.push({
      dimension: "ATTRIBUTION_CAMPAIGN_HID",
      values: [],
      include: !isExclude,
      searchByString: isExclude ? search.substring(1) : search
    });
  }
  
  return filters;
}

function buildHeaders(bearerToken) {
  return {
    Accept: 'application/json, text/plain, */*',
    'Accept-Language': 'en-US,en;q=0.9',
    Authorization: `Bearer ${bearerToken}`,
    Connection: 'keep-alive',
    DNT: '1',
    Origin: 'https://app.appodeal.com',
    Referer: 'https://app.appodeal.com/analytics/reports?reloadTime=' + Date.now(),
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
    'x-requested-with': 'XMLHttpRequest',
    'Trace-Id': Utilities.getUuid()
  };
}

function getDateDimension(projectName) {
  const dateDimensionProjects = ['GOOGLE_ADS', 'APPLOVIN', 'INCENT', 'OVERALL'];
  return dateDimensionProjects.includes(projectName) ? 'DATE' : 'INSTALL_DATE';
}

function executeApiRequestWithRetry(url, options, projectName, startTime, maxRetries = 3) {
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const responseCode = resp.getResponseCode();
      const responseText = resp.getContentText();
      
      if (responseCode === 200) {
        const parsedResponse = JSON.parse(responseText);
        if (parsedResponse.errors?.length > 0) {
          throw new Error(`GraphQL errors: ${JSON.stringify(parsedResponse.errors)}`);
        }
        console.log(`${projectName}: API request completed in ${(Date.now() - startTime) / 1000}s`);
        return parsedResponse;
      }
      
      if (responseCode >= 400 && responseCode < 500) {
        const errorMessages = {
          401: 'Unauthorized: Bearer token may be expired or invalid',
          403: 'Forbidden: Insufficient permissions',
          429: 'Rate limited: Too many requests'
        };
        throw new Error(errorMessages[responseCode] || `Client error ${responseCode}: ${responseText.substring(0, 200)}`);
      }
      
      lastError = new Error(`Server error ${responseCode}`);
      if (attempt < maxRetries) {
        Utilities.sleep(Math.min(1000 * Math.pow(2, attempt - 1), 10000));
      }
    } catch (e) {
      lastError = e;
      if (attempt < maxRetries && e.toString().includes('timed out')) {
        Utilities.sleep(2000 * attempt);
      }
    }
  }
  
  throw lastError || new Error('API request failed after all retries');
}

// Обработка данных
function processApiData(rawData, includeLastWeek = null) {
  const stats = rawData.data.analytics.richStats.stats;
  const today = new Date();
  const currentWeekStart = formatDateForAPI(getMondayOfWeek(today));
  const lastWeekStart = formatDateForAPI(getMondayOfWeek(new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000)));
  const shouldIncludeLastWeek = includeLastWeek !== null ? includeLastWeek : (today.getDay() >= 2 || today.getDay() === 0);

  console.log(`Processing ${stats.length} records for ${CURRENT_PROJECT}...`);

  // Специальная обработка для TRICKY
  if (CURRENT_PROJECT === 'TRICKY') {
    return processTrickyData(stats, currentWeekStart, lastWeekStart, shouldIncludeLastWeek);
  }
  
  // Общая обработка для остальных проектов
  return processStandardData(stats, currentWeekStart, lastWeekStart, shouldIncludeLastWeek);
}

function processStandardData(stats, currentWeekStart, lastWeekStart, shouldIncludeLastWeek) {
  const appData = {};
  let processedCount = 0;
  const BATCH_SIZE = 500;
  
  for (let batchStart = 0; batchStart < stats.length; batchStart += BATCH_SIZE) {
    const batch = stats.slice(batchStart, Math.min(batchStart + BATCH_SIZE, stats.length));
    
    batch.forEach(row => {
      const date = row[0].value;
      const weekKey = formatDateForAPI(getMondayOfWeek(new Date(date)));
      
      if (weekKey >= currentWeekStart || (!shouldIncludeLastWeek && weekKey >= lastWeekStart)) return;
      
      const rowData = parseRowData(row);
      const appKey = rowData.app.id;
      
      if (!appData[appKey]) {
        appData[appKey] = createAppStructure(rowData.app);
      }
      
      if (!appData[appKey].weeks[weekKey]) {
        appData[appKey].weeks[weekKey] = createWeekStructure(date);
      }
      
      addDataToWeek(appData[appKey].weeks[weekKey], rowData, date);
      processedCount++;
    });
  }
  
  // Конвертация для INCENT_TRAFFIC
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    return convertToNetworkStructure(appData);
  }
  
  console.log(`${CURRENT_PROJECT}: Processed ${processedCount} records`);
  return appData;
}

function processTrickyData(stats, currentWeekStart, lastWeekStart, shouldIncludeLastWeek) {
  ensureBundleIdCacheLoaded();
  const appsDbCache = getOptimizedAppsDbForTricky();
  const appData = {};
  const newBundleIds = new Map();
  let processedCount = 0;
  
  stats.forEach(row => {
    const date = row[0].value;
    const weekKey = formatDateForAPI(getMondayOfWeek(new Date(date)));
    
    if (weekKey >= currentWeekStart || (!shouldIncludeLastWeek && weekKey >= lastWeekStart)) return;
    
    const rowData = parseRowData(row);
    const bundleId = getCachedBundleId(rowData.campaignName, rowData.campaignId) || 'unknown';
    
    if (bundleId && !BUNDLE_ID_CACHE.has(rowData.campaignName)) {
      newBundleIds.set(rowData.campaignName, { campaignId: rowData.campaignId, bundleId });
    }
    
    const sourceAppDisplayName = getOptimizedSourceAppDisplayName(bundleId, appsDbCache);
    const appKey = rowData.app.id;
    
    if (!appData[appKey]) {
      appData[appKey] = createAppStructure(rowData.app);
    }
    
    if (!appData[appKey].weeks[weekKey]) {
      appData[appKey].weeks[weekKey] = createWeekStructure(date);
      appData[appKey].weeks[weekKey].sourceApps = {};
    }
    
    if (!appData[appKey].weeks[weekKey].sourceApps[bundleId]) {
      appData[appKey].weeks[weekKey].sourceApps[bundleId] = {
        sourceAppId: bundleId,
        sourceAppName: sourceAppDisplayName,
        campaigns: []
      };
    }
    
    appData[appKey].weeks[weekKey].sourceApps[bundleId].campaigns.push(createCampaignData(rowData, date));
    processedCount++;
  });
  
  if (newBundleIds.size > 0) saveBundleIdCache(newBundleIds);
  
  console.log(`TRICKY: Processed ${processedCount} records`);
  return appData;
}

// Вспомогательные функции обработки
function parseRowData(row) {
  const isOverallOrIncent = ['OVERALL', 'INCENT_TRAFFIC'].includes(CURRENT_PROJECT);
  const campaign = isOverallOrIncent ? null : row[1];
  const network = isOverallOrIncent ? row[1] : null;
  const app = row[2];
  const metricsStartIndex = 3;
  
  const campaignName = campaign ? (campaign.campaignName || campaign.value || 'Unknown') : 'Unknown';
  const campaignId = campaign ? (campaign.campaignId || campaign.id || 'Unknown') : 'Unknown';
  
  return {
    campaign, network, app, campaignName, campaignId,
    geo: extractGeoFromCampaign(campaignName),
    sourceApp: extractSourceApp(campaignName),
    metrics: extractMetrics(row, metricsStartIndex),
    status: campaign?.status || 'Unknown',
    type: campaign?.type || 'Unknown',
    isAutomated: campaign?.isAutomated || false
  };
}

function extractMetrics(row, startIndex) {
  return {
    cpi: parseFloat(row[startIndex].value) || 0,
    installs: parseInt(row[startIndex + 1].value) || 0,
    ipm: parseFloat(row[startIndex + 2].value) || 0,
    spend: parseFloat(row[startIndex + 3].value) || 0,
    rrD1: parseFloat(row[startIndex + 4].value) || 0,
    roasD1: parseFloat(row[startIndex + 5].value) || 0,
    roasD3: parseFloat(row[startIndex + 6].value) || 0,
    rrD7: parseFloat(row[startIndex + 7].value) || 0,
    roasD7: parseFloat(row[startIndex + 8].value) || 0,
    roasD30: parseFloat(row[startIndex + 9].value) || 0,
    eArpuForecast: parseFloat(row[startIndex + 10].value) || 0,
    eRoasForecast: parseFloat(row[startIndex + 11].value) || 0,
    eProfitForecast: parseFloat(row[startIndex + 12].value) || 0,
    eRoasForecastD730: parseFloat(row[startIndex + 13].value) || 0
  };
}

function createAppStructure(app) {
  return {
    appId: app.id,
    appName: app.name,
    platform: app.platform,
    bundleId: app.bundleId,
    weeks: {}
  };
}

function createWeekStructure(date) {
  const monday = getMondayOfWeek(new Date(date));
  const sunday = getSundayOfWeek(new Date(date));
  return {
    weekStart: formatDateForAPI(monday),
    weekEnd: formatDateForAPI(sunday),
    campaigns: []
  };
}

function createCampaignData(rowData, date) {
  return {
    date: date,
    campaignId: rowData.campaignId,
    campaignName: rowData.campaignName,
    ...rowData.metrics,
    status: rowData.status,
    type: rowData.type,
    geo: rowData.geo,
    sourceApp: rowData.sourceApp,
    isAutomated: rowData.isAutomated
  };
}

function addDataToWeek(week, rowData, date) {
  if (['OVERALL', 'INCENT_TRAFFIC'].includes(CURRENT_PROJECT)) {
    const networkId = rowData.network?.id || 'unknown';
    const networkName = rowData.network?.value || 'Unknown Network';
    
    if (!week.networks) week.networks = {};
    if (!week.networks[networkId]) {
      week.networks[networkId] = {
        networkId: networkId,
        networkName: networkName,
        campaigns: []
      };
    }
    
    week.networks[networkId].campaigns.push({
      date: date,
      campaignId: `network_${networkId}_${rowData.app.id}_${week.weekStart}`,
      campaignName: networkName,
      ...rowData.metrics,
      status: 'Active',
      type: 'Network',
      geo: 'ALL',
      sourceApp: networkName,
      isAutomated: false
    });
  } else {
    week.campaigns.push(createCampaignData(rowData, date));
  }
}

function convertToNetworkStructure(appData) {
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
  
  return networkData;
}

// GEO extraction
function extractGeoFromCampaign(campaignName) {
  if (!campaignName) return 'OTHER';
  
  const project = CURRENT_PROJECT === 'REGULAR' ? 'TRICKY' : CURRENT_PROJECT;
  const config = GEO_PATTERNS[project] || GEO_PATTERNS.DEFAULT;
  
  if (project === 'OVERALL' || project === 'INCENT_TRAFFIC') return 'ALL';
  
  if (config.type === 'pipe') {
    for (const [pattern, geo] of Object.entries(config.map)) {
      if (campaignName.includes(pattern)) return geo;
    }
  } else if (config.type === 'keywords') {
    for (const {pattern, geo} of config.patterns) {
      if (campaignName.includes(pattern)) return geo;
    }
  } else {
    const upperName = campaignName.toUpperCase();
    for (const pattern of config.patterns) {
      const upperPattern = pattern.toUpperCase();
      if (upperName.includes('_' + upperPattern + '_') || upperName.includes('-' + upperPattern + '-') ||
          upperName.includes('_' + upperPattern) || upperName.includes('-' + upperPattern) ||
          upperName.includes(upperPattern + '_') || upperName.includes(upperPattern + '-') ||
          upperName === upperPattern) {
        return pattern;
      }
    }
  }
  
  return 'OTHER';
}

function extractSourceApp(campaignName) {
  if (['OVERALL', 'INCENT_TRAFFIC'].includes(CURRENT_PROJECT)) return campaignName;
  if (campaignName.startsWith('APD_')) return campaignName;
  if (['REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT'].includes(CURRENT_PROJECT)) return campaignName;
  
  // TRICKY logic
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
}

// Bundle ID функции
function ensureBundleIdCacheLoaded() {
  const now = Date.now();
  if (BUNDLE_ID_CACHE_LOADED && BUNDLE_ID_CACHE_TIME && (now - BUNDLE_ID_CACHE_TIME) < BUNDLE_ID_CACHE_DURATION) return;
  
  try {
    const sheet = SpreadsheetApp.openById(BUNDLE_ID_CACHE_SHEET_ID).getSheetByName('Bundle ID Cache');
    if (!sheet) {
      createBundleIdCacheSheet();
    } else {
      const data = sheet.getDataRange().getValues();
      BUNDLE_ID_CACHE.clear();
      for (let i = 1; i < data.length; i++) {
        const [campaignName, campaignId, bundleId] = data[i];
        if (campaignName && bundleId) {
          BUNDLE_ID_CACHE.set(campaignName, { campaignId, bundleId });
        }
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
    const sheet = SpreadsheetApp.openById(BUNDLE_ID_CACHE_SHEET_ID).insertSheet('Bundle ID Cache');
    sheet.getRange(1, 1, 1, 4).setValues([['Campaign Name', 'Campaign ID', 'Bundle ID', 'Last Updated']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
    sheet.setColumnWidths(1, 4, [300, 150, 200, 150]);
  } catch (e) {
    console.error('Error creating Bundle ID Cache sheet:', e);
  }
}

function saveBundleIdCache(newCache) {
  if (newCache.size === 0) return;
  
  try {
    const sheet = SpreadsheetApp.openById(BUNDLE_ID_CACHE_SHEET_ID).getSheetByName('Bundle ID Cache');
    if (!sheet) return;
    
    const newEntries = [];
    const now = new Date();
    
    newCache.forEach((value, campaignName) => {
      if (!BUNDLE_ID_CACHE.has(campaignName)) {
        newEntries.push([campaignName, value.campaignId || '', value.bundleId, now]);
        BUNDLE_ID_CACHE.set(campaignName, value);
      }
    });
    
    if (newEntries.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newEntries.length, 4).setValues(newEntries);
    }
  } catch (e) {
    console.error('Error saving Bundle ID Cache:', e);
  }
}

function getCachedBundleId(campaignName, campaignId = '') {
  if (BUNDLE_ID_CACHE.has(campaignName)) return BUNDLE_ID_CACHE.get(campaignName).bundleId;
  return extractBundleIdFromCampaign(campaignName);
}

function getOptimizedAppsDbForTricky() {
  const now = Date.now();
  if (APPS_DB_CACHE && APPS_DB_CACHE_TIME && (now - APPS_DB_CACHE_TIME) < BUNDLE_ID_CACHE_DURATION) {
    return APPS_DB_CACHE;
  }
  
  try {
    const appsDb = new AppsDatabase('TRICKY');
    appsDb.ensureCacheUpToDate();
    APPS_DB_CACHE = appsDb.loadFromCache();
    APPS_DB_CACHE_TIME = now;
    return APPS_DB_CACHE;
  } catch (e) {
    console.error('Error loading Apps Database:', e);
    return {};
  }
}

function getOptimizedSourceAppDisplayName(bundleId, appsDbCache) {
  if (!bundleId || CURRENT_PROJECT !== 'TRICKY' || !appsDbCache) return bundleId || 'Unknown';
  
  const appInfo = appsDbCache[bundleId];
  if (appInfo && appInfo.publisher !== bundleId) {
    const publisher = appInfo.publisher || '';
    const appName = appInfo.appName || '';
    
    if (publisher && appName && publisher !== appName) return `${publisher} ${appName}`;
    if (publisher) return publisher;
    if (appName) return appName;
  }
  
  return bundleId;
}

// Legacy функции для совместимости
function processProjectApiData(projectName, rawData, includeLastWeek = null) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    return processApiData(rawData, includeLastWeek);
  } finally {
    setCurrentProject(originalProject);
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

// GraphQL запрос (не трогаем по просьбе)
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