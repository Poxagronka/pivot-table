var BUNDLE_ID_CACHE = {};
var APPS_DB_CACHE = null;
var APPS_DB_CACHE_TIME = null;
var TRICKY_OPTIMIZED_CACHE = null;

function fetchCampaignData(dateRange) {
  const config = getCurrentConfig();
  const apiConfig = getCurrentApiConfig();
  const filters = buildFilters(apiConfig);
  const dateDimension = ['GOOGLE_ADS','APPLOVIN','INCENT','OVERALL'].includes(CURRENT_PROJECT) ? 'DATE' : 'INSTALL_DATE';
  
  const payload = {
    operationName: apiConfig.OPERATION_NAME,
    variables: {
      dateFilters: [{dimension: dateDimension, from: dateRange.from, to: dateRange.to, include: true}],
      filters: filters,
      groupBy: apiConfig.GROUP_BY,
      measures: apiConfig.MEASURES,
      havingFilters: [{measure: {id: "spend", day: null}, operator: "MORE", value: 0}],
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
    headers: buildHeaders(config),
    payload: JSON.stringify(payload)
  };

  const resp = UrlFetchApp.fetch(config.API_URL, options);
  if (resp.getResponseCode() !== 200) throw new Error('API request failed: ' + resp.getContentText());
  return JSON.parse(resp.getContentText());
}

function fetchProjectCampaignData(projectName, dateRange) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    return fetchCampaignData(dateRange);
  } finally {
    setCurrentProject(originalProject);
  }
}

function buildFilters(apiConfig) {
  const filters = [
    {dimension: "USER", values: apiConfig.FILTERS.USER, include: true},
    {dimension: "ATTRIBUTION_PARTNER", values: apiConfig.FILTERS.ATTRIBUTION_PARTNER, include: true}
  ];
  
  if (apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID !== null && apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID?.length > 0) {
    filters.push({dimension: "ATTRIBUTION_NETWORK_HID", values: apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID, include: true});
  }
  
  if (apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH) {
    const searchPattern = apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH;
    filters.push({
      dimension: "ATTRIBUTION_CAMPAIGN_HID", 
      values: [], 
      include: !searchPattern.startsWith('!'),
      searchByString: searchPattern.startsWith('!') ? searchPattern.substring(1) : searchPattern
    });
  }
  
  return filters;
}

function buildHeaders(config) {
  return {
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
  };
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

function processApiData(rawData, includeLastWeek = null) {
  if (CURRENT_PROJECT === 'OVERALL') {
    return processOverallApiData(rawData, includeLastWeek);
  }
  if (CURRENT_PROJECT === 'TRICKY') {
    return processApiDataTrickyOptimized(rawData, includeLastWeek);
  }
  return processApiDataStandard(rawData, includeLastWeek);
}

function processOverallApiData(rawData, includeLastWeek = null) {
  const stats = rawData.data.analytics.richStats.stats;
  const appData = {};
  const today = new Date();
  const currentWeekStart = formatDateForAPI(getMondayOfWeek(today));
  const lastWeekStart = formatDateForAPI(getMondayOfWeek(new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000)));
  const dayOfWeek = today.getDay();
  const shouldIncludeLastWeek = includeLastWeek !== null ? includeLastWeek : (dayOfWeek >= 2 || dayOfWeek === 0);

  stats.forEach((row, index) => {
    try {
      const date = row[0].value;
      const monday = getMondayOfWeek(new Date(date));
      const weekKey = formatDateForAPI(monday);

      if (weekKey >= currentWeekStart) return;
      if (!shouldIncludeLastWeek && weekKey >= lastWeekStart) return;

      const network = row[1];
      const app = row[2];
      const metricsStartIndex = 3;
      
      if (row.length < metricsStartIndex + 8) return;
      
      const metrics = extractOverallMetrics(row, metricsStartIndex);
      const sunday = getSundayOfWeek(new Date(date));
      const appKey = app.id;

      const networkId = network.id || network.value || 'unknown';
      const networkName = getNetworkDisplayName(networkId);

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
          networks: {}
        };
      }

      if (!appData[appKey].weeks[weekKey].networks[networkId]) {
        appData[appKey].weeks[weekKey].networks[networkId] = {
          networkId: networkId,
          networkName: networkName,
          ...metrics
        };
      } else {
        const existing = appData[appKey].weeks[weekKey].networks[networkId];
        existing.cpi = ((existing.cpi * existing.installs) + (metrics.cpi * metrics.installs)) / (existing.installs + metrics.installs);
        existing.installs += metrics.installs;
        existing.spend += metrics.spend;
        existing.rrD1 = ((existing.rrD1 * existing.installs) + (metrics.rrD1 * metrics.installs)) / (existing.installs + metrics.installs);
        existing.roas = ((existing.roas * existing.spend) + (metrics.roas * metrics.spend)) / (existing.spend + metrics.spend);
        existing.rrD7 = ((existing.rrD7 * existing.installs) + (metrics.rrD7 * metrics.installs)) / (existing.installs + metrics.installs);
        existing.eRoasForecast = ((existing.eRoasForecast * existing.spend) + (metrics.eRoasForecast * metrics.spend)) / (existing.spend + metrics.spend);
        existing.eProfitForecast += metrics.eProfitForecast;
      }

    } catch (error) {
      console.error('Error processing OVERALL row:', error);
    }
  });

  return appData;
}

function getNetworkDisplayName(networkId) {
  const networkMap = {
    'appgrowth_int': 'AppGrowth',
    'inmobidsp_int': 'InMobidsp', 
    'thespotlight_int': 'TheSpotlight',
    'aura_int': 'Aura',
    'moloco_int': 'Moloco',
    'mintegral_int': 'Mintegral',
    'googleadwords_int': 'Google Ads',
    'applovin_int': 'AppLovin',
    'ayetstudios_int': 'Ayet Studios',
    'engagerewards_int': 'Engage Rewards',
    'adjoe_int': 'Adjoe',
    'mistplay_int': 'Mistplay',
    '234187180623265792': 'AppGrowth',
    '445856363109679104': 'Moloco',
    '378302368699121664': 'Google Ads',
    '261208778387488768': 'AppLovin',
    '756604737398243328': 'Mintegral',
    '1580763469207044096': 'Incent Network 1',
    '932245122865692672': 'Incent Network 2',
    '6958061424287416320': 'Incent Network 3',
    '6070852297695428608': 'Incent Network 4',
    '5354779956943519744': 'Incent Network 5'
  };
  
  return networkMap[networkId] || networkId;
}

function extractOverallMetrics(row, startIndex) {
  return {
    cpi: parseFloat(row[startIndex].value) || 0,
    installs: parseInt(row[startIndex + 1].value) || 0,
    spend: parseFloat(row[startIndex + 2].value) || 0,
    rrD1: parseFloat(row[startIndex + 3].value) || 0,
    roas: parseFloat(row[startIndex + 4].value) || 0,
    rrD7: parseFloat(row[startIndex + 5].value) || 0,
    eRoasForecast: parseFloat(row[startIndex + 6].value) || 0,
    eProfitForecast: parseFloat(row[startIndex + 7].value) || 0
  };
}

function initTrickyOptimizedCache() {
  if (CURRENT_PROJECT !== 'TRICKY' || TRICKY_OPTIMIZED_CACHE) return TRICKY_OPTIMIZED_CACHE;
  
  const appsDb = new AppsDatabase('TRICKY');
  appsDb.ensureCacheUpToDate();
  const appsDbCache = appsDb.loadFromCache();
  
  TRICKY_OPTIMIZED_CACHE = {
    appsDbCache: appsDbCache,
    bundleIdCache: {},
    processed: 0,
    cacheHits: 0
  };
  
  return TRICKY_OPTIMIZED_CACHE;
}

function batchExtractBundleIds(campaignNames) {
  if (CURRENT_PROJECT !== 'TRICKY') return {};
  
  const cache = initTrickyOptimizedCache();
  const results = {};
  
  campaignNames.forEach(campaignName => {
    if (cache.bundleIdCache[campaignName] !== undefined) {
      results[campaignName] = cache.bundleIdCache[campaignName];
      cache.cacheHits++;
    } else {
      const bundleId = extractBundleIdFromCampaign(campaignName);
      cache.bundleIdCache[campaignName] = bundleId;
      results[campaignName] = bundleId;
    }
  });
  
  return results;
}

function getOptimizedSourceAppDisplayName(bundleId, appsDbCache) {
  if (!bundleId || CURRENT_PROJECT !== 'TRICKY') return bundleId || 'Unknown';
  
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

function processApiDataTrickyOptimized(rawData, includeLastWeek = null) {
  const stats = rawData.data.analytics.richStats.stats;
  const appData = {};
  const today = new Date();
  const currentWeekStart = formatDateForAPI(getMondayOfWeek(today));
  const lastWeekStart = formatDateForAPI(getMondayOfWeek(new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000)));
  const dayOfWeek = today.getDay();
  const shouldIncludeLastWeek = includeLastWeek !== null ? includeLastWeek : (dayOfWeek >= 2 || dayOfWeek === 0);
  const cache = initTrickyOptimizedCache();
  
  const campaignNames = [];
  const validRows = [];
  
  stats.forEach((row, index) => {
    try {
      const date = row[0].value;
      const monday = getMondayOfWeek(new Date(date));
      const weekKey = formatDateForAPI(monday);

      if (weekKey >= currentWeekStart) return;
      if (!shouldIncludeLastWeek && weekKey >= lastWeekStart) return;

      const campaign = row[1];
      const app = row[2];
      const metricsStartIndex = 3;
      
      if (row.length < metricsStartIndex + 12) return;
      
      const spendValue = parseFloat(row[metricsStartIndex + 3].value) || 0;
      if (spendValue <= 0) return;
      
      let campaignName = 'Unknown';
      if (campaign?.campaignName) {
        campaignName = campaign.campaignName;
      } else if (campaign?.value) {
        campaignName = campaign.value;
      }
      
      campaignNames.push(campaignName);
      validRows.push({ row, index, date, weekKey, campaign, app, campaignName, spendValue });
      
    } catch (error) {}
  });

  const bundleIdMap = batchExtractBundleIds(campaignNames);
  const groupedData = {};
  
  validRows.forEach(({ row, date, weekKey, campaign, app, campaignName, spendValue }) => {
    try {
      const metricsStartIndex = 3;
      const metrics = extractMetrics(row, metricsStartIndex);
      const sunday = getSundayOfWeek(new Date(date));
      const appKey = app.id;
      const bundleId = bundleIdMap[campaignName];
      const geo = extractGeoFromCampaign(campaignName);
      const sourceApp = extractSourceApp(campaignName);
      const campaignId = campaign?.campaignId || campaign?.id || 'Unknown';

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
        extractedBundleId: bundleId
      };
      
      if (!groupedData[appKey]) {
        groupedData[appKey] = {
          appId: app.id,
          appName: app.name,
          platform: app.platform,
          bundleId: app.bundleId,
          weeks: {}
        };
      }
      
      if (!groupedData[appKey].weeks[weekKey]) {
        groupedData[appKey].weeks[weekKey] = {
          weekStart: formatDateForAPI(getMondayOfWeek(new Date(date))),
          weekEnd: formatDateForAPI(sunday),
          campaigns: []
        };
      }
      
      groupedData[appKey].weeks[weekKey].campaigns.push(campaignData);
      cache.processed++;
      
    } catch (error) {}
  });

  Object.keys(groupedData).forEach(appKey => {
    const appInfo = groupedData[appKey];
    
    if (!appData[appKey]) {
      appData[appKey] = {
        appId: appInfo.appId,
        appName: appInfo.appName,
        platform: appInfo.platform,
        bundleId: appInfo.bundleId,
        weeks: {}
      };
    }
    
    Object.keys(appInfo.weeks).forEach(weekKey => {
      const weekInfo = appInfo.weeks[weekKey];
      const bundleGroups = {};
      
      weekInfo.campaigns.forEach(campaign => {
        const bundleId = campaign.extractedBundleId || 'unknown';
        if (!bundleGroups[bundleId]) bundleGroups[bundleId] = [];
        bundleGroups[bundleId].push(campaign);
      });
      
      const sourceApps = {};
      const sortedBundleIds = Object.keys(bundleGroups).sort();
      
      sortedBundleIds.forEach(bundleId => {
        const campaigns = bundleGroups[bundleId];
        campaigns.sort((a, b) => b.spend - a.spend);
        
        const sourceAppDisplayName = getOptimizedSourceAppDisplayName(bundleId, cache.appsDbCache);
        
        sourceApps[bundleId] = {
          sourceAppId: bundleId,
          sourceAppName: sourceAppDisplayName,
          sourceAppIconUrl: '',
          sourceAppStoreUrl: '',
          campaigns: campaigns
        };
      });
      
      appData[appKey].weeks[weekKey] = {
        weekStart: weekInfo.weekStart,
        weekEnd: weekInfo.weekEnd,
        sourceApps: sourceApps,
        campaigns: []
      };
    });
  });

  return appData;
}

function processApiDataStandard(rawData, includeLastWeek = null) {
  const stats = rawData.data.analytics.richStats.stats;
  const appData = {};
  const today = new Date();
  const currentWeekStart = formatDateForAPI(getMondayOfWeek(today));
  const lastWeekStart = formatDateForAPI(getMondayOfWeek(new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000)));
  const dayOfWeek = today.getDay();
  const shouldIncludeLastWeek = includeLastWeek !== null ? includeLastWeek : (dayOfWeek >= 2 || dayOfWeek === 0);

  stats.forEach((row, index) => {
    try {
      const date = row[0].value;
      const monday = getMondayOfWeek(new Date(date));
      const weekKey = formatDateForAPI(monday);

      if (weekKey >= currentWeekStart) return;
      if (!shouldIncludeLastWeek && weekKey >= lastWeekStart) return;

      const campaign = row[1];
      const app = row[2];
      const metricsStartIndex = 3;
      
      const metrics = extractMetrics(row, metricsStartIndex);
      const sunday = getSundayOfWeek(new Date(date));
      const appKey = app.id;
      
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
      
      appData[appKey].weeks[weekKey].campaigns.push(campaignData);

    } catch (error) {}
  });

  return appData;
}

function extractMetrics(row, startIndex) {
  return {
    cpi: parseFloat(row[startIndex].value) || 0,
    installs: parseInt(row[startIndex + 1].value) || 0,
    ipm: parseFloat(row[startIndex + 2].value) || 0,
    spend: parseFloat(row[startIndex + 3].value) || 0,
    rrD1: parseFloat(row[startIndex + 4].value) || 0,
    roas: parseFloat(row[startIndex + 5].value) || 0,
    rrD7: parseFloat(row[startIndex + 6].value) || 0,
    roasD7: parseFloat(row[startIndex + 7].value) || 0,
    eArpuForecast: parseFloat(row[startIndex + 8].value) || 0,
    eRoasForecast: parseFloat(row[startIndex + 9].value) || 0,
    eProfitForecast: parseFloat(row[startIndex + 10].value) || 0,
    eRoasForecastD730: parseFloat(row[startIndex + 11].value) || 0
  };
}

function processProjectApiData(projectName, rawData, includeLastWeek = null) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    return processApiData(rawData, includeLastWeek);
  } finally {
    setCurrentProject(originalProject);
  }
}

function extractGeoFromCampaign(campaignName) {
  if (!campaignName) return 'OTHER';
  
 if (CURRENT_PROJECT === 'TRICKY' || CURRENT_PROJECT === 'REGULAR') {
  const geoMap = {
    '| USA |': 'USA',  // United States
    '| CAN |': 'CAN',  // Canada
    '| GBR |': 'GBR',  // United Kingdom
    '| DEU |': 'DEU',  // Germany
    '| FRA |': 'FRA',  // France
    '| ITA |': 'ITA',  // Italy
    '| ESP |': 'ESP',  // Spain
    '| AUS |': 'AUS',  // Australia
    '| NZL |': 'NZL',  // New Zealand
    '| JPN |': 'JPN',  // Japan
    '| KOR |': 'KOR',  // South Korea
    '| CHN |': 'CHN',  // China
    '| IND |': 'IND',  // India
    '| BRA |': 'BRA',  // Brazil
    '| MEX |': 'MEX',  // Mexico
    '| RUS |': 'RUS',  // Russia
    '| ZAF |': 'ZAF',  // South Africa
    '| SAU |': 'SAU',  // Saudi Arabia
    '| TUR |': 'TUR',  // Turkey
    '| NLD |': 'NLD',  // Netherlands
    '| BEL |': 'BEL',  // Belgium
    '| SWE |': 'SWE',  // Sweden
    '| NOR |': 'NOR',  // Norway
    '| DNK |': 'DNK',  // Denmark
    '| FIN |': 'FIN',  // Finland
    '| CHE |': 'CHE',  // Switzerland
    '| AUT |': 'AUT',  // Austria
    '| SGP |': 'SGP',  // Singapore
    '| ARE |': 'ARE',  // UAE
    '| ARG |': 'ARG',  // Argentina
    '| COL |': 'COL'   // Colombia
  };

  for (const [pattern, geo] of Object.entries(geoMap)) {
    if (campaignName.includes(pattern)) return geo;
  }
  return 'OTHER';
}
  
  if (CURRENT_PROJECT === 'OVERALL') return 'ALL';
  
  if (CURRENT_PROJECT === 'GOOGLE_ADS') {
    const geoPatterns = [
      { pattern: 'LatAm', geo: 'LatAm' }, { pattern: 'UK,GE', geo: 'UK,GE' }, { pattern: 'BR (PT)', geo: 'BR' },
      { pattern: 'US ', geo: 'US' }, { pattern: ' US ', geo: 'US' }, { pattern: 'WW ', geo: 'WW' },
      { pattern: ' WW ', geo: 'WW' }, { pattern: 'UK', geo: 'UK' }, { pattern: 'GE', geo: 'GE' }, { pattern: 'BR', geo: 'BR' }
    ];
    for (const {pattern, geo} of geoPatterns) {
      if (campaignName.includes(pattern)) return geo;
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
    if (CURRENT_PROJECT === 'OVERALL') return campaignName;
    if (campaignName.startsWith('APD_')) return campaignName;
    if (['REGULAR','GOOGLE_ADS','APPLOVIN','MINTEGRAL','INCENT'].includes(CURRENT_PROJECT)) return campaignName;
    
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
  BUNDLE_ID_CACHE = {};
  APPS_DB_CACHE = null;
  APPS_DB_CACHE_TIME = null;
  TRICKY_OPTIMIZED_CACHE = null;
}