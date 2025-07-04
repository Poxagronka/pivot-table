/**
 * API Client - Multi Project Support - ОБНОВЛЕНО: новая логика GEO для проектов кроме Tricky/Regular
 * Handles all API communication and data fetching
 */

/**
 * Fetch data from API with updated headers using current project configuration
 */
function fetchCampaignData(dateRange) {
  const config = getCurrentConfig();
  const apiConfig = getCurrentApiConfig();
  
  // Build filters array
  const filters = [
    { dimension: "USER", values: apiConfig.FILTERS.USER, include: true },
    { dimension: "ATTRIBUTION_PARTNER", values: apiConfig.FILTERS.ATTRIBUTION_PARTNER, include: true },
    { dimension: "ATTRIBUTION_NETWORK_HID", values: apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID, include: true }
  ];
  
  // Add campaign filter based on project type
  if (apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH) {
    const searchPattern = apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH;
    
    if (searchPattern.startsWith('!')) {
      // Negative filter (for Regular/Google_Ads/Applovin/Mintegral - exclude campaigns)
      const excludePattern = searchPattern.substring(1);
      filters.push({
        dimension: "ATTRIBUTION_CAMPAIGN_HID", 
        values: [], 
        include: false,
        searchByString: excludePattern
      });
    } else {
      // Positive filter (for Tricky - include campaigns)
      filters.push({
        dimension: "ATTRIBUTION_CAMPAIGN_HID", 
        values: [], 
        include: true, 
        searchByString: searchPattern
      });
    }
  }
  
  // Determine date dimension based on project
  const dateDimension = (CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN') ? 'DATE' : 'INSTALL_DATE';
  
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
      havingFilters: [],
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
    payload: JSON.stringify(payload)
  };

  const resp = UrlFetchApp.fetch(config.API_URL, options);
  if (resp.getResponseCode() !== 200) {
    throw new Error('API request failed: ' + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
}

/**
 * Fetch data for specific project
 */
function fetchProjectCampaignData(projectName, dateRange) {
  const config = getProjectConfig(projectName);
  const apiConfig = getProjectApiConfig(projectName);
  
  if (!config.BEARER_TOKEN) {
    throw new Error(`${projectName} project is not configured: missing BEARER_TOKEN`);
  }
  
  if (!apiConfig.FILTERS.USER || apiConfig.FILTERS.USER.length === 0) {
    throw new Error(`${projectName} project is not configured: missing USER filters`);
  }
  
  // Build filters array
  const filters = [
    { dimension: "USER", values: apiConfig.FILTERS.USER, include: true },
    { dimension: "ATTRIBUTION_PARTNER", values: apiConfig.FILTERS.ATTRIBUTION_PARTNER, include: true },
    { dimension: "ATTRIBUTION_NETWORK_HID", values: apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID, include: true }
  ];
  
  // Add campaign filter based on project type
  if (apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH) {
    const searchPattern = apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH;
    
    if (searchPattern.startsWith('!')) {
      // Negative filter (for Regular/Google_Ads/Applovin/Mintegral - exclude campaigns)
      const excludePattern = searchPattern.substring(1);
      filters.push({
        dimension: "ATTRIBUTION_CAMPAIGN_HID", 
        values: [], 
        include: false,
        searchByString: excludePattern
      });
    } else {
      // Positive filter (for Tricky - include campaigns)
      filters.push({
        dimension: "ATTRIBUTION_CAMPAIGN_HID", 
        values: [], 
        include: true, 
        searchByString: searchPattern
      });
    }
  }
  
  // Determine date dimension based on project
  const dateDimension = (projectName === 'GOOGLE_ADS' || projectName === 'APPLOVIN') ? 'DATE' : 'INSTALL_DATE';
  
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
      havingFilters: [],
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
    payload: JSON.stringify(payload)
  };

  const resp = UrlFetchApp.fetch(config.API_URL, options);
  if (resp.getResponseCode() !== 200) {
    throw new Error(`${projectName} API request failed: ` + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
}

/**
 * Get GraphQL query string
 */
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

/**
 * Process API data and group by apps, then weeks
 * Skip current week data as it's incomplete
 * ИСПРАВЛЕНО: Добавлен фильтр по спенду > 0 и поддержка всех проектов
 */
function processApiData(rawData) {
  const stats = rawData.data.analytics.richStats.stats;
  const appData = {};

  // Calculate current week start (Monday)
  const today = new Date();
  const currentWeekStart = formatDateForAPI(getMondayOfWeek(today));

  stats.forEach((row, index) => {
    try {
      const date = row[0].value;
      const monday = getMondayOfWeek(new Date(date));
      const weekKey = formatDateForAPI(monday);

      // Skip current week as it's incomplete
      if (weekKey >= currentWeekStart) {
        return;
      }

      const campaign = row[1];
      const app = row[2];
      
      // Extract metrics based on project type - Applovin uses same format as Google Ads
      let metrics;
      if (CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN') {
        // Google Ads/Applovin format: cpi, installs, spend, rr_d1, roas_d1, rr_d7, e_roas_365, e_profit_730
        metrics = {
          cpi: parseFloat(row[3].value) || 0,
          installs: parseInt(row[4].value) || 0,
          spend: parseFloat(row[5].value) || 0,
          rrD1: parseFloat(row[6].value) || 0,
          roas: parseFloat(row[7].value) || 0,
          rrD7: parseFloat(row[8].value) || 0,
          eRoasForecast: parseFloat(row[9].value) || 0,
          eProfitForecast: parseFloat(row[10].value) || 0,
          // Default values for missing metrics
          ipm: 0,
          eArpuForecast: 0
        };
      } else {
        // Original format: cpi, installs, ipm, spend, roas, e_arpu, e_roas, e_profit
        metrics = {
          cpi: parseFloat(row[3].value) || 0,
          installs: parseInt(row[4].value) || 0,
          ipm: parseFloat(row[5].value) || 0,
          spend: parseFloat(row[6].value) || 0,
          roas: parseFloat(row[7].value) || 0,
          eArpuForecast: parseFloat(row[8].value) || 0,
          eRoasForecast: parseFloat(row[9].value) || 0,
          eProfitForecast: parseFloat(row[10].value) || 0,
          // Default values for Google Ads/Applovin metrics
          rrD1: 0,
          rrD7: 0
        };
      }

      // ИСПРАВЛЕНО: Пропускаем кампании с нулевым или отрицательным спендом
      if (metrics.spend <= 0) {
        return;
      }

      const sunday = getSundayOfWeek(new Date(date));

      const appKey = app.id;
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

      // Extract campaign name - different structure for different projects
      let campaignName = 'Unknown';
      let campaignId = 'Unknown';
      
      if (campaign) {
        // For UaCampaign objects (Tricky/Regular/Google_Ads/Applovin/Mintegral)
        if (campaign.campaignName) {
          campaignName = campaign.campaignName;
          campaignId = campaign.campaignId || campaign.id || 'Unknown';
        }
        // For StatsValue objects (Moloco) 
        else if (campaign.value) {
          campaignName = campaign.value;
          campaignId = campaign.id || 'Unknown';
        }
      }

      const geo = extractGeoFromCampaign(campaignName);
      const sourceApp = extractSourceApp(campaignName);

      appData[appKey].weeks[weekKey].campaigns.push({
        date: date,
        campaignId: campaignId,
        campaignName: campaignName,
        ...metrics,
        status: campaign.status || 'Unknown',
        type: campaign.type || 'Unknown',
        geo,
        sourceApp,
        isAutomated: campaign.isAutomated || false
      });

    } catch (error) {
      console.error(`Error processing row ${index}:`, error);
    }
  });

  return appData;
}

/**
 * Process API data for specific project
 */
function processProjectApiData(projectName, rawData) {
  // Set current project context for processing
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    const result = processApiData(rawData);
    return result;
  } finally {
    // Restore original project context
    setCurrentProject(originalProject);
  }
}

/**
 * Extract geo information from campaign name
 * ОБНОВЛЕНО: Разная логика для разных проектов
 */
function extractGeoFromCampaign(campaignName) {
  if (!campaignName) return 'OTHER';
  
  // Для Tricky и Regular - используем старую логику
  if (CURRENT_PROJECT === 'TRICKY' || CURRENT_PROJECT === 'REGULAR') {
    const geoMap = {
      '| USA |': 'USA',
      '| MEX |': 'MEX',
      '| AUS |': 'AUS',
      '| DEU |': 'DEU',
      '| JPN |': 'JPN',
      '| KOR |': 'KOR',
      '| BRA |': 'BRA',
      '| CAN |': 'CAN',
      '| GBR |': 'GBR'
    };

    for (const [pattern, geo] of Object.entries(geoMap)) {
      if (campaignName.includes(pattern)) {
        return geo;
      }
    }
    return 'OTHER';
  } 
  
  // Для остальных проектов - новая логика поиска GEO в названии
  // Порядок важен! От более специфичных к менее специфичным
  const geoPatterns = [
    'WW_ru', 'WW_es', 'WW_de', 'WW_pt',  // Сначала более специфичные WW_
    'Asia T1', 'T2-ES', 'T1-EN',         // Затем составные
    'LatAm', 'TopGeo', 'Europe',         // Затем регионы
    'US', 'RU', 'UK', 'GE', 'FR', 'PT', 'ES', 'DE', 'T1', 'WW'  // Последними простые коды
  ];
  
  // Ищем паттерны в названии кампании (регистронезависимо)
  const upperCampaignName = campaignName.toUpperCase();
  
  for (const pattern of geoPatterns) {
    const upperPattern = pattern.toUpperCase();
    
    // Ищем точное совпадение как отдельное слово или разделенное подчеркиванием/дефисом
    if (upperCampaignName.includes('_' + upperPattern + '_') ||
        upperCampaignName.includes('-' + upperPattern + '-') ||
        upperCampaignName.includes('_' + upperPattern) ||
        upperCampaignName.includes('-' + upperPattern) ||
        upperCampaignName.includes(upperPattern + '_') ||
        upperCampaignName.includes(upperPattern + '-') ||
        upperCampaignName === upperPattern) {
      return pattern;
    }
  }
  
  return 'OTHER';
}

/**
 * Extract source app from campaign name
 * This function handles different campaign naming patterns for different projects
 */
function extractSourceApp(campaignName) {
  try {
    // Handle Moloco APD_ campaigns: show full campaign name
    if (campaignName.startsWith('APD_')) {
      return campaignName;
    }
    
    // Handle Regular, Google_Ads, Applovin, and Mintegral campaigns: DO NOT modify campaign names - return as is
    if (CURRENT_PROJECT === 'REGULAR' || CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN' || CURRENT_PROJECT === 'MINTEGRAL') {
      return campaignName;
    }
    
    // Handle Tricky campaigns (original logic)
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

/**
 * Project-specific source app extraction (for future use)
 */
function extractProjectSourceApp(projectName, campaignName) {
  // Set project context for extraction
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    return extractSourceApp(campaignName);
  } finally {
    setCurrentProject(originalProject);
  }
}

/**
 * Project-specific geo extraction (for future use)
 */
function extractProjectGeoFromCampaign(projectName, campaignName) {
  // Set project context for extraction
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    return extractGeoFromCampaign(campaignName);
  } finally {
    setCurrentProject(originalProject);
  }
}