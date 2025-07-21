/**
 * API Client - ОБНОВЛЕНО: обработка ROAS D-1, D-3, D-7, D-30 + поддержка сеток для OVERALL + отладка INCENT_TRAFFIC
 */

var BUNDLE_ID_CACHE = {};
var APPS_DB_CACHE = null;
var APPS_DB_CACHE_TIME = null;

function fetchCampaignData(dateRange) {
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
  
  const dateDimension = (CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN' || CURRENT_PROJECT === 'INCENT' || CURRENT_PROJECT === 'INCENT_TRAFFIC' || CURRENT_PROJECT === 'OVERALL') ? 'DATE' : 'INSTALL_DATE';
  
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
    payload: JSON.stringify(payload)
  };

  const resp = UrlFetchApp.fetch(config.API_URL, options);
  if (resp.getResponseCode() !== 200) {
    throw new Error('API request failed: ' + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
}

function fetchProjectCampaignData(projectName, dateRange) {
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
  
  const dateDimension = (projectName === 'GOOGLE_ADS' || projectName === 'APPLOVIN' || projectName === 'INCENT' || projectName === 'INCENT_TRAFFIC' || projectName === 'OVERALL') ? 'DATE' : 'INSTALL_DATE';
  
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
    payload: JSON.stringify(payload)
  };

  const resp = UrlFetchApp.fetch(config.API_URL, options);
  if (resp.getResponseCode() !== 200) {
    throw new Error(`${projectName} API request failed: ` + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
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

function getOptimizedAppsDb() {
  const now = new Date().getTime();
  
  if (APPS_DB_CACHE && APPS_DB_CACHE_TIME && (now - APPS_DB_CACHE_TIME) < 600000) {
    return APPS_DB_CACHE;
  }
  
  console.log('Loading Apps Database...');
  const appsDb = new AppsDatabase('TRICKY');
  appsDb.ensureCacheUpToDate();
  APPS_DB_CACHE = appsDb.loadFromCache();
  APPS_DB_CACHE_TIME = now;
  
  console.log(`Apps Database loaded: ${Object.keys(APPS_DB_CACHE).length} apps`);
  return APPS_DB_CACHE;
}

function getCachedBundleId(campaignName) {
  if (BUNDLE_ID_CACHE[campaignName]) {
    return BUNDLE_ID_CACHE[campaignName];
  }
  
  const bundleId = extractBundleIdFromCampaign(campaignName);
  BUNDLE_ID_CACHE[campaignName] = bundleId;
  return bundleId;
}

function getOptimizedSourceAppDisplayName(bundleId, appsDbCache) {
  if (!bundleId || CURRENT_PROJECT !== 'TRICKY') {
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
  const stats = rawData.data.analytics.richStats.stats;
  const appData = {};

  const today = new Date();
  const currentWeekStart = formatDateForAPI(getMondayOfWeek(today));
  const lastWeekStart = formatDateForAPI(getMondayOfWeek(new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000)));

  const dayOfWeek = today.getDay();
  const shouldIncludeLastWeek = includeLastWeek !== null ? includeLastWeek : (dayOfWeek >= 2 || dayOfWeek === 0);

  console.log(`Processing ${stats.length} records for ${CURRENT_PROJECT}...`);
  console.log(`Current week start: ${currentWeekStart}`);
  console.log(`Last week start: ${lastWeekStart}`);
  console.log(`Include last week: ${shouldIncludeLastWeek}`);

  let appsDbCache = null;
  if (CURRENT_PROJECT === 'TRICKY') {
    appsDbCache = getOptimizedAppsDb();
  }

  const trickyWeeklyData = {};
  let processedCount = 0;

  stats.forEach((row, index) => {
    try {
      const date = row[0].value;
      const monday = getMondayOfWeek(new Date(date));
      const weekKey = formatDateForAPI(monday);
      
      // Детальное логирование для INCENT_TRAFFIC
      if (CURRENT_PROJECT === 'INCENT_TRAFFIC' && index < 5) {
        console.log(`INCENT_TRAFFIC record ${index}:`);
        console.log(`  - date: ${date}`);
        console.log(`  - weekKey: ${weekKey}`);
        console.log(`  - currentWeekStart: ${currentWeekStart}`);
        console.log(`  - lastWeekStart: ${lastWeekStart}`);
        console.log(`  - shouldIncludeLastWeek: ${shouldIncludeLastWeek}`);
        console.log(`  - will be filtered: ${weekKey >= currentWeekStart || (!shouldIncludeLastWeek && weekKey >= lastWeekStart)}`);
      }

      if (weekKey >= currentWeekStart) {
        return;
      }
      
      if (!shouldIncludeLastWeek && weekKey >= lastWeekStart) {
        return;
      }

      let campaign, app, network, metricsStartIndex;
      
      // ОТЛАДКА: проверка условия для INCENT_TRAFFIC
      if (CURRENT_PROJECT === 'INCENT_TRAFFIC' && index < 5) {
        console.log(`INCENT_TRAFFIC processing row ${index}, checking condition:`, CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'INCENT_TRAFFIC');
      }
      
      if (CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'INCENT_TRAFFIC') {
        campaign = null;
        network = row[1];  // Attribution Network HID
        app = row[2];
        metricsStartIndex = 3;
        
        // Отладка для INCENT_TRAFFIC
        if (CURRENT_PROJECT === 'INCENT_TRAFFIC' && index < 5) {
          console.log('Processing as OVERALL/INCENT_TRAFFIC format');
          console.log(`INCENT_TRAFFIC row ${index} structure:`, {
            network: network ? `${network.__typename} (${network.id || network.value})` : 'null',
            app: app ? `${app.__typename} (${app.name})` : 'null',
            metricsCount: row.length - metricsStartIndex
          });
        }
      } else {
        campaign = row[1];
        app = row[2];
        network = null;
        metricsStartIndex = 3;
        
        // Отладка для INCENT_TRAFFIC - должно НЕ попадать сюда
        if (CURRENT_PROJECT === 'INCENT_TRAFFIC' && index < 5) {
          console.log('WARNING: INCENT_TRAFFIC processing as regular format!');
        }
      }
      
      // ОБНОВЛЕНО: новая структура метрик с ROAS D-1, D-3, D-7, D-30
      const metrics = {
        cpi: parseFloat(row[metricsStartIndex].value) || 0,         // 0: cpi
        installs: parseInt(row[metricsStartIndex + 1].value) || 0,  // 1: installs
        ipm: parseFloat(row[metricsStartIndex + 2].value) || 0,     // 2: ipm
        spend: parseFloat(row[metricsStartIndex + 3].value) || 0,   // 3: spend
        rrD1: parseFloat(row[metricsStartIndex + 4].value) || 0,    // 4: retention_rate D1
        roasD1: parseFloat(row[metricsStartIndex + 5].value) || 0,  // 5: roas D1
        roasD3: parseFloat(row[metricsStartIndex + 6].value) || 0,  // 6: roas D3
        rrD7: parseFloat(row[metricsStartIndex + 7].value) || 0,    // 7: retention_rate D7
        roasD7: parseFloat(row[metricsStartIndex + 8].value) || 0,  // 8: roas D7
        roasD30: parseFloat(row[metricsStartIndex + 9].value) || 0, // 9: roas D30
        eArpuForecast: parseFloat(row[metricsStartIndex + 10].value) || 0,  // 10: e_arpu_forecast D365
        eRoasForecast: parseFloat(row[metricsStartIndex + 11].value) || 0,  // 11: e_roas_forecast D365
        eProfitForecast: parseFloat(row[metricsStartIndex + 12].value) || 0, // 12: e_profit_forecast D730
        eRoasForecastD730: parseFloat(row[metricsStartIndex + 13].value) || 0 // 13: e_roas_forecast D730
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
            weeks: {}
          };
        }

        if (!appData[appKey].weeks[weekKey]) {
          appData[appKey].weeks[weekKey] = {
            weekStart: formatDateForAPI(monday),
            weekEnd: formatDateForAPI(sunday),
            networks: {} // Добавляем группировку по сеткам
          };
        }
        
        // Получаем информацию о сетке
        const networkId = network?.id || 'unknown';
        const networkName = network?.value || 'Unknown Network';
        
        if (!appData[appKey].weeks[weekKey].networks[networkId]) {
          appData[appKey].weeks[weekKey].networks[networkId] = {
            networkId: networkId,
            networkName: networkName,
            campaigns: []
          };
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
        
        appData[appKey].weeks[weekKey].networks[networkId].campaigns.push(networkCampaignData);
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

        if (CURRENT_PROJECT === 'TRICKY' && appsDbCache) {
          const bundleId = getCachedBundleId(campaignName);
          
          if (!trickyWeeklyData[appKey]) {
            trickyWeeklyData[appKey] = {
              appId: app.id,
              appName: app.name,
              platform: app.platform,
              bundleId: app.bundleId,
              weeks: {}
            };
          }
          
          if (!trickyWeeklyData[appKey].weeks[weekKey]) {
            trickyWeeklyData[appKey].weeks[weekKey] = {
              weekStart: formatDateForAPI(monday),
              weekEnd: formatDateForAPI(sunday),
              campaigns: []
            };
          }
          
          campaignData.extractedBundleId = bundleId;
          trickyWeeklyData[appKey].weeks[weekKey].campaigns.push(campaignData);
          
        } else {
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
        }

        processedCount++;
        
        if (processedCount % 100 === 0) {
          console.log(`Processed ${processedCount}/${stats.length} records...`);
        }
      }

    } catch (error) {
      console.error(`Error processing row ${index}:`, error);
    }
  });

  if (CURRENT_PROJECT === 'TRICKY' && appsDbCache) {
    console.log('Optimized TRICKY grouping...');
    
    Object.keys(trickyWeeklyData).forEach(appKey => {
      const appInfo = trickyWeeklyData[appKey];
      
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
          if (!bundleGroups[bundleId]) {
            bundleGroups[bundleId] = [];
          }
          bundleGroups[bundleId].push(campaign);
        });
        
        const sourceApps = {};
        const sortedBundleIds = Object.keys(bundleGroups).sort();
        
        sortedBundleIds.forEach(bundleId => {
          const campaigns = bundleGroups[bundleId];
          campaigns.sort((a, b) => b.spend - a.spend);
          
          const sourceAppDisplayName = getOptimizedSourceAppDisplayName(bundleId, appsDbCache);
          
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
    
    console.log('TRICKY optimized grouping completed');
  }

  // СПЕЦИАЛЬНАЯ ОБРАБОТКА ДЛЯ INCENT_TRAFFIC - перегруппировка сетка → неделя → приложение
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    console.log('=== INCENT_TRAFFIC Regrouping Debug ===');
    console.log('appData before regrouping:', {
      appCount: Object.keys(appData).length,
      apps: Object.keys(appData).map(key => appData[key].appName)
    });
    
    // Проверим структуру данных
    Object.values(appData).forEach(app => {
      console.log(`App: ${app.appName}`);
      Object.values(app.weeks).forEach(week => {
        console.log(`  Week: ${week.weekStart}`);
        if (week.networks) {
          console.log(`    Networks:`, Object.keys(week.networks));
        } else {
          console.log(`    NO NETWORKS! campaigns:`, week.campaigns?.length || 0);
        }
      });
    });
    
    const networkData = {};
    
    // Перегруппируем данные: сетка → неделя → приложение
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
        } else {
          console.log('WARNING: No networks in week data for INCENT_TRAFFIC!');
        }
      });
    });
    
    console.log('networkData after regrouping:', {
      networkCount: Object.keys(networkData).length,
      networks: Object.keys(networkData)
    });
    
    return networkData;
  }

  console.log(`Processing completed: ${processedCount} records processed`);
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
      '| JPN |': 'JPN', '| KOR |': 'KOR', '| BRA |': 'BRA', '| CAN |': 'CAN', '| GBR |': 'GBR'
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
  BUNDLE_ID_CACHE = {};
  APPS_DB_CACHE = null;
  APPS_DB_CACHE_TIME = null;
  console.log('TRICKY caches cleared');
}