// Cache management
var BUNDLE_ID_CACHE = new Map();
var BUNDLE_ID_CACHE_LOADED = false;
var BUNDLE_ID_CACHE_TIME = null;
var BUNDLE_ID_CACHE_DURATION = 21600000;
var APPS_DB_CACHE = null;
var APPS_DB_CACHE_TIME = null;
var BUNDLE_ID_CACHE_SHEET_ID = '1Z5pJgtg--9EACJL8PVZgJsmeUemv6PKhSsyx9ArChrM';

// ========== КОНФИГУРАЦИИ ==========

// Конфигурация парсеров строк для каждого проекта
const ROW_PARSERS = {
  INCENT_TRAFFIC: {
    structure: ['date', 'country', 'network', 'campaign', 'app'],
    metricsStart: 5,
    extractors: {
      countryCode: row => row[1]?.code || 'OTHER',
      countryName: row => row[1]?.value || 'Other',
      networkId: row => row[2]?.hid || row[2]?.id || 'unknown',
      networkName: row => row[2]?.value || row[2]?.name || 'Unknown Network',
      campaignId: row => row[3]?.hid || row[3]?.id || 'Unknown',
      campaignName: row => row[3]?.campaignName || row[3]?.value || 'Unknown',
      app: row => row[4]
    }
  },
  APPLOVIN_TEST: {
    structure: ['date', 'country', 'campaign', 'app'],
    metricsStart: 4,
    typeCheck: row => row[1]?.__typename === 'UaCampaignCountry',
    extractors: {
      countryCode: row => row[1]?.code || 'OTHER',
      countryName: row => row[1]?.country || row[1]?.value || 'Other',
      campaignId: row => row[2]?.hid || row[2]?.id || 'Unknown',
      campaignName: row => row[2]?.campaignName || 'Unknown',
      app: row => row[3]
    }
  },
  OVERALL: {
    structure: ['date', 'network', 'app'],
    metricsStart: 3,
    isNetworkLevel: true,
    extractors: {
      networkId: row => row[1]?.id || 'unknown',
      networkName: row => row[1]?.value || 'Unknown Network',
      app: row => row[2]
    }
  },
  DEFAULT: {
    structure: ['date', 'campaign', 'app'],
    metricsStart: 3,
    extractors: {
      campaignId: row => row[1]?.campaignId || row[1]?.hid || row[1]?.id || 'Unknown',
      campaignName: row => row[1]?.campaignName || row[1]?.value || 'Unknown',
      app: row => row[2]
    }
  }
};

// Конфигурация метрик
const METRICS_MAP = [
  ['cpi', 'float'], ['installs', 'int'], ['ipm', 'float'], ['spend', 'float'],
  ['rrD1', 'float'], ['roasD1', 'float'], ['roasD3', 'float'], ['rrD7', 'float'],
  ['roasD7', 'float'], ['roasD14', 'float'], ['roasD30', 'float'], ['eArpuForecast', 'float'],
  ['eRoasForecast', 'float'], ['eProfitForecast', 'float'], ['eRoasForecastD730', 'float']
];

// Процессоры данных для проектов
const DATA_PROCESSORS = {
  INCENT_TRAFFIC: {
    keyCountries: ['United States', 'United Kingdom', 'Germany', 'Canada',
                   'Australia', 'France', 'South Korea', 'Japan'],
    builder: 'networkBuilder'
  },
  APPLOVIN_TEST: {
    builder: 'appBuilder',
    postProcess: 'restructureToCampaignFirst'
  },
  TRICKY: {
    customProcessor: true
  },
  OVERALL: {
    builder: 'appWithNetworksBuilder'
  },
  DEFAULT: {
    builder: 'appBuilder'
  }
};

// GEO конфигурация
const GEO_CONFIGS = {
  TRICKY: {
    patterns: ['USA','MEX','AUS','DEU','JPN','KOR','BRA','CAN','GBR','FRA','ITA','ESP','RUS','CHN','IND',
               'TUR','POL','NLD','SWE','NOR','DNK','FIN','CHE','AUT','BEL','PRT','GRC','CZE','HUN','ROU',
               'BGR','HRV','SVK','SVN','LTU','LVA','EST','UKR','BLR','ISR','SAU','ARE','QAT','KWT','EGY',
               'ZAF','NGA','KEN','MAR','THA','VNM','IDN','MYS','SGP','PHL','TWN','HKG','ARG','CHL','COL',
               'PER','VEN','URY','ECU','BOL','PRY','CRI','GTM','DOM','PAN','NZL'],
    extract: name => {
      for (const geo of GEO_CONFIGS.TRICKY.patterns) {
        if (name.includes(`| ${geo} |`)) return geo;
      }
      return 'OTHER';
    }
  },
  GOOGLE_ADS: {
    patterns: [['LatAm','LatAm'],['UK,GE','UK,GE'],['BR (PT)','BR'],['US ','US'],[' US ','US'],
               ['WW ','WW'],[' WW ','WW'],['UK','UK'],['GE','GE'],['BR','BR']],
    extract: name => {
      for (const [pattern, geo] of GEO_CONFIGS.GOOGLE_ADS.patterns) {
        if (name.includes(pattern)) return geo;
      }
      return 'OTHER';
    }
  },
  DEFAULT: {
    patterns: ['WW_ru','WW_es','WW_de','WW_pt','Asia T1','T2-ES','T1-EN','LatAm',
               'TopGeo','Europe','US','RU','UK','GE','FR','PT','ES','DE','T1','WW'],
    extract: name => {
      const upper = name.toUpperCase();
      for (const pattern of GEO_CONFIGS.DEFAULT.patterns) {
        const up = pattern.toUpperCase();
        if (upper.includes('_'+up+'_') || upper.includes('-'+up+'-') || 
            upper.includes('_'+up) || upper.includes('-'+up) ||
            upper.includes(up+'_') || upper.includes(up+'-') || upper === up) {
          return pattern;
        }
      }
      return 'OTHER';
    }
  }
};

// ========== ОСНОВНЫЕ ФУНКЦИИ (сохраняем сигнатуры) ==========

function fetchCampaignData(dateRange, projectName = null) {
  const project = projectName || CURRENT_PROJECT;
  const config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
  const apiConfig = projectName ? getProjectApiConfig(projectName) : getCurrentApiConfig();
  
  if (!config.BEARER_TOKEN) throw new Error(`${project} missing BEARER_TOKEN`);
  if (!apiConfig.FILTERS.USER?.length) throw new Error(`${project} missing USER filters`);
  
  const payload = buildPayload(apiConfig, dateRange, project);
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
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
      'x-requested-with': 'XMLHttpRequest',
      'Trace-Id': Utilities.getUuid()
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  return executeWithRetry(config.API_URL, options, project);
}

function fetchProjectCampaignData(projectName, dateRange) {
  return fetchCampaignData(dateRange, projectName);
}

function buildPayload(apiConfig, dateRange, project) {
  const dateDim = ['GOOGLE_ADS','APPLOVIN','APPLOVIN_TEST','INCENT','INCENT_TRAFFIC','OVERALL'].includes(project) ? 'DATE' : 'INSTALL_DATE';
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
  
  return {
    operationName: apiConfig.OPERATION_NAME,
    variables: {
      dateFilters: [{ dimension: dateDim, from: dateRange.from, to: dateRange.to, include: true }],
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
}

function executeWithRetry(url, options, project, maxRetries = 3) {
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      
      if (code === 200) {
        const parsed = JSON.parse(resp.getContentText());
        if (parsed.errors?.length > 0) throw new Error(`GraphQL errors: ${JSON.stringify(parsed.errors)}`);
        console.log(`${project}: API request completed`);
        return parsed;
      }
      
      if (code >= 400 && code < 500) {
        const errors = { 401: 'Unauthorized', 403: 'Forbidden', 429: 'Rate limited' };
        throw new Error(errors[code] || `Client error ${code}`);
      }
      
      lastError = new Error(`Server error ${code}`);
      if (attempt < maxRetries) Utilities.sleep(Math.min(1000 * Math.pow(2, attempt - 1), 10000));
    } catch (e) {
      lastError = e;
      if (attempt < maxRetries && e.toString().includes('timed out')) {
        Utilities.sleep(2000 * attempt);
      }
    }
  }
  
  throw lastError || new Error('API request failed');
}

// ========== ОБРАБОТКА ДАННЫХ ==========

function processApiData(rawData) {
  const stats = rawData.data.analytics.richStats.stats;
  const today = new Date();
  const sundayOfCurrentWeek = getSundayOfWeek(today);
  const futureDate = formatDateForAPI(new Date(sundayOfCurrentWeek.getTime() + 24 * 60 * 60 * 1000)); // Понедельник следующей недели
  
  console.log(`Processing ${stats.length} records for ${CURRENT_PROJECT}...`);
  
  // Специальный процессор для TRICKY
  if (CURRENT_PROJECT === 'TRICKY') {
    return processTrickyStrategy(stats, sundayOfCurrentWeek);
  }
  
  // Универсальный процессор
  const processor = DATA_PROCESSORS[CURRENT_PROJECT] || DATA_PROCESSORS.DEFAULT;
  const data = processWithBuilder(stats, sundayOfCurrentWeek, processor);
  
  // Постобработка если нужна
  if (processor.postProcess === 'restructureToCampaignFirst') {
    return restructureToCampaignFirst(data);
  }
  
  return data;
}

function processWithBuilder(stats, sundayOfCurrentWeek, processor) {
  const builders = {
    networkBuilder: buildNetworkStructure,
    appBuilder: buildAppStructure,
    appWithNetworksBuilder: buildAppWithNetworksStructure
  };
  
  const builder = builders[processor.builder] || builders.appBuilder;
  return builder(stats, sundayOfCurrentWeek, processor);
}

// ========== BUILDERS ==========

function buildNetworkStructure(stats, sundayOfCurrentWeek, processor) {
  const networkData = {};
  const keyCountries = processor.keyCountries || [];
  
  stats.forEach(row => {
    const date = row[0].value;
    const weekKey = formatDateForAPI(getMondayOfWeek(new Date(date)));
    
    // Исключаем только будущие недели (после текущей)
    if (new Date(date) > sundayOfCurrentWeek) return;
    
    const data = parseRowUnified(row);
    
    // Обработка ключевых стран для INCENT_TRAFFIC
    const isKeyCountry = keyCountries.includes(data.countryName);
    const finalCountryCode = isKeyCountry ? data.countryCode : 'OTHER';
    const finalCountryName = isKeyCountry ? data.countryName : 'Other';
    
    // Инициализация структур
    ensureNestedStructure(networkData, [
      [data.networkId, () => ({ networkId: data.networkId, networkName: data.networkName, countries: {} })],
      [`${data.networkId}.countries.${finalCountryCode}`, () => ({ 
        countryCode: finalCountryCode, countryName: finalCountryName, campaigns: {} 
      })],
      [`${data.networkId}.countries.${finalCountryCode}.campaigns.${data.campaignId}`, () => ({ 
        campaignId: data.campaignId, campaignName: data.campaignName,
        sourceApp: data.sourceApp, geo: data.geo, weeks: {} 
      })],
      [`${data.networkId}.countries.${finalCountryCode}.campaigns.${data.campaignId}.weeks.${weekKey}`, () => ({
        weekStart: formatDateForAPI(getMondayOfWeek(new Date(date))),
        weekEnd: formatDateForAPI(getSundayOfWeek(new Date(date))),
        data: []
      })]
    ]);
    
    // Добавляем данные
    networkData[data.networkId].countries[finalCountryCode].campaigns[data.campaignId].weeks[weekKey].data.push({
      date, ...data.metrics, appId: data.app.id, appName: data.app.name,
      status: data.status, type: data.type, isAutomated: data.isAutomated
    });
  });
  
  return networkData;
}

function buildAppStructure(stats, sundayOfCurrentWeek) {
  const appData = {};
  
  stats.forEach(row => {
    const date = row[0].value;
    const weekKey = formatDateForAPI(getMondayOfWeek(new Date(date)));
    
    // Исключаем только будущие недели
    if (new Date(date) > sundayOfCurrentWeek) return;
    
    const data = parseRowUnified(row);
    const appKey = data.app.id;
    
    ensureNestedStructure(appData, [
      [appKey, () => ({
        appId: data.app.id, appName: data.app.name,
        platform: data.app.platform, bundleId: data.app.bundleId, weeks: {}
      })],
      [`${appKey}.weeks.${weekKey}`, () => ({
        weekStart: formatDateForAPI(getMondayOfWeek(new Date(date))),
        weekEnd: formatDateForAPI(getSundayOfWeek(new Date(date))),
        campaigns: []
      })]
    ]);
    
    appData[appKey].weeks[weekKey].campaigns.push({
      date, campaignId: data.campaignId, campaignName: data.campaignName,
      ...data.metrics, status: data.status, type: data.type,
      geo: data.geo, sourceApp: data.sourceApp, isAutomated: data.isAutomated,
      countryCode: data.countryCode, countryName: data.countryName
    });
  });
  
  return appData;
}

function buildAppWithNetworksStructure(stats, sundayOfCurrentWeek) {
  const appData = {};
  
  stats.forEach(row => {
    const date = row[0].value;
    const weekKey = formatDateForAPI(getMondayOfWeek(new Date(date)));
    
    // Исключаем только будущие недели
    if (new Date(date) > sundayOfCurrentWeek) return;
    
    const data = parseRowUnified(row);
    const appKey = data.app.id;
    const networkId = data.networkId || 'unknown';
    
    ensureNestedStructure(appData, [
      [appKey, () => ({
        appId: data.app.id, appName: data.app.name,
        platform: data.app.platform, bundleId: data.app.bundleId, weeks: {}
      })],
      [`${appKey}.weeks.${weekKey}`, () => ({
        weekStart: formatDateForAPI(getMondayOfWeek(new Date(date))),
        weekEnd: formatDateForAPI(getSundayOfWeek(new Date(date))),
        networks: {}
      })],
      [`${appKey}.weeks.${weekKey}.networks.${networkId}`, () => ({
        networkId, networkName: data.networkName, campaigns: []
      })]
    ]);
    
    appData[appKey].weeks[weekKey].networks[networkId].campaigns.push({
      date, campaignId: `network_${networkId}_${data.app.id}_${weekKey}`,
      campaignName: data.networkName, ...data.metrics,
      status: 'Active', type: 'Network', geo: 'ALL', 
      sourceApp: data.networkName, isAutomated: false
    });
  });
  
  return appData;
}

// ========== ПАРСЕРЫ ==========

function parseRowUnified(row) {
  const parser = ROW_PARSERS[CURRENT_PROJECT] || ROW_PARSERS.DEFAULT;
  
  // Проверка типов для APPLOVIN_TEST
  if (parser.typeCheck && !parser.typeCheck(row)) {
    return parseRowUnified.call({ CURRENT_PROJECT: 'DEFAULT' }, row);
  }
  
  // Извлечение базовых данных
  const result = {
    metrics: extractMetrics(row, parser.metricsStart),
    status: 'Unknown',
    type: 'Unknown',
    isAutomated: false
  };
  
  // Применение экстракторов
  Object.entries(parser.extractors).forEach(([key, extractor]) => {
    result[key] = extractor(row);
  });
  
  // Добавление campaign данных если есть
  const campaign = row[parser.structure.indexOf('campaign')];
  if (campaign) {
    result.status = campaign.status || 'Unknown';
    result.type = campaign.type || 'Unknown';
    result.isAutomated = campaign.isAutomated || false;
  }
  
  // GEO и sourceApp
  result.geo = extractGeoFromCampaign(result.campaignName);
  result.sourceApp = extractSourceApp(result.campaignName);
  
  return result;
}

function extractMetrics(row, startIndex) {
  const metrics = {};
  METRICS_MAP.forEach(([name, type], index) => {
    const value = row[startIndex + index]?.value || 0;
    metrics[name] = type === 'int' ? parseInt(value) || 0 : parseFloat(value) || 0;
  });
  return metrics;
}

// ========== УТИЛИТЫ ==========

function ensureNestedStructure(obj, pathConfigs) {
  pathConfigs.forEach(([path, initializer]) => {
    const parts = path.split('.');
    let current = obj;
    
    for (let i = 0; i < parts.length; i++) {
      if (i === 0 && !current[parts[0]]) {
        current[parts[0]] = initializer();
      } else if (i > 0) {
        if (!current[parts[i]]) {
          current[parts[i]] = i === parts.length - 1 ? initializer() : {};
        }
        current = current[parts[i]];
      } else {
        current = current[parts[0]];
      }
    }
  });
}

// ========== LEGACY ПАРСЕР (для обратной совместимости) ==========

function parseRow(row, isOverallOrIncent) {
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC' || CURRENT_PROJECT === 'APPLOVIN_TEST') {
    return parseRowUnified(row);
  }
  
  // Оригинальная логика для остальных проектов
  const campaign = (isOverallOrIncent && CURRENT_PROJECT !== 'INCENT_TRAFFIC') ? null : row[1];
  const network = (isOverallOrIncent && CURRENT_PROJECT !== 'INCENT_TRAFFIC') ? row[1] : null;
  const app = row[2];
  const metricsStartIndex = 3;
  
  const campaignName = campaign ? (campaign.campaignName || campaign.value || 'Unknown') : 'Unknown';
  const campaignId = campaign ? (campaign.campaignId || campaign.hid || campaign.id || 'Unknown') : 'Unknown';
  
  const metrics = extractMetrics(row, metricsStartIndex);
  
  return {
    campaign, network, app, campaignName, campaignId, metrics,
    geo: extractGeoFromCampaign(campaignName),
    sourceApp: extractSourceApp(campaignName),
    status: campaign?.status || 'Unknown',
    type: campaign?.type || 'Unknown',
    isAutomated: campaign?.isAutomated || false,
    countryCode: null,
    countryName: null
  };
}

// ========== СПЕЦИАЛЬНЫЕ ПРОЦЕССОРЫ (без изменений) ==========

function processTrickyStrategy(stats, sundayOfCurrentWeek) {
  ensureBundleIdCacheLoaded();
  const appsDbCache = getOptimizedAppsDbForTricky();
  const appData = {};
  const newBundleIds = new Map();
  
  stats.forEach(row => {
    const date = row[0].value;
    const weekKey = formatDateForAPI(getMondayOfWeek(new Date(date)));
    
    // Исключаем только будущие недели
    if (new Date(date) > sundayOfCurrentWeek) return;
    
    const data = parseRow(row, false);
    const bundleId = getCachedBundleId(data.campaignName, data.campaignId) || 'unknown';
    
    if (bundleId && !BUNDLE_ID_CACHE.has(data.campaignName)) {
      newBundleIds.set(data.campaignName, { campaignId: data.campaignId, bundleId });
    }
    
    const sourceAppDisplayName = getOptimizedSourceAppDisplayName(bundleId, appsDbCache);
    const appKey = data.app.id;
    
    if (!appData[appKey]) {
      appData[appKey] = {
        appId: data.app.id, appName: data.app.name,
        platform: data.app.platform, bundleId: data.app.bundleId, weeks: {}
      };
    }
    
    if (!appData[appKey].weeks[weekKey]) {
      const monday = getMondayOfWeek(new Date(date));
      const sunday = getSundayOfWeek(new Date(date));
      appData[appKey].weeks[weekKey] = {
        weekStart: formatDateForAPI(monday),
        weekEnd: formatDateForAPI(sunday),
        sourceApps: {}
      };
    }
    
    if (!appData[appKey].weeks[weekKey].sourceApps[bundleId]) {
      appData[appKey].weeks[weekKey].sourceApps[bundleId] = {
        sourceAppId: bundleId, sourceAppName: sourceAppDisplayName, campaigns: []
      };
    }
    
    appData[appKey].weeks[weekKey].sourceApps[bundleId].campaigns.push({
      date, campaignId: data.campaignId, campaignName: data.campaignName,
      ...data.metrics, status: data.status, type: data.type,
      geo: data.geo, sourceApp: data.sourceApp, isAutomated: data.isAutomated
    });
  });
  
  if (newBundleIds.size > 0) saveBundleIdCache(newBundleIds);
  console.log(`TRICKY: Processed ${stats.length} records`);
  return appData;
}

function restructureToCampaignFirst(appData) {
  const restructured = {};
  
  Object.keys(appData).forEach(appKey => {
    const app = appData[appKey];
    const campaignGroups = {};
    const allCampaigns = new Map();
    
    Object.keys(app.weeks).forEach(weekKey => {
      const week = app.weeks[weekKey];
      if (!week.campaigns) return;
      
      week.campaigns.forEach(campaign => {
        const campaignKey = campaign.campaignId;
        if (!allCampaigns.has(campaignKey)) {
          allCampaigns.set(campaignKey, {
            campaignId: campaign.campaignId,
            campaignName: campaign.campaignName,
            sourceApp: campaign.sourceApp
          });
        }
      });
    });
    
    allCampaigns.forEach((campaignInfo, campaignId) => {
      campaignGroups[campaignId] = {
        campaignId: campaignId,
        campaignName: campaignInfo.campaignName,
        sourceApp: campaignInfo.sourceApp,
        geo: '',
        weeks: {}
      };
      
      Object.keys(app.weeks).forEach(weekKey => {
        const week = app.weeks[weekKey];
        const campaignDataForWeek = week.campaigns.filter(c => c.campaignId === campaignId);
        
        if (campaignDataForWeek.length > 0) {
          campaignGroups[campaignId].weeks[weekKey] = {
            weekStart: week.weekStart,
            weekEnd: week.weekEnd,
            countries: {}
          };
          
          campaignDataForWeek.forEach(campaign => {
            const countryCode = campaign.countryCode || 'OTHER';
            const countryName = campaign.countryName || 'Other';
            
            if (!campaignGroups[campaignId].weeks[weekKey].countries[countryCode]) {
              campaignGroups[campaignId].weeks[weekKey].countries[countryCode] = {
                countryCode: countryCode,
                countryName: countryName,
                campaigns: []
              };
            }
            
            campaignGroups[campaignId].weeks[weekKey].countries[countryCode].campaigns.push(campaign);
          });
          
          const countriesArray = Object.values(campaignGroups[campaignId].weeks[weekKey].countries);
          countriesArray.forEach(country => {
            country.totalSpend = country.campaigns.reduce((sum, c) => sum + c.spend, 0);
          });
          countriesArray.sort((a, b) => b.totalSpend - a.totalSpend);
          
          const top10 = countriesArray.slice(0, 10);
          const others = countriesArray.slice(10);
          
          if (others.length > 0) {
            const othersSpend = others.reduce((sum, c) => sum + c.totalSpend, 0);
            const othersCampaigns = [];
            others.forEach(country => othersCampaigns.push(...country.campaigns));
            
            const othersAggregated = {
              countryCode: 'OTHERS',
              countryName: 'Others',
              totalSpend: othersSpend,
              campaigns: [{
                ...aggregateCampaigns(othersCampaigns),
                countryCode: 'OTHERS',
                countryName: 'Others'
              }]
            };
            
            const finalCountries = [];
            let othersInserted = false;
            
            for (const country of top10) {
              if (!othersInserted && othersSpend > country.totalSpend) {
                finalCountries.push(othersAggregated);
                othersInserted = true;
              }
              finalCountries.push(country);
            }
            
            if (!othersInserted && othersSpend > 0) {
              finalCountries.push(othersAggregated);
            }
            
            campaignGroups[campaignId].weeks[weekKey].countries = {};
            finalCountries.forEach(country => {
              campaignGroups[campaignId].weeks[weekKey].countries[country.countryCode] = country;
            });
          }
        }
      });
    });
    
    restructured[appKey] = {
      appId: app.appId,
      appName: app.appName,
      platform: app.platform,
      bundleId: app.bundleId,
      weeks: {},
      campaignGroups
    };
  });
  
  return restructured;
}

function aggregateCampaigns(campaigns) {
  const totalSpend = campaigns.reduce((s, c) => s + c.spend, 0);
  const totalInstalls = campaigns.reduce((s, c) => s + c.installs, 0);
  
  return {
    date: campaigns[0]?.date,
    campaignId: campaigns[0]?.campaignId,
    campaignName: campaigns[0]?.campaignName,
    spend: totalSpend,
    installs: totalInstalls,
    cpi: totalInstalls ? totalSpend / totalInstalls : 0,
    ipm: campaigns.reduce((s, c) => s + c.ipm, 0) / (campaigns.length || 1),
    roasD1: campaigns.reduce((s, c) => s + c.roasD1, 0) / (campaigns.length || 1),
    roasD3: campaigns.reduce((s, c) => s + c.roasD3, 0) / (campaigns.length || 1),
    roasD7: campaigns.reduce((s, c) => s + c.roasD7, 0) / (campaigns.length || 1),
    roasD14: campaigns.reduce((s, c) => s + c.roasD14, 0) / (campaigns.length || 1),
    roasD30: campaigns.reduce((s, c) => s + c.roasD30, 0) / (campaigns.length || 1),
    rrD1: campaigns.reduce((s, c) => s + c.rrD1, 0) / (campaigns.length || 1),
    rrD7: campaigns.reduce((s, c) => s + c.rrD7, 0) / (campaigns.length || 1),
    eArpuForecast: campaigns.reduce((s, c) => s + c.eArpuForecast, 0) / (campaigns.length || 1),
    eRoasForecast: campaigns.reduce((s, c) => s + c.eRoasForecast, 0) / (campaigns.length || 1),
    eProfitForecast: campaigns.reduce((s, c) => s + c.eProfitForecast, 0),
    eRoasForecastD730: campaigns.reduce((s, c) => s + c.eRoasForecastD730, 0) / (campaigns.length || 1),
    status: 'Active',
    type: 'Aggregated',
    geo: 'OTHERS',
    sourceApp: campaigns[0]?.sourceApp || 'Unknown',
    isAutomated: false
  };
}

// ========== GEO И SOURCE APP ==========

function extractGeoFromCampaign(campaignName) {
  if (!campaignName) return 'OTHER';
  if (['OVERALL','INCENT_TRAFFIC'].includes(CURRENT_PROJECT)) return 'ALL';
  
  const project = CURRENT_PROJECT === 'REGULAR' ? 'TRICKY' : CURRENT_PROJECT;
  const config = GEO_CONFIGS[project] || GEO_CONFIGS.DEFAULT;
  return config.extract(campaignName);
}

function extractSourceApp(campaignName) {
  if (['OVERALL','INCENT_TRAFFIC'].includes(CURRENT_PROJECT)) return campaignName;
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
}

// ========== BUNDLE ID КЕШИРОВАНИЕ (без изменений) ==========

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

// ========== LEGACY ФУНКЦИИ (для совместимости) ==========

function processProjectApiData(projectName, rawData) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    return processApiData(rawData);
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

// ========== GRAPHQL QUERY ==========

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