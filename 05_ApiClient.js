// Cache management
var BUNDLE_ID_CACHE = new Map();
var BUNDLE_ID_CACHE_LOADED = false;
var BUNDLE_ID_CACHE_TIME = null;
var BUNDLE_ID_CACHE_DURATION = 21600000;
var APPS_DB_CACHE = null;
var APPS_DB_CACHE_TIME = null;
var BUNDLE_ID_CACHE_SHEET_ID = '1Z5pJgtg--9EACJL8PVZgJsmeUemv6PKhSsyx9ArChrM';

// GEO configuration
const GEO_CONFIGS = {
  TRICKY: {
    patterns: ['USA','MEX','AUS','DEU','JPN','KOR','BRA','CAN','GBR','FRA','ITA','ESP','RUS','CHN','IND',
               'TUR','POL','NLD','SWE','NOR','DNK','FIN','CHE','AUT','BEL','PRT','GRC','CZE','HUN','ROU',
               'BGR','HRV','SVK','SVN','LTU','LVA','EST','UKR','BLR','ISR','SAU','ARE','QAT','KWT','EGY',
               'ZAF','NGA','KEN','MAR','THA','VNM','IDN','MYS','SGP','PHL','TWN','HKG','ARG','CHL','COL',
               'PER','VEN','URY','ECU','BOL','PRY','CRI','GTM','DOM','PAN','NZL'],
    extract: (name) => {
      for (const geo of GEO_CONFIGS.TRICKY.patterns) {
        if (name.includes(`| ${geo} |`)) return geo;
      }
      return 'OTHER';
    }
  },
  GOOGLE_ADS: {
    patterns: [['LatAm','LatAm'],['UK,GE','UK,GE'],['BR (PT)','BR'],['US ','US'],[' US ','US'],
               ['WW ','WW'],[' WW ','WW'],['UK','UK'],['GE','GE'],['BR','BR']],
    extract: (name) => {
      for (const [pattern, geo] of GEO_CONFIGS.GOOGLE_ADS.patterns) {
        if (name.includes(pattern)) return geo;
      }
      return 'OTHER';
    }
  },
  DEFAULT: {
    patterns: ['WW_ru','WW_es','WW_de','WW_pt','Asia T1','T2-ES','T1-EN','LatAm',
               'TopGeo','Europe','US','RU','UK','GE','FR','PT','ES','DE','T1','WW'],
    extract: (name) => {
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

// Main API functions (keep signatures!)
function fetchCampaignData(dateRange, projectName = null) {
  const project = projectName || CURRENT_PROJECT;
  
  // APPLOVIN_TEST должен использовать свою конфигурацию со странами!
  const config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
  const apiConfig = projectName ? getProjectApiConfig(projectName) : getCurrentApiConfig();
  
  if (!config.BEARER_TOKEN) throw new Error(`${project} missing BEARER_TOKEN`);
  if (!apiConfig.FILTERS.USER?.length) throw new Error(`${project} missing USER filters`);
  
  const payload = buildPayload(apiConfig, dateRange, project); // используем project
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

  return executeWithRetry(config.API_URL, options, project); // используем project
}

function fetchProjectCampaignData(projectName, dateRange) {
  return fetchCampaignData(dateRange, projectName);
}

// Simplified payload builder
function buildPayload(apiConfig, dateRange, project) {
  const dateDim = ['GOOGLE_ADS','APPLOVIN','APPLOVIN_TEST','INCENT','OVERALL'].includes(project) ? 'DATE' : 'INSTALL_DATE';
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

// Simplified retry logic
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

function processApiData(rawData, includeLastWeek = null) {
  const stats = rawData.data.analytics.richStats.stats;
  const today = new Date();
  const currentWeekStart = formatDateForAPI(getMondayOfWeek(today));
  const lastWeekStart = formatDateForAPI(getMondayOfWeek(new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000)));
  const shouldIncludeLastWeek = includeLastWeek !== null ? includeLastWeek : (today.getDay() >= 2 || today.getDay() === 0);
  
  console.log(`Processing ${stats.length} records for ${CURRENT_PROJECT}...`);
  
  // ДЕТАЛЬНАЯ ОТЛАДКА для APPLOVIN_TEST
  if (CURRENT_PROJECT === 'APPLOVIN_TEST') {
    console.log('=== APPLOVIN_TEST RAW DATA ANALYSIS ===');
    console.log('Total stats records:', stats.length);
    
    if (stats.length > 0) {
      // Показываем ПОЛНЫЙ первый объект для отладки
      console.log('\n=== FULL FIRST RECORD DEBUG ===');
      console.log('First record complete object:');
      console.log(JSON.stringify(stats[0], null, 2));
      console.log('=== END FULL RECORD ===\n');
      
      // Проверяем первые 3 записи
      for (let i = 0; i < Math.min(3, stats.length); i++) {
        console.log(`\nRecord ${i + 1}:`);
        const row = stats[i];
        console.log('  Row length:', row.length);
        
        row.forEach((item, index) => {
          if (item && typeof item === 'object') {
            console.log(`  [${index}]: __typename = ${item.__typename}`);
            
            // Детально логируем каждый тип
            if (item.__typename === 'UaCampaignCountry') {
              console.log(`    >>> COUNTRY FOUND! code: "${item.code}", country: "${item.country}"`);
            } else if (item.__typename === 'UaCampaign') {
              console.log(`    Campaign: "${item.campaignName}", hid: ${item.hid}`);
            } else if (item.__typename === 'AppInfo') {
              console.log(`    App: "${item.name}"`);
            } else if (item.__typename === 'StatsValue') {
              console.log(`    Value: ${item.value}`);
            }
          }
        });
      }
    }
    console.log('=== END RAW DATA ANALYSIS ===\n');
  }
  
  // Data processor strategy
  const processor = CURRENT_PROJECT === 'TRICKY' ? processTrickyStrategy : processStandardStrategy;
  const appData = processor(stats, currentWeekStart, lastWeekStart, shouldIncludeLastWeek);
  
  // Добавить отладку для APPLOVIN_TEST
  if (CURRENT_PROJECT === 'APPLOVIN_TEST') {
    console.log('APPLOVIN_TEST Debug: Before restructure');
    console.log('Total apps in appData:', Object.keys(appData).length);
    
    // Собираем статистику по странам
    const countryStats = {};
    let totalRecords = 0;
    let recordsWithCountry = 0;
    let recordsWithoutCountry = 0;
    
    Object.values(appData).forEach(app => {
      Object.values(app.weeks).forEach(week => {
        week.campaigns.forEach(c => {
          totalRecords++;
          
          if (c.countryCode && c.countryCode !== 'OTHER') {
            recordsWithCountry++;
          } else {
            recordsWithoutCountry++;
          }
          
          const cc = c.countryCode || 'OTHER';
          const cn = c.countryName || 'Other';
          
          if (!countryStats[cc]) {
            countryStats[cc] = { name: cn, count: 0, spend: 0 };
          }
          countryStats[cc].count++;
          countryStats[cc].spend += c.spend;
        });
      });
    });
    
    console.log(`Total records: ${totalRecords}, With country: ${recordsWithCountry}, Without: ${recordsWithoutCountry}`);
    console.log('Country distribution by spend:');
    
    const sortedCountries = Object.entries(countryStats)
      .sort((a, b) => b[1].spend - a[1].spend);
        
    sortedCountries.slice(0, 15).forEach(([code, stats]) => {
      console.log(`  ${code} (${stats.name}): ${stats.count} records, ${stats.spend.toFixed(2)}`);
    });
    
    if (sortedCountries.length > 15) {
      console.log(`  ... and ${sortedCountries.length - 15} more countries`);
    }
    
    // Детальная проверка первых записей
    const firstApp = Object.values(appData)[0];
    if (firstApp) {
      const firstWeek = Object.values(firstApp.weeks)[0];
      if (firstWeek && firstWeek.campaigns.length > 0) {
        console.log('First 3 campaign records:');
        firstWeek.campaigns.slice(0, 3).forEach((c, i) => {
          console.log(`  Record ${i + 1}:`);
          console.log(`    Campaign: ${c.campaignName}`);
          console.log(`    Country Code: ${c.countryCode}`);
          console.log(`    Country Name: ${c.countryName}`);
          console.log(`    Spend: ${c.spend}`);
        });
      }
    }
    
    return restructureToCampaignFirst(appData);
  }
  
  return CURRENT_PROJECT === 'INCENT_TRAFFIC' ? convertToNetworkStructure(appData) : appData;
}

function processStandardStrategy(stats, currentWeekStart, lastWeekStart, shouldIncludeLastWeek) {
  const appData = {};
  const isOverallOrIncent = ['OVERALL','INCENT_TRAFFIC'].includes(CURRENT_PROJECT);
  
  stats.forEach(row => {
    const date = row[0].value;
    const weekKey = formatDateForAPI(getMondayOfWeek(new Date(date)));
    
    if (weekKey >= currentWeekStart || (!shouldIncludeLastWeek && weekKey >= lastWeekStart)) return;
    
    const data = parseRow(row, isOverallOrIncent);
    const appKey = data.app.id;
    
    // Initialize structures
    if (!appData[appKey]) {
      appData[appKey] = {
        appId: data.app.id,
        appName: data.app.name,
        platform: data.app.platform,
        bundleId: data.app.bundleId,
        weeks: {}
      };
    }
    
    if (!appData[appKey].weeks[weekKey]) {
      const monday = getMondayOfWeek(new Date(date));
      const sunday = getSundayOfWeek(new Date(date));
      appData[appKey].weeks[weekKey] = {
        weekStart: formatDateForAPI(monday),
        weekEnd: formatDateForAPI(sunday),
        campaigns: [],
        networks: isOverallOrIncent ? {} : undefined
      };
    }
    
    // Add data
    if (isOverallOrIncent) {
      const networkId = data.network?.id || 'unknown';
      const networkName = data.network?.value || 'Unknown Network';
      
      if (!appData[appKey].weeks[weekKey].networks[networkId]) {
        appData[appKey].weeks[weekKey].networks[networkId] = {
          networkId, networkName, campaigns: []
        };
      }
      
      appData[appKey].weeks[weekKey].networks[networkId].campaigns.push({
        date, campaignId: `network_${networkId}_${data.app.id}_${weekKey}`,
        campaignName: networkName, ...data.metrics,
        status: 'Active', type: 'Network', geo: 'ALL', sourceApp: networkName, isAutomated: false
      });
    } else {
      // Для всех проектов, включая APPLOVIN_TEST с группировкой по странам
      const countryCode = data.countryCode || 'OTHER';
      const countryName = data.countryName || 'Other';
      
      appData[appKey].weeks[weekKey].campaigns.push({
        date, 
        campaignId: data.campaignId, 
        campaignName: data.campaignName,
        ...data.metrics, 
        status: data.status, 
        type: data.type,
        geo: data.geo, 
        sourceApp: data.sourceApp, 
        isAutomated: data.isAutomated,
        countryCode: countryCode,
        countryName: countryName
      });
    }
  });
  
  return CURRENT_PROJECT === 'INCENT_TRAFFIC' ? convertToNetworkStructure(appData) : appData;
}

function restructureToCampaignFirst(appData) {
  console.log('APPLOVIN_TEST: Starting restructure, apps count:', Object.keys(appData).length);
  
  // КРИТИЧЕСКАЯ ОТЛАДКА: проверяем наличие стран в данных
  let countriesFound = new Set();
  let samplesWithCountry = [];
  let samplesWithoutCountry = [];
  
  Object.values(appData).forEach(app => {
    Object.values(app.weeks).forEach(week => {
      week.campaigns.forEach(c => {
        if (c.countryCode && c.countryCode !== 'OTHER') {
          countriesFound.add(`${c.countryCode}:${c.countryName}`);
          if (samplesWithCountry.length < 3) {
            samplesWithCountry.push({
              campaign: c.campaignName,
              country: `${c.countryCode} - ${c.countryName}`,
              spend: c.spend
            });
          }
        } else if (samplesWithoutCountry.length < 3) {
          samplesWithoutCountry.push({
            campaign: c.campaignName,
            countryCode: c.countryCode,
            countryName: c.countryName
          });
        }
      });
    });
  });
  
  console.log('Countries found in data:', Array.from(countriesFound));
  if (samplesWithCountry.length > 0) {
    console.log('Sample records WITH country:', samplesWithCountry);
  }
  if (samplesWithoutCountry.length > 0) {
    console.log('Sample records WITHOUT country:', samplesWithoutCountry);
  }
  
  const restructured = {};
  
  Object.keys(appData).forEach(appKey => {
    const app = appData[appKey];
    const campaignGroups = {};
    
    console.log(`Processing app: ${app.appName}, weeks: ${Object.keys(app.weeks).length}`);
    
    // Собираем все уникальные кампании
    const allCampaigns = new Map();
    
    Object.keys(app.weeks).forEach(weekKey => {
      const week = app.weeks[weekKey];
      if (!week.campaigns) {
        console.log(`WARNING: No campaigns in week ${weekKey}`);
        return;
      }
      
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
    
    // Создаем структуру campaignGroups
    allCampaigns.forEach((campaignInfo, campaignId) => {
      campaignGroups[campaignId] = {
        campaignId: campaignId,
        campaignName: campaignInfo.campaignName,
        sourceApp: campaignInfo.sourceApp,
        geo: '', // GEO теперь на уровне стран, не кампаний
        weeks: {}
      };
      
      // Добавляем недели для каждой кампании
      Object.keys(app.weeks).forEach(weekKey => {
        const week = app.weeks[weekKey];
        const campaignDataForWeek = week.campaigns.filter(c => c.campaignId === campaignId);
        
        if (campaignDataForWeek.length > 0) {
          campaignGroups[campaignId].weeks[weekKey] = {
            weekStart: week.weekStart,
            weekEnd: week.weekEnd,
            countries: {}
          };
          
          // Группируем по странам
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
          
          // Обработка топ-10 стран для этой недели
          const countriesArray = Object.values(campaignGroups[campaignId].weeks[weekKey].countries);
          
          // Считаем спенд для каждой страны
          countriesArray.forEach(country => {
            country.totalSpend = country.campaigns.reduce((sum, c) => sum + c.spend, 0);
          });
          
          // Сортируем по спенду
          countriesArray.sort((a, b) => b.totalSpend - a.totalSpend);
          
          // Топ-10 и остальные
          const top10 = countriesArray.slice(0, 10);
          const others = countriesArray.slice(10);
          
          if (others.length > 0) {
            const othersSpend = others.reduce((sum, c) => sum + c.totalSpend, 0);
            const othersCampaigns = [];
            others.forEach(country => othersCampaigns.push(...country.campaigns));
            
            // Агрегируем метрики для Others
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
            
            // Создаем финальный список с правильной сортировкой
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
            
            // Перестраиваем countries объект с правильным порядком
            campaignGroups[campaignId].weeks[weekKey].countries = {};
            finalCountries.forEach(country => {
              campaignGroups[campaignId].weeks[weekKey].countries[country.countryCode] = country;
            });
          }
        }
      });
    });
    
    console.log(`App ${app.appName}: found ${Object.keys(campaignGroups).length} unique campaigns`);
    
    restructured[appKey] = {
      appId: app.appId,
      appName: app.appName,
      platform: app.platform,
      bundleId: app.bundleId,
      weeks: {},
      campaignGroups
    };
  });
  
  console.log('APPLOVIN_TEST: Restructure complete');
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


function processTrickyStrategy(stats, currentWeekStart, lastWeekStart, shouldIncludeLastWeek) {
  ensureBundleIdCacheLoaded();
  const appsDbCache = getOptimizedAppsDbForTricky();
  const appData = {};
  const newBundleIds = new Map();
  
  stats.forEach(row => {
    const date = row[0].value;
    const weekKey = formatDateForAPI(getMondayOfWeek(new Date(date)));
    
    if (weekKey >= currentWeekStart || (!shouldIncludeLastWeek && weekKey >= lastWeekStart)) return;
    
    const data = parseRow(row, false);
    const bundleId = getCachedBundleId(data.campaignName, data.campaignId) || 'unknown';
    
    if (bundleId && !BUNDLE_ID_CACHE.has(data.campaignName)) {
      newBundleIds.set(data.campaignName, { campaignId: data.campaignId, bundleId });
    }
    
    const sourceAppDisplayName = getOptimizedSourceAppDisplayName(bundleId, appsDbCache);
    const appKey = data.app.id;
    
    // Initialize structures
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

// Simplified row parser
function parseRow(row, isOverallOrIncent) {
  // СПЕЦИАЛЬНАЯ ОБРАБОТКА ДЛЯ APPLOVIN_TEST
  if (CURRENT_PROJECT === 'APPLOVIN_TEST') {
    // В новой структуре: [0] - дата, [1] - страна, [2] - кампания, [3] - приложение, [4+] - метрики
    const countryObj = row[1];
    const campaign = row[2];
    const app = row[3];
    
    // Проверяем, что у нас правильная структура
    if (countryObj && countryObj.__typename === 'UaCampaignCountry') {
      const campaignName = campaign ? (campaign.campaignName || 'Unknown') : 'Unknown';
      const campaignId = campaign ? (campaign.hid || campaign.id || 'Unknown') : 'Unknown';
      
      const metrics = {
        cpi: parseFloat(row[4]?.value || 0) || 0,
        installs: parseInt(row[5]?.value || 0) || 0,
        ipm: parseFloat(row[6]?.value || 0) || 0,
        spend: parseFloat(row[7]?.value || 0) || 0,
        rrD1: parseFloat(row[8]?.value || 0) || 0,
        roasD1: parseFloat(row[9]?.value || 0) || 0,
        roasD3: parseFloat(row[10]?.value || 0) || 0,
        rrD7: parseFloat(row[11]?.value || 0) || 0,
        roasD7: parseFloat(row[12]?.value || 0) || 0,
        roasD30: parseFloat(row[13]?.value || 0) || 0,
        eArpuForecast: parseFloat(row[14]?.value || 0) || 0,
        eRoasForecast: parseFloat(row[15]?.value || 0) || 0,
        eProfitForecast: parseFloat(row[16]?.value || 0) || 0,
        eRoasForecastD730: parseFloat(row[17]?.value || 0) || 0
      };
      
      return {
        campaign,
        network: null,
        app,
        campaignName,
        campaignId,
        metrics,
        geo: countryObj.code || 'OTHER',
        sourceApp: extractSourceApp(campaignName),
        status: campaign?.status || 'Unknown',
        type: campaign?.type || 'Unknown',
        isAutomated: campaign?.isAutomated || false,
        countryCode: countryObj.code,
        countryName: countryObj.country
      };
    }
  }
  
  // ОРИГИНАЛЬНАЯ ЛОГИКА ДЛЯ ДРУГИХ ПРОЕКТОВ
  const hasCountry = false; // Для других проектов стран нет
  const country = null;
  const campaign = isOverallOrIncent ? null : row[1];
  const network = isOverallOrIncent ? row[1] : null;
  const app = row[2];
  const metricsStartIndex = 3;
  
  const campaignName = campaign ? (campaign.campaignName || campaign.value || 'Unknown') : 'Unknown';
  const campaignId = campaign ? (campaign.campaignId || campaign.hid || campaign.id || 'Unknown') : 'Unknown';
  
  const metrics = {
    cpi: parseFloat(row[metricsStartIndex]?.value || 0) || 0,
    installs: parseInt(row[metricsStartIndex + 1]?.value || 0) || 0,
    ipm: parseFloat(row[metricsStartIndex + 2]?.value || 0) || 0,
    spend: parseFloat(row[metricsStartIndex + 3]?.value || 0) || 0,
    rrD1: parseFloat(row[metricsStartIndex + 4]?.value || 0) || 0,
    roasD1: parseFloat(row[metricsStartIndex + 5]?.value || 0) || 0,
    roasD3: parseFloat(row[metricsStartIndex + 6]?.value || 0) || 0,
    rrD7: parseFloat(row[metricsStartIndex + 7]?.value || 0) || 0,
    roasD7: parseFloat(row[metricsStartIndex + 8]?.value || 0) || 0,
    roasD30: parseFloat(row[metricsStartIndex + 9]?.value || 0) || 0,
    eArpuForecast: parseFloat(row[metricsStartIndex + 10]?.value || 0) || 0,
    eRoasForecast: parseFloat(row[metricsStartIndex + 11]?.value || 0) || 0,
    eProfitForecast: parseFloat(row[metricsStartIndex + 12]?.value || 0) || 0,
    eRoasForecastD730: parseFloat(row[metricsStartIndex + 13]?.value || 0) || 0
  };
  
  return {
    campaign,
    network,
    app,
    campaignName,
    campaignId,
    metrics,
    geo: extractGeoFromCampaign(campaignName),
    sourceApp: extractSourceApp(campaignName),
    status: campaign?.status || 'Unknown',
    type: campaign?.type || 'Unknown',
    isAutomated: campaign?.isAutomated || false,
    countryCode: null,
    countryName: null
  };
}

function convertToNetworkStructure(appData) {
  const networkData = {};
  
  Object.values(appData).forEach(app => {
    Object.values(app.weeks).forEach(week => {
      if (week.networks) {
        Object.values(week.networks).forEach(network => {
          const key = network.networkId;
          
          if (!networkData[key]) {
            networkData[key] = { networkId: network.networkId, networkName: network.networkName, weeks: {} };
          }
          
          if (!networkData[key].weeks[week.weekStart]) {
            networkData[key].weeks[week.weekStart] = {
              weekStart: week.weekStart, weekEnd: week.weekEnd, apps: {}
            };
          }
          
          networkData[key].weeks[week.weekStart].apps[app.appId] = {
            appId: app.appId, appName: app.appName,
            platform: app.platform, bundleId: app.bundleId,
            campaigns: network.campaigns
          };
        });
      }
    });
  });
  
  return networkData;
}

// GEO extraction (keep signature!)
function extractGeoFromCampaign(campaignName) {
  if (!campaignName) return 'OTHER';
  if (['OVERALL','INCENT_TRAFFIC'].includes(CURRENT_PROJECT)) return 'ALL';
  
  const project = CURRENT_PROJECT === 'REGULAR' ? 'TRICKY' : CURRENT_PROJECT;
  const config = GEO_CONFIGS[project] || GEO_CONFIGS.DEFAULT;
  return config.extract(campaignName);
}

// Source app extraction (keep signature!)
function extractSourceApp(campaignName) {
  if (['OVERALL','INCENT_TRAFFIC'].includes(CURRENT_PROJECT)) return campaignName;
  if (campaignName.startsWith('APD_')) return campaignName;
  if (['REGULAR','GOOGLE_ADS','APPLOVIN','MINTEGRAL','INCENT'].includes(CURRENT_PROJECT)) return campaignName;
  
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

// Bundle ID cache functions (keep all!)
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

// Legacy functions (keep all signatures!)
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

// GraphQL query (keep as is!)
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