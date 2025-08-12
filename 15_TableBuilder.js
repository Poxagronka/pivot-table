// Cache management
const CACHE = {
  weekTotals: new Map(),
  wowMetrics: new Map(),
  precomputedTotals: new Map(),
  precomputedRows: new Map(),
  wowKeys: new Map()
};

function clearTableBuilderCaches() {
  Object.values(CACHE).forEach(cache => cache.clear());
}

// Main building function
function buildUnifiedTable(data, tableData, formatData, wow, initialMetricsCache) {
  const startTime = Date.now();

  clearTableBuilderCaches();
  precomputeAllData(data, wow);
  
  const appsDbCache = CURRENT_PROJECT === 'TRICKY' ? loadAppsDbCache() : null;
  
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    buildIncentTrafficTable(data, tableData, formatData, initialMetricsCache);
  } else {
    buildStandardTable(data, tableData, formatData, initialMetricsCache, appsDbCache);
  }
  
  console.log(`buildUnifiedTable completed: ${tableData.length} rows in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
}

function loadAppsDbCache() {
  try {
    const appsDb = new AppsDatabase('TRICKY');
    return appsDb.loadFromCache();
  } catch (e) {
    console.error('Error loading AppsDatabase:', e);
    return {};
  }
}

function buildIncentTrafficTable(data, tableData, formatData, initialMetricsCache) {
  const networkKeys = Object.keys(data).sort((a, b) => 
    data[a].networkName.localeCompare(data[b].networkName)
  );
  
  networkKeys.forEach(networkKey => {
    const network = data[networkKey];
    
    // Add network row
    formatData.push({ row: tableData.length + 1, type: 'NETWORK' });
    tableData.push(createEmptyRow('NETWORK', network.networkName));
    
    // Process weeks
    Object.keys(network.weeks).sort().forEach(weekKey => {
      const week = network.weeks[weekKey];
      
      // Week row
      const weekTotals = getCachedTotals(`network_${networkKey}_${weekKey}`);
      const weekWoW = getCachedWoW(`${networkKey}_${weekKey}`, 'weekWoW');
      
      formatData.push({ row: tableData.length + 1, type: 'WEEK' });
      tableData.push(createDataRow('WEEK', week, weekTotals, weekWoW, network.networkName, initialMetricsCache));
      
      // App rows
      const appKeys = Object.keys(week.apps).sort((a, b) => {
        const spendA = week.apps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
        const spendB = week.apps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
        return spendB - spendA;
      });
      
      appKeys.forEach(appKey => {
        const app = week.apps[appKey];
        const appTotals = getCachedTotals(`incent_app_${networkKey}_${weekKey}_${appKey}`);
        const appWoW = getCachedWoW(`${networkKey}_${weekKey}_${appKey}`, 'appWoW');
        
        formatData.push({ row: tableData.length + 1, type: 'APP' });
        tableData.push(createDataRow('APP', { weekStart: week.weekStart, weekEnd: week.weekEnd }, 
          appTotals, appWoW, network.networkName, initialMetricsCache, app.appId, app.appName));
      });
    });
  });
}

function buildStandardTable(data, tableData, formatData, initialMetricsCache, appsDbCache) {
  const appKeys = Object.keys(data).sort((a, b) => data[a].appName.localeCompare(data[b].appName));
  
  appKeys.forEach(appKey => {
    const app = data[appKey];
    
    // Add app row
    formatData.push({ row: tableData.length + 1, type: 'APP' });
    tableData.push(createEmptyRow('APP', app.appName));
    
    // Process weeks
    Object.keys(app.weeks).sort().forEach(weekKey => {
      const week = app.weeks[weekKey];
      
      // Week row
      const weekTotals = getCachedTotals(`app_${appKey}_${weekKey}`);
      const weekWoW = getCachedWoW(`${app.appName}_${weekKey}`, 'appWeekWoW');
      
      formatData.push({ row: tableData.length + 1, type: 'WEEK' });
      tableData.push(createDataRow('WEEK', week, weekTotals, weekWoW, app.appName, initialMetricsCache));
      
      // Sub-rows (source apps, networks, or campaigns)
      addSubRows(tableData, week, weekKey, formatData, app.appName, initialMetricsCache, appsDbCache);
    });
  });
}

function addSubRows(tableData, week, weekKey, formatData, appName, initialMetricsCache, appsDbCache) {
  if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
    // Process source apps
    const sourceAppKeys = Object.keys(week.sourceApps).sort((a, b) => {
      const spendA = week.sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
      const spendB = week.sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
      return spendB - spendA;
    });
    
    sourceAppKeys.forEach(sourceAppKey => {
      const sourceApp = week.sourceApps[sourceAppKey];
      const totals = getCachedTotals(`sourceapp_${appName}_${weekKey}_${sourceApp.sourceAppId}`);
      const wow = getCachedWoW(`${sourceApp.sourceAppId}_${weekKey}`, 'sourceAppWoW');
      
      formatData.push({ row: tableData.length + 1, type: 'SOURCE_APP' });
      
      // Handle hyperlinks for TRICKY
      let displayName = sourceApp.sourceAppName;
      if (appsDbCache) {
        const appInfo = appsDbCache[sourceApp.sourceAppId];
        if (appInfo?.linkApp) {
          displayName = `=HYPERLINK("${appInfo.linkApp}", "${sourceApp.sourceAppName}")`;
          formatData.push({ row: tableData.length + 1, type: 'HYPERLINK' });
        }
      }
      
      tableData.push(createDataRow('SOURCE_APP', week, totals, wow, appName, 
        initialMetricsCache, sourceApp.sourceAppId, displayName));
      
      // Add campaigns
      addCampaignRows(tableData, sourceApp.campaigns, week, weekKey, formatData, appName, initialMetricsCache);
    });
    
  } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
    // Process networks for OVERALL
    const networkKeys = Object.keys(week.networks).sort((a, b) => {
      const spendA = week.networks[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
      const spendB = week.networks[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
      return spendB - spendA;
    });
    
    networkKeys.forEach(networkKey => {
      const network = week.networks[networkKey];
      const totals = getCachedTotals(`overall_network_${appName}_${weekKey}_${network.networkId}`);
      const wow = getCachedWoW(`${networkKey}_${weekKey}`, 'campaignWoW');
      
      formatData.push({ row: tableData.length + 1, type: 'NETWORK' });
      tableData.push(createDataRow('NETWORK', week, totals, wow, appName, 
        initialMetricsCache, network.networkId, network.networkName));
    });
    
  } else if (!['OVERALL', 'INCENT_TRAFFIC'].includes(CURRENT_PROJECT)) {
    // Add campaigns for other projects
    addCampaignRows(tableData, week.campaigns, week, weekKey, formatData, appName, initialMetricsCache);
  }
}

function addCampaignRows(tableData, campaigns, week, weekKey, formatData, appName, initialMetricsCache) {
  if (!campaigns || ['OVERALL', 'INCENT_TRAFFIC'].includes(CURRENT_PROJECT)) return;
  
  const sorted = campaigns.sort((a, b) => b.spend - a.spend);
  
  sorted.forEach(campaign => {
    const campaignIdValue = ['TRICKY', 'REGULAR'].includes(CURRENT_PROJECT) ? 
      `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")` :
      campaign.campaignId;
    
    const wow = getCachedWoW(`${campaign.campaignId}_${weekKey}`, 'campaignWoW');
    
    formatData.push({ row: tableData.length + 1, type: 'CAMPAIGN' });
    tableData.push(createCampaignRow(campaign, week, wow, appName, initialMetricsCache, campaignIdValue));
  });
}

// Row creation functions
function createEmptyRow(level, name) {
  const row = new Array(getUnifiedHeaders().length).fill('');
  row[0] = level;
  row[1] = name;
  return row;
}

function createDataRow(level, week, data, wow, appName, initialMetricsCache, identifier = '', displayName = '') {
  const headers = getUnifiedHeaders();
  const row = new Array(headers.length).fill('');
  
  row[0] = level;
  
  if (level === 'APP' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
    row[1] = displayName || identifier;
    return row;
  }
  
  // Common data for all levels
  const weekRange = `${week.weekStart} - ${week.weekEnd}`;
  const combinedRoas = `${data.avgRoasD1.toFixed(0)}% → ${data.avgRoasD3.toFixed(0)}% → ${data.avgRoasD7.toFixed(0)}% → ${data.avgRoasD30.toFixed(0)}%`;
  
  // Format eROAS and eProfit with initial values
  let eROAS730Display = `${data.avgEROASD730.toFixed(0)}%`;
  let eProfit730Display = formatSmartCurrency(data.totalProfit);
  
  if (initialMetricsCache && appName) {
    eROAS730Display = initialMetricsCache.formatEROASWithInitial(level, appName, weekRange, data.avgEROASD730, identifier, displayName);
    eProfit730Display = initialMetricsCache.formatProfitWithInitial(level, appName, weekRange, data.totalProfit, identifier, displayName);
  }
  
  // Fill row based on level
  row[1] = level === 'WEEK' ? weekRange : (displayName || identifier);
  row[2] = ''; // ID column - filled for campaigns
  row[3] = ''; // GEO column - filled for campaigns
  row[4] = formatSmartCurrency(data.totalSpend);
  row[5] = wow.spendChangePercent !== undefined ? `${wow.spendChangePercent.toFixed(0)}%` : '';
  row[6] = data.totalInstalls;
  row[7] = data.avgCpi.toFixed(3);
  row[8] = combinedRoas;
  row[9] = data.avgIpm.toFixed(1);
  row[10] = `${data.avgRrD1.toFixed(0)}%`;
  row[11] = `${data.avgRrD7.toFixed(0)}%`;
  row[12] = data.avgArpu.toFixed(3);
  row[13] = `${data.avgERoas.toFixed(0)}%`;
  row[14] = eROAS730Display;
  row[15] = eProfit730Display;
  row[16] = wow.eProfitChangePercent !== undefined ? `${wow.eProfitChangePercent.toFixed(0)}%` : '';
  row[17] = wow.growthStatus || '';
  row[18] = ''; // Comments
  
  return row;
}

function createCampaignRow(campaign, week, wow, appName, initialMetricsCache, campaignIdValue) {
  const headers = getUnifiedHeaders();
  const row = new Array(headers.length).fill('');
  const weekRange = `${week.weekStart} - ${week.weekEnd}`;
  
  row[0] = 'CAMPAIGN';
  row[1] = campaign.sourceApp;
  row[2] = campaignIdValue;
  row[3] = campaign.geo;
  row[4] = formatSmartCurrency(campaign.spend);
  row[5] = wow.spendChangePercent !== undefined ? `${wow.spendChangePercent.toFixed(0)}%` : '';
  row[6] = campaign.installs;
  row[7] = campaign.cpi ? campaign.cpi.toFixed(3) : '0.000';
  row[8] = `${campaign.roasD1.toFixed(0)}% → ${campaign.roasD3.toFixed(0)}% → ${campaign.roasD7.toFixed(0)}% → ${campaign.roasD30.toFixed(0)}%`;
  row[9] = campaign.ipm.toFixed(1);
  row[10] = `${campaign.rrD1.toFixed(0)}%`;
  row[11] = `${campaign.rrD7.toFixed(0)}%`;
  row[12] = campaign.eArpuForecast.toFixed(3);
  row[13] = `${campaign.eRoasForecast.toFixed(0)}%`;
  
  // Format with initial values
  if (initialMetricsCache && appName) {
    row[14] = initialMetricsCache.formatEROASWithInitial('CAMPAIGN', appName, weekRange, 
      campaign.eRoasForecastD730, campaign.campaignId, campaign.sourceApp);
    row[15] = initialMetricsCache.formatProfitWithInitial('CAMPAIGN', appName, weekRange, 
      campaign.eProfitForecast, campaign.campaignId, campaign.sourceApp);
  } else {
    row[14] = `${campaign.eRoasForecastD730.toFixed(0)}%`;
    row[15] = formatSmartCurrency(campaign.eProfitForecast);
  }
  
  row[16] = wow.eProfitChangePercent !== undefined ? `${wow.eProfitChangePercent.toFixed(0)}%` : '';
  row[17] = wow.growthStatus || '';
  row[18] = ''; // Comments
  
  return row;
}

// Cache functions
function precomputeAllData(data, wow) {
  precomputeTotals(data);
  precomputeWoW(wow);
}

function precomputeTotals(data) {
  const processLevel = (items, prefix) => {
    Object.entries(items).forEach(([key, item]) => {
      if (item.campaigns) {
        const cacheKey = `${prefix}_${key}`;
        CACHE.precomputedTotals.set(cacheKey, calculateWeekTotals(item.campaigns));
      }
      
      // Process nested structures
      if (item.weeks) processLevel(item.weeks, `${prefix}_${key}`);
      if (item.sourceApps) processLevel(item.sourceApps, `${prefix}_${key}`);
      if (item.apps) processLevel(item.apps, `${prefix}_${key}`);
      if (item.networks) processLevel(item.networks, `${prefix}_${key}`);
    });
  };
  
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    Object.keys(data).forEach(networkKey => {
      const network = data[networkKey];
      Object.keys(network.weeks).forEach(weekKey => {
        const week = network.weeks[weekKey];
        
        // Week totals
        const allCampaigns = [];
        Object.values(week.apps).forEach(app => allCampaigns.push(...app.campaigns));
        CACHE.precomputedTotals.set(`network_${networkKey}_${weekKey}`, calculateWeekTotals(allCampaigns));
        
        // App totals
        Object.keys(week.apps).forEach(appKey => {
          const app = week.apps[appKey];
          CACHE.precomputedTotals.set(`incent_app_${networkKey}_${weekKey}_${appKey}`, calculateWeekTotals(app.campaigns));
        });
      });
    });
  } else {
    processLevel(data, CURRENT_PROJECT === 'TRICKY' ? 'sourceapp' : 
                      CURRENT_PROJECT === 'OVERALL' ? 'overall_network' : 'app');
  }
}

function precomputeWoW(wow) {
  Object.entries(wow).forEach(([type, data]) => {
    Object.entries(data).forEach(([key, value]) => {
      CACHE.wowMetrics.set(`${type}_${key}`, value);
    });
  });
}

function getCachedTotals(cacheKey) {
  return CACHE.precomputedTotals.get(cacheKey) || calculateWeekTotals([]);
}

function getCachedWoW(key, type) {
  return CACHE.wowMetrics.get(`${type}_${key}`) || { spendChangePercent: 0, eProfitChangePercent: 0, growthStatus: 'First Week' };
}

function calculateWeekTotals(campaigns) {
  if (!campaigns || campaigns.length === 0) {
    return {
      totalSpend: 0, totalInstalls: 0, avgCpi: 0, avgRoasD1: 0, avgRoasD3: 0, avgRoasD7: 0, avgRoasD30: 0,
      avgIpm: 0, avgRrD1: 0, avgRrD7: 0, avgArpu: 0, avgERoas: 0, avgEROASD730: 0, totalProfit: 0
    };
  }
  
  const totalSpend = campaigns.reduce((s, c) => s + c.spend, 0);
  const totalInstalls = campaigns.reduce((s, c) => s + c.installs, 0);
  const avgCpi = totalInstalls ? totalSpend / totalInstalls : 0;
  
  const avgMetrics = ['roasD1', 'roasD3', 'roasD7', 'roasD30', 'ipm', 'rrD1', 'rrD7', 'eArpuForecast'];
  const averages = {};
  avgMetrics.forEach(metric => {
    averages[metric] = campaigns.length ? campaigns.reduce((s, c) => s + c[metric], 0) / campaigns.length : 0;
  });
  
  // Weighted eROAS calculations
  const calculateWeightedEROAS = (field, min = 1, max = 1000) => {
    const valid = campaigns.filter(c => c[field] >= min && c[field] <= max && c.spend > 0);
    if (!valid.length) return 0;
    const weighted = valid.reduce((sum, c) => sum + (c[field] * c.spend), 0);
    const totalSpend = valid.reduce((sum, c) => sum + c.spend, 0);
    return totalSpend > 0 ? weighted / totalSpend : 0;
  };
  
  return {
    totalSpend, totalInstalls, avgCpi,
    avgRoasD1: averages.roasD1, avgRoasD3: averages.roasD3,
    avgRoasD7: averages.roasD7, avgRoasD30: averages.roasD30,
    avgIpm: averages.ipm, avgRrD1: averages.rrD1, avgRrD7: averages.rrD7,
    avgArpu: averages.eArpuForecast,
    avgERoas: calculateWeightedEROAS('eRoasForecast'),
    avgEROASD730: calculateWeightedEROAS('eRoasForecastD730'),
    totalProfit: campaigns.reduce((s, c) => s + c.eProfitForecast, 0)
  };
}

// Headers function
function getUnifiedHeaders() {
  return TABLE_CONFIG.HEADERS;
}

// Legacy functions for compatibility
function addOptimizedSubRows(tableData, week, weekKey, formatData, appName = '', initialMetricsCache = null, appsDbCache = null) {
  return addSubRows(tableData, week, weekKey, formatData, appName, initialMetricsCache, appsDbCache);
}

function addOptimizedCampaignRows(tableData, campaigns, week, weekKey, formatData, appName = '', initialMetricsCache = null) {
  return addCampaignRows(tableData, campaigns, week, weekKey, formatData, appName, initialMetricsCache);
}

function addUnifiedSubRows(tableData, week, weekKey, wow, formatData, appName = '', initialMetricsCache = null, appsDbCache = null) {
  return addSubRows(tableData, week, weekKey, formatData, appName, initialMetricsCache, appsDbCache);
}

function addCampaignRowsBatched(tableData, campaigns, week, weekKey, wow, formatData, appName = '', initialMetricsCache = null) {
  return addCampaignRows(tableData, campaigns, week, weekKey, formatData, appName, initialMetricsCache);
}

function createUnifiedRow(level, week, data, spendWoW, profitWoW, status, appName = '', initialMetricsCache = null, identifier = '', displayName = '', campaignIdValue = '') {
  const wow = { spendChangePercent: parseFloat(spendWoW) || 0, eProfitChangePercent: parseFloat(profitWoW) || 0, growthStatus: status };
  
  if (level === 'CAMPAIGN') {
    return createCampaignRow(data, week, wow, appName, initialMetricsCache, campaignIdValue);
  } else {
    return createDataRow(level, week, data, wow, appName, initialMetricsCache, identifier, displayName);
  }
}

function getCachedWeekTotals(campaigns) {
  return calculateWeekTotals(campaigns);
}

function getPrecomputedTotals(campaigns, cacheKey) {
  return getCachedTotals(cacheKey);
}

function getOptimizedWoW(key, type) {
  return getCachedWoW(key, type);
}