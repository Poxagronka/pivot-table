const WEEK_TOTALS_CACHE = new Map();
const WOW_METRICS_CACHE = new Map();
const PRECOMPUTED_TOTALS = new Map();
const PRECOMPUTED_ROWS = new Map();
const WOW_KEYS_CACHE = new Map();

function buildUnifiedTable(data, tableData, formatData, wow, initialEROASCache) {
  const startTime = Date.now();
  console.log(`🔧 buildUnifiedTable started for ${CURRENT_PROJECT}`);

  clearTableBuilderCaches();
  
  console.log(`⏱️ Precomputing totals and WoW... (${((Date.now() - startTime) / 1000).toFixed(1)}s)`);
  precomputeAllTotals(data);
  precomputeWoWCache(wow);
  
  let appsDbCache = null;
  if (CURRENT_PROJECT === 'TRICKY') {
    console.log(`⏱️ Loading AppsDatabase cache... (${((Date.now() - startTime) / 1000).toFixed(1)}s)`);
    try {
      const appsDb = new AppsDatabase('TRICKY');
      appsDbCache = appsDb.loadFromCache();
      console.log(`📱 AppsDatabase loaded: ${Object.keys(appsDbCache).length} apps`);
    } catch (e) {
      console.error('Error loading AppsDatabase:', e);
      appsDbCache = {};
    }
  }
  
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    console.log(`⏱️ Processing INCENT_TRAFFIC networks... (${((Date.now() - startTime) / 1000).toFixed(1)}s)`);
    
    const networkKeys = Object.keys(data).sort((a, b) => 
      data[a].networkName.localeCompare(data[b].networkName)
    );
    console.log(`📊 Found ${networkKeys.length} networks to process`);
    
    networkKeys.forEach((networkKey, networkIndex) => {
      const network = data[networkKey];
      
      formatData.push({ row: tableData.length + 1, type: 'NETWORK' });
      const emptyRow = new Array(getUnifiedHeaders().length).fill('');
      emptyRow[0] = 'NETWORK';
      emptyRow[1] = network.networkName;
      tableData.push(emptyRow);
      
      const weekKeys = Object.keys(network.weeks).sort();
      console.log(`  Network ${networkIndex + 1}/${networkKeys.length}: ${network.networkName} (${weekKeys.length} weeks)`);
      
      weekKeys.forEach(weekKey => {
        const week = network.weeks[weekKey];
        
        const allCampaigns = [];
        Object.values(week.apps).forEach(app => {
          allCampaigns.push(...app.campaigns);
        });
        
        const weekTotals = getPrecomputedTotals(allCampaigns, `network_${networkKey}_${weekKey}`);
        const weekWoWKey = `${networkKey}_${weekKey}`;
        const weekWoW = getOptimizedWoW(weekWoWKey, 'weekWoW');
        
        const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
        const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
        const status = weekWoW.growthStatus || '';
        
        formatData.push({ row: tableData.length + 1, type: 'WEEK' });
        const weekRow = createUnifiedRow('WEEK', week, weekTotals, spendWoW, profitWoW, status, network.networkName, initialEROASCache);
        tableData.push(weekRow);
        
        const appKeys = Object.keys(week.apps).sort((a, b) => {
          const totalSpendA = week.apps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
          const totalSpendB = week.apps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
          return totalSpendB - totalSpendA;
        });
        
        appKeys.forEach(appKey => {
          const app = week.apps[appKey];
          const appTotals = getPrecomputedTotals(app.campaigns, `incent_app_${networkKey}_${weekKey}_${appKey}`);
          
          const appWoWKey = `${networkKey}_${weekKey}_${appKey}`;
          const appWoW = getOptimizedWoW(appWoWKey, 'appWoW');
          
          const spendWoW = appWoW.spendChangePercent !== undefined ? `${appWoW.spendChangePercent.toFixed(0)}%` : '';
          const profitWoW = appWoW.eProfitChangePercent !== undefined ? `${appWoW.eProfitChangePercent.toFixed(0)}%` : '';
          const status = appWoW.growthStatus || '';
          
          formatData.push({ row: tableData.length + 1, type: 'APP' });
          
          const appRow = createUnifiedRow('APP', { weekStart: week.weekStart, weekEnd: week.weekEnd }, appTotals, spendWoW, profitWoW, status, network.networkName, initialEROASCache, app.appId, app.appName);
          tableData.push(appRow);
        });
      });
    });

    console.log(`✅ INCENT_TRAFFIC processing completed in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
    
  } else {
    console.log(`⏱️ Processing regular projects... (${((Date.now() - startTime) / 1000).toFixed(1)}s)`);
    
    const appKeys = Object.keys(data).sort((a, b) => data[a].appName.localeCompare(data[b].appName));
    console.log(`📊 Found ${appKeys.length} apps to process`);
    
    appKeys.forEach((appKey, appIndex) => {
      const app = data[appKey];
      
      formatData.push({ row: tableData.length + 1, type: 'APP' });
      const emptyRow = new Array(getUnifiedHeaders().length).fill('');
      emptyRow[0] = 'APP';
      emptyRow[1] = app.appName;
      tableData.push(emptyRow);

      const weekKeys = Object.keys(app.weeks).sort();
      console.log(`  App ${appIndex + 1}/${appKeys.length}: ${app.appName} (${weekKeys.length} weeks)`);
      
      weekKeys.forEach((weekKey, weekIndex) => {
        const week = app.weeks[weekKey];
        
        formatData.push({ row: tableData.length + 1, type: 'WEEK' });
        
        let allCampaigns = [];
        if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
          Object.values(week.sourceApps).forEach(sourceApp => {
            allCampaigns.push(...sourceApp.campaigns);
          });
        } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
          Object.values(week.networks).forEach(network => {
            allCampaigns.push(...network.campaigns);
          });
        } else {
          allCampaigns = week.campaigns || [];
        }
        
        const weekTotals = getPrecomputedTotals(allCampaigns, `app_${appKey}_${weekKey}`);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = getOptimizedWoW(appWeekKey, 'appWeekWoW');
        
        const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
        const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
        const status = weekWoW.growthStatus || '';
        
        const weekRow = createUnifiedRow('WEEK', week, weekTotals, spendWoW, profitWoW, status, app.appName, initialEROASCache);
        tableData.push(weekRow);
        
        addOptimizedSubRows(tableData, week, weekKey, formatData, app.appName, initialEROASCache, appsDbCache);
      });
    });

    console.log(`✅ Regular projects processing completed in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
  }
  
  console.log(`🔧 buildUnifiedTable completed: ${tableData.length} rows in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
}

function precomputeAllTotals(data) {
  const startTime = Date.now();
  let computedCount = 0;
  
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    Object.keys(data).forEach(networkKey => {
      const network = data[networkKey];
      Object.keys(network.weeks).forEach(weekKey => {
        const week = network.weeks[weekKey];
        
        const allCampaigns = [];
        Object.values(week.apps).forEach(app => {
          allCampaigns.push(...app.campaigns);
        });
        
        const cacheKey = `network_${networkKey}_${weekKey}`;
        const totals = calculateWeekTotals(allCampaigns);
        PRECOMPUTED_TOTALS.set(cacheKey, totals);
        computedCount++;
        
        Object.keys(week.apps).forEach(appKey => {
          const app = week.apps[appKey];
          const appCacheKey = `incent_app_${networkKey}_${weekKey}_${appKey}`;
          const appTotals = calculateWeekTotals(app.campaigns);
          PRECOMPUTED_TOTALS.set(appCacheKey, appTotals);
          computedCount++;
        });
      });
    });
  } else {
    Object.keys(data).forEach(appKey => {
      const app = data[appKey];
      Object.keys(app.weeks).forEach(weekKey => {
        const week = app.weeks[weekKey];
        
        let allCampaigns = [];
        if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
          Object.values(week.sourceApps).forEach(sourceApp => {
            allCampaigns.push(...sourceApp.campaigns);
            
            const sourceAppCacheKey = `sourceapp_${appKey}_${weekKey}_${sourceApp.sourceAppId}`;
            const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
            PRECOMPUTED_TOTALS.set(sourceAppCacheKey, sourceAppTotals);
            computedCount++;
            
            WOW_KEYS_CACHE.set(`sourceApp_${sourceApp.sourceAppId}_${weekKey}`, `sourceAppWoW_${sourceApp.sourceAppId}_${weekKey}`);
            
            sourceApp.campaigns.forEach(campaign => {
              const campaignKey = `campaign_${campaign.campaignId}_${weekKey}`;
              WOW_KEYS_CACHE.set(campaignKey, `campaignWoW_${campaign.campaignId}_${weekKey}`);
            });
          });
        } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
          Object.values(week.networks).forEach(network => {
            allCampaigns.push(...network.campaigns);
            
            const networkCacheKey = `overall_network_${appKey}_${weekKey}_${network.networkId}`;
            const networkTotals = calculateWeekTotals(network.campaigns);
            PRECOMPUTED_TOTALS.set(networkCacheKey, networkTotals);
            computedCount++;
            
            WOW_KEYS_CACHE.set(`network_${network.networkId}_${weekKey}`, `campaignWoW_${network.networkId}_${weekKey}`);
          });
        } else {
          allCampaigns = week.campaigns || [];
          if (allCampaigns.length > 0) {
            allCampaigns.forEach(campaign => {
              const campaignKey = `campaign_${campaign.campaignId}_${weekKey}`;
              WOW_KEYS_CACHE.set(campaignKey, `campaignWoW_${campaign.campaignId}_${weekKey}`);
            });
          }
        }
        
        const cacheKey = `app_${appKey}_${weekKey}`;
        const totals = calculateWeekTotals(allCampaigns);
        PRECOMPUTED_TOTALS.set(cacheKey, totals);
        computedCount++;
        
        WOW_KEYS_CACHE.set(`appWeek_${app.appName}_${weekKey}`, `appWeekWoW_${app.appName}_${weekKey}`);
      });
    });
  }
  
  console.log(`🚀 Precomputed ${computedCount} totals and ${WOW_KEYS_CACHE.size} WoW keys in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
}

function precomputeWoWCache(wow) {
  const startTime = Date.now();
  let cacheCount = 0;
  
  if (wow.campaignWoW) {
    Object.keys(wow.campaignWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`campaignWoW_${key}`, wow.campaignWoW[key]);
      cacheCount++;
    });
  }
  
  if (wow.appWeekWoW) {
    Object.keys(wow.appWeekWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`appWeekWoW_${key}`, wow.appWeekWoW[key]);
      cacheCount++;
    });
  }
  
  if (wow.sourceAppWoW) {
    Object.keys(wow.sourceAppWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`sourceAppWoW_${key}`, wow.sourceAppWoW[key]);
      cacheCount++;
    });
  }
  
  if (wow.weekWoW) {
    Object.keys(wow.weekWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`weekWoW_${key}`, wow.weekWoW[key]);
      cacheCount++;
    });
  }
  
  if (wow.appWoW) {
    Object.keys(wow.appWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`appWoW_${key}`, wow.appWoW[key]);
      cacheCount++;
    });
  }
  
  if (wow.networkWoW) {
    Object.keys(wow.networkWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`networkWoW_${key}`, wow.networkWoW[key]);
      cacheCount++;
    });
  }
  
  console.log(`🚀 Precomputed ${cacheCount} WoW cache entries in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
}

function getOptimizedWoW(key, type) {
  const cacheKey = `${type}_${key}`;
  const cached = WOW_METRICS_CACHE.get(cacheKey);
  if (cached) {
    return cached;
  }
  
  return { spendChangePercent: 0, eProfitChangePercent: 0, growthStatus: 'First Week' };
}

function getCachedWoW(key, type, fallbackWow) {
  return getOptimizedWoW(key, type);
}

function getPrecomputedTotals(campaigns, cacheKey) {
  const cached = PRECOMPUTED_TOTALS.get(cacheKey);
  if (cached) {
    return cached;
  }
  
  return getCachedWeekTotals(campaigns);
}

function addOptimizedSubRows(tableData, week, weekKey, formatData, appName = '', initialEROASCache = null, appsDbCache = null) {
  if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
    const sourceAppKeys = Object.keys(week.sourceApps).sort((a, b) => {
      const totalSpendA = week.sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
      const totalSpendB = week.sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
      return totalSpendB - totalSpendA;
    });
    
    sourceAppKeys.forEach(sourceAppKey => {
      const sourceApp = week.sourceApps[sourceAppKey];
      const sourceAppTotals = getPrecomputedTotals(sourceApp.campaigns, `sourceapp_${appName}_${weekKey}_${sourceApp.sourceAppId}`);
      
      const sourceAppWoW = getOptimizedWoW(`${sourceApp.sourceAppId}_${weekKey}`, 'sourceAppWoW');
      
      const spendWoW = sourceAppWoW.spendChangePercent !== undefined ? `${sourceAppWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = sourceAppWoW.eProfitChangePercent !== undefined ? `${sourceAppWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = sourceAppWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'SOURCE_APP' });
      
      let sourceAppDisplayName = sourceApp.sourceAppName;
      if (CURRENT_PROJECT === 'TRICKY' && appsDbCache) {
        const appInfo = appsDbCache[sourceApp.sourceAppId];
        if (appInfo && appInfo.linkApp) {
          sourceAppDisplayName = `=HYPERLINK("${appInfo.linkApp}", "${sourceApp.sourceAppName}")`;
          formatData.push({ row: tableData.length + 1, type: 'HYPERLINK' });
        }
      }
      
      const sourceAppRow = createUnifiedRow('SOURCE_APP', 
        { weekStart: week.weekStart, weekEnd: week.weekEnd }, 
        sourceAppTotals, spendWoW, profitWoW, status, appName, 
        initialEROASCache, sourceApp.sourceAppId, sourceApp.sourceAppName);
      tableData.push(sourceAppRow);
      
      addOptimizedCampaignRows(tableData, sourceApp.campaigns, { weekStart: week.weekStart, weekEnd: week.weekEnd }, weekKey, formatData, appName, initialEROASCache);
    });
  } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
    const networkKeys = Object.keys(week.networks).sort((a, b) => {
      const totalSpendA = week.networks[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
      const totalSpendB = week.networks[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
      return totalSpendB - totalSpendA;
    });
    
    networkKeys.forEach(networkKey => {
      const network = week.networks[networkKey];
      const networkTotals = getPrecomputedTotals(network.campaigns, `overall_network_${appName}_${weekKey}_${network.networkId}`);
      
      const networkWoW = getOptimizedWoW(`${networkKey}_${weekKey}`, 'campaignWoW');
      
      const spendWoW = networkWoW.spendChangePercent !== undefined ? `${networkWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = networkWoW.eProfitChangePercent !== undefined ? `${networkWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = networkWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'NETWORK' });
      
      const networkRow = createUnifiedRow('NETWORK', 
        { weekStart: week.weekStart, weekEnd: week.weekEnd }, 
        networkTotals, spendWoW, profitWoW, status, appName, 
        initialEROASCache, network.networkId, network.networkName);
      tableData.push(networkRow);
    });
  } else if (CURRENT_PROJECT !== 'OVERALL' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
    addOptimizedCampaignRows(tableData, week.campaigns, week, weekKey, formatData, appName, initialEROASCache);
  }
}

function addOptimizedCampaignRows(tableData, campaigns, week, weekKey, formatData, appName = '', initialEROASCache = null) {
  if (CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    return;
  }
  
  const sortedCampaigns = campaigns.sort((a, b) => b.spend - a.spend);
  const batchSize = 1000;
  
  for (let i = 0; i < sortedCampaigns.length; i += batchSize) {
    const batch = sortedCampaigns.slice(i, i + batchSize);
    
    batch.forEach(campaign => {
      let campaignIdValue;
      if (CURRENT_PROJECT === 'TRICKY' || CURRENT_PROJECT === 'REGULAR') {
        campaignIdValue = `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`;
      } else {
        campaignIdValue = campaign.campaignId;
      }
      
      const campaignWoW = getOptimizedWoW(`${campaign.campaignId}_${weekKey}`, 'campaignWoW');
      
      const spendPct = campaignWoW.spendChangePercent !== undefined ? `${campaignWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitPct = campaignWoW.eProfitChangePercent !== undefined ? `${campaignWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const growthStatus = campaignWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'CAMPAIGN' });
      
      const campaignRow = createUnifiedRow('CAMPAIGN', week, campaign, spendPct, profitPct, growthStatus, appName, initialEROASCache, campaign.campaignId, campaign.sourceApp, campaignIdValue);
      tableData.push(campaignRow);
    });
  }
}

function addUnifiedSubRows(tableData, week, weekKey, wow, formatData, appName = '', initialEROASCache = null, appsDbCache = null) {
  return addOptimizedSubRows(tableData, week, weekKey, formatData, appName, initialEROASCache, appsDbCache);
}

function addCampaignRowsBatched(tableData, campaigns, week, weekKey, wow, formatData, appName = '', initialEROASCache = null) {
  return addOptimizedCampaignRows(tableData, campaigns, week, weekKey, formatData, appName, initialEROASCache);
}

function createUnifiedRow(level, week, data, spendWoW, profitWoW, status, appName = '', initialEROASCache = null, identifier = '', displayName = '', campaignIdValue = '') {
  const headers = getUnifiedHeaders();
  const row = new Array(headers.length).fill('');
  
  row[0] = level;
  
  if (level === 'APP' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
    row[1] = displayName || identifier;
    return row;
  } else if (level === 'WEEK') {
    row[1] = `${week.weekStart} - ${week.weekEnd}`;
    const combinedRoas = `${data.avgRoasD1.toFixed(0)}% → ${data.avgRoasD3.toFixed(0)}% → ${data.avgRoasD7.toFixed(0)}% → ${data.avgRoasD30.toFixed(0)}%`;
    
    let eROAS730Display = `${data.avgEROASD730.toFixed(0)}%`;
    if (initialEROASCache && appName) {
      const weekRange = `${week.weekStart} - ${week.weekEnd}`;
      eROAS730Display = initialEROASCache.formatEROASWithInitial('WEEK', appName, weekRange, data.avgEROASD730);
    }
    
    row[4] = formatSmartCurrency(data.totalSpend); row[5] = spendWoW; row[6] = data.totalInstalls; row[7] = data.avgCpi.toFixed(3);
    row[8] = combinedRoas; row[9] = data.avgIpm.toFixed(1); row[10] = `${data.avgRrD1.toFixed(0)}%`; row[11] = `${data.avgRrD7.toFixed(0)}%`;
    row[12] = data.avgArpu.toFixed(3); row[13] = `${data.avgERoas.toFixed(0)}%`; row[14] = eROAS730Display;
    row[15] = formatSmartCurrency(data.totalProfit); row[16] = profitWoW; row[17] = status;
  } else if (level === 'CAMPAIGN') {
    row[1] = data.sourceApp; row[2] = campaignIdValue; row[3] = data.geo;
    const combinedRoas = `${data.roasD1.toFixed(0)}% → ${data.roasD3.toFixed(0)}% → ${data.roasD7.toFixed(0)}% → ${data.roasD30.toFixed(0)}%`;
    
    let eROAS730Display = `${data.eRoasForecastD730.toFixed(0)}%`;
    if (initialEROASCache && appName) {
      const weekRange = `${week.weekStart} - ${week.weekEnd}`;
      eROAS730Display = initialEROASCache.formatEROASWithInitial('CAMPAIGN', appName, weekRange, data.eRoasForecastD730, data.campaignId, data.sourceApp);
    }
    
    row[4] = formatSmartCurrency(data.spend); row[5] = spendWoW; row[6] = data.installs; row[7] = data.cpi ? data.cpi.toFixed(3) : '0.000';
    row[8] = combinedRoas; row[9] = data.ipm.toFixed(1); row[10] = `${data.rrD1.toFixed(0)}%`; row[11] = `${data.rrD7.toFixed(0)}%`;
    row[12] = data.eArpuForecast.toFixed(3); row[13] = `${data.eRoasForecast.toFixed(0)}%`; row[14] = eROAS730Display;
    row[15] = formatSmartCurrency(data.eProfitForecast); row[16] = profitWoW; row[17] = status;
  } else {
    row[1] = displayName || identifier;
    const combinedRoas = `${data.avgRoasD1.toFixed(0)}% → ${data.avgRoasD3.toFixed(0)}% → ${data.avgRoasD7.toFixed(0)}% → ${data.avgRoasD30.toFixed(0)}%`;
    
    let eROAS730Display = `${data.avgEROASD730.toFixed(0)}%`;
    if (initialEROASCache && appName) {
      const weekRange = `${week.weekStart} - ${week.weekEnd}`;
      eROAS730Display = initialEROASCache.formatEROASWithInitial(level, appName, weekRange, data.avgEROASD730, identifier, displayName);
    }
    
    row[4] = formatSmartCurrency(data.totalSpend); row[5] = spendWoW; row[6] = data.totalInstalls; row[7] = data.avgCpi.toFixed(3);
    row[8] = combinedRoas; row[9] = data.avgIpm.toFixed(1); row[10] = `${data.avgRrD1.toFixed(0)}%`; row[11] = `${data.avgRrD7.toFixed(0)}%`;
    row[12] = data.avgArpu.toFixed(3); row[13] = `${data.avgERoas.toFixed(0)}%`; row[14] = eROAS730Display;
    row[15] = formatSmartCurrency(data.totalProfit); row[16] = profitWoW; row[17] = status;
  }
  
  row[18] = '';
  
  const weekRange = `${week.weekStart} - ${week.weekEnd}`;
  
  if (level === 'WEEK') {
    row[19] = generateCommentHash('WEEK', appName, weekRange);
  } else if (level === 'CAMPAIGN') {
    let cleanCampaignId = data.campaignId;
    if (CURRENT_PROJECT === 'TRICKY' && campaignIdValue && campaignIdValue.includes('HYPERLINK')) {
      const match = campaignIdValue.match(/campaigns\/([^"]+)/);
      cleanCampaignId = match ? match[1] : data.campaignId;
    }
    row[19] = generateDetailedCommentHash('CAMPAIGN', appName, weekRange, 
      cleanCampaignId, data.sourceApp, data.campaignName || data.sourceApp);
  } else if (level === 'SOURCE_APP') {
    row[19] = generateDetailedCommentHash('SOURCE_APP', appName, weekRange, 
      identifier, displayName || identifier, '');
  } else if (level === 'NETWORK') {
    row[19] = generateDetailedCommentHash('NETWORK', appName, weekRange, 
      identifier, '', displayName || identifier);
  } else if (level === 'APP') {
    row[19] = generateDetailedCommentHash('APP', appName, weekRange, 
      identifier, displayName || identifier, '');
  } else {
    row[19] = '';
  }
  
  return row;
}

function getCachedWeekTotals(campaigns) {
  if (!campaigns || campaigns.length === 0) {
    return {
      totalSpend: 0, totalInstalls: 0, avgCpi: 0, avgRoasD1: 0, avgRoasD3: 0, avgRoasD7: 0, avgRoasD30: 0,
      avgIpm: 0, avgRrD1: 0, avgRrD7: 0, avgArpu: 0, avgERoas: 0, avgEROASD730: 0, totalProfit: 0
    };
  }
  
  const cacheKey = campaigns.map(c => `${c.campaignId}_${c.spend}_${c.installs}`).join('|');
  
  if (WEEK_TOTALS_CACHE.has(cacheKey)) {
    return WEEK_TOTALS_CACHE.get(cacheKey);
  }
  
  const result = calculateWeekTotals(campaigns);
  WEEK_TOTALS_CACHE.set(cacheKey, result);
  return result;
}

function calculateWeekTotals(campaigns) {
  const totalSpend = campaigns.reduce((s, c) => s + c.spend, 0);
  const totalInstalls = campaigns.reduce((s, c) => s + c.installs, 0);
  const avgCpi = totalInstalls ? totalSpend / totalInstalls : 0;
  
  const avgRoasD1 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD1, 0) / campaigns.length : 0;
  const avgRoasD3 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD3, 0) / campaigns.length : 0;
  const avgRoasD7 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD7, 0) / campaigns.length : 0;
  const avgRoasD30 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD30, 0) / campaigns.length : 0;
  
  const avgIpm = campaigns.length ? campaigns.reduce((s, c) => s + c.ipm, 0) / campaigns.length : 0;
  const avgRrD1 = campaigns.length ? campaigns.reduce((s, c) => s + c.rrD1, 0) / campaigns.length : 0;
  const avgRrD7 = campaigns.length ? campaigns.reduce((s, c) => s + c.rrD7, 0) / campaigns.length : 0;
  const avgArpu = campaigns.length ? campaigns.reduce((s, c) => s + c.eArpuForecast, 0) / campaigns.length : 0;
  
  const validForEROAS = campaigns.filter(c => 
    c.eRoasForecast >= 1 && 
    c.eRoasForecast <= 1000 && 
    c.spend > 0
  );
  
  let avgERoas = 0;
  if (validForEROAS.length > 0) {
    const totalWeightedEROAS = validForEROAS.reduce((sum, c) => sum + (c.eRoasForecast * c.spend), 0);
    const totalSpendForEROAS = validForEROAS.reduce((sum, c) => sum + c.spend, 0);
    avgERoas = totalSpendForEROAS > 0 ? totalWeightedEROAS / totalSpendForEROAS : 0;
  }
  
  const validForEROASD730 = campaigns.filter(c => 
    c.eRoasForecastD730 >= 1 && 
    c.eRoasForecastD730 <= 1000 && 
    c.spend > 0
  );
  
  let avgEROASD730 = 0;
  if (validForEROASD730.length > 0) {
    const totalWeightedEROASD730 = validForEROASD730.reduce((sum, c) => sum + (c.eRoasForecastD730 * c.spend), 0);
    const totalSpendForEROASD730 = validForEROASD730.reduce((sum, c) => sum + c.spend, 0);
    avgEROASD730 = totalSpendForEROASD730 > 0 ? totalWeightedEROASD730 / totalSpendForEROASD730 : 0;
  }
  
  const totalProfit = campaigns.reduce((s, c) => s + c.eProfitForecast, 0);

  return {
    totalSpend, totalInstalls, avgCpi, avgRoasD1, avgRoasD3, avgRoasD7, avgRoasD30, avgIpm, avgRrD1, avgRrD7,
    avgArpu, avgERoas, avgEROASD730, totalProfit
  };
}

function clearTableBuilderCaches() {
  WEEK_TOTALS_CACHE.clear();
  WOW_METRICS_CACHE.clear();
  PRECOMPUTED_TOTALS.clear();
  PRECOMPUTED_ROWS.clear();
  WOW_KEYS_CACHE.clear();
}

function getUnifiedHeaders() {
  return [
    'Level', 'Week Range / Source App', 'ID', 'GEO',
    'Spend', 'Spend WoW %', 'Installs', 'CPI', 'ROAS D1→D3→D7→D30', 'IPM',
    'RR D-1', 'RR D-7', 'eARPU 365d', 'eROAS 365d', 'eROAS 730d', 'eProfit 730d', 'eProfit 730d WoW %', 'Growth Status', 'Comments', 'RowHash'
  ];
}