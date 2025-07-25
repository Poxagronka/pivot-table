const WEEK_TOTALS_CACHE = new Map();
const TABLE_APPS_DB_CACHE = new Map();
const TABLE_DISPLAY_NAME_CACHE = new Map();

function buildUnifiedTable(data, tableData, formatData, wow, initialEROASCache) {
  const startTime = Date.now();
  console.log(`🔧 buildUnifiedTable started for ${CURRENT_PROJECT}`);

  clearTableBuilderCaches();
  
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
        
        const weekTotals = getCachedWeekTotals(allCampaigns);
        const weekWoWKey = `${networkKey}_${weekKey}`;
        const weekWoW = wow.weekWoW[weekWoWKey] || {};
        
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
          const appTotals = getCachedWeekTotals(app.campaigns);
          
          const appWoWKey = `${networkKey}_${weekKey}_${appKey}`;
          const appWoW = wow.appWoW[appWoWKey] || {};
          
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
    
    preloadAppsDbCache();
    
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
      
      const appStartTime = Date.now();
      
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
        
        const weekTotals = getCachedWeekTotals(allCampaigns);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = wow.appWeekWoW[appWeekKey] || {};
        
        const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
        const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
        const status = weekWoW.growthStatus || '';
        
        const weekRow = createUnifiedRow('WEEK', week, weekTotals, spendWoW, profitWoW, status, app.appName, initialEROASCache);
        tableData.push(weekRow);
        
        const subRowStartTime = Date.now();
        addUnifiedSubRows(tableData, week, weekKey, wow, formatData, app.appName, initialEROASCache);
        
        if (weekIndex === 0 || (Date.now() - subRowStartTime) > 1000) {
          console.log(`    Week ${weekIndex + 1}/${weekKeys.length}: ${weekKey} processed in ${((Date.now() - subRowStartTime) / 1000).toFixed(1)}s`);
        }
      });
      
      const appProcessTime = (Date.now() - appStartTime) / 1000;
      if (appProcessTime > 5) {
        console.log(`  ⚠️  App ${appIndex + 1} (${app.appName}) took ${appProcessTime.toFixed(1)}s - slow!`);
      }
    });

    console.log(`✅ Regular projects processing completed in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
  }
  
  console.log(`🔧 buildUnifiedTable completed: ${tableData.length} rows in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
}

function addUnifiedSubRows(tableData, week, weekKey, wow, formatData, appName = '', initialEROASCache = null) {
  const startTime = Date.now();
  
  if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
    const sourceAppKeys = Object.keys(week.sourceApps).sort((a, b) => {
      const totalSpendA = week.sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
      const totalSpendB = week.sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
      return totalSpendB - totalSpendA;
    });
    
    console.log(`      Processing ${sourceAppKeys.length} source apps for TRICKY...`);
    
    sourceAppKeys.forEach(sourceAppKey => {
      const sourceApp = week.sourceApps[sourceAppKey];
      const sourceAppTotals = getCachedWeekTotals(sourceApp.campaigns);
      
      const sourceAppWoWKey = `${sourceApp.sourceAppId}_${weekKey}`;
      const sourceAppWoW = wow.sourceAppWoW[sourceAppWoWKey] || {};
      
      const spendWoW = sourceAppWoW.spendChangePercent !== undefined ? `${sourceAppWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = sourceAppWoW.eProfitChangePercent !== undefined ? `${sourceAppWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = sourceAppWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'SOURCE_APP' });
      
      const sourceAppDisplayInfo = getCachedDisplayName(sourceApp.sourceAppId, sourceApp.sourceAppName);
      
      if (sourceAppDisplayInfo.isHyperlink) {
        formatData.push({ row: tableData.length + 1, type: 'HYPERLINK' });
      }
      
      const sourceAppRow = createUnifiedRow('SOURCE_APP', week, sourceAppTotals, spendWoW, profitWoW, status, appName, initialEROASCache, sourceApp.sourceAppId, sourceAppDisplayInfo.displayName);
      tableData.push(sourceAppRow);
      
      const campaignStartTime = Date.now();
      addCampaignRowsBatched(tableData, sourceApp.campaigns, { weekStart: weekKey.split('-').join('/'), weekEnd: '' }, weekKey, wow, formatData, appName, initialEROASCache);
      
      if (sourceApp.campaigns.length > 50 || (Date.now() - campaignStartTime) > 500) {
        console.log(`        Source app ${sourceApp.sourceAppName}: ${sourceApp.campaigns.length} campaigns in ${((Date.now() - campaignStartTime) / 1000).toFixed(1)}s`);
      }
    });
  } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
    const networkKeys = Object.keys(week.networks).sort((a, b) => {
      const totalSpendA = week.networks[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
      const totalSpendB = week.networks[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
      return totalSpendB - totalSpendA;
    });
    
    networkKeys.forEach(networkKey => {
      const network = week.networks[networkKey];
      const networkTotals = getCachedWeekTotals(network.campaigns);
      
      const networkWoWKey = `${networkKey}_${weekKey}`;
      const networkWoW = wow.campaignWoW[networkWoWKey] || {};
      
      const spendWoW = networkWoW.spendChangePercent !== undefined ? `${networkWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = networkWoW.eProfitChangePercent !== undefined ? `${networkWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = networkWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'NETWORK' });
      
      const networkRow = createUnifiedRow('NETWORK', week, networkTotals, spendWoW, profitWoW, status, appName, initialEROASCache, network.networkId, network.networkName);
      tableData.push(networkRow);
    });
  } else if (CURRENT_PROJECT !== 'OVERALL' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
    const campaignStartTime = Date.now();
    addCampaignRowsBatched(tableData, week.campaigns, week, weekKey, wow, formatData, appName, initialEROASCache);
    
    if (week.campaigns && week.campaigns.length > 50 || (Date.now() - campaignStartTime) > 500) {
      console.log(`        Regular campaigns: ${week.campaigns ? week.campaigns.length : 0} campaigns in ${((Date.now() - campaignStartTime) / 1000).toFixed(1)}s`);
    }
  }
  
  const elapsed = (Date.now() - startTime) / 1000;
  if (elapsed > 1) {
    console.log(`      addUnifiedSubRows took ${elapsed.toFixed(1)}s`);
  }
}

function addCampaignRowsBatched(tableData, campaigns, week, weekKey, wow, formatData, appName = '', initialEROASCache = null) {
  if (CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    return;
  }
  
  const sortedCampaigns = campaigns.sort((a, b) => b.spend - a.spend);
  const batchSize = 100;
  
  for (let i = 0; i < sortedCampaigns.length; i += batchSize) {
    const batch = sortedCampaigns.slice(i, i + batchSize);
    
    batch.forEach(campaign => {
      let campaignIdValue;
      if (CURRENT_PROJECT === 'TRICKY' || CURRENT_PROJECT === 'REGULAR') {
        campaignIdValue = `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`;
      } else {
        campaignIdValue = campaign.campaignId;
      }
      
      const key = `${campaign.campaignId}_${weekKey}`;
      const campaignWoW = wow.campaignWoW[key] || {};
      
      const spendPct = campaignWoW.spendChangePercent !== undefined ? `${campaignWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitPct = campaignWoW.eProfitChangePercent !== undefined ? `${campaignWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const growthStatus = campaignWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'CAMPAIGN' });
      
      const campaignRow = createUnifiedRow('CAMPAIGN', week, campaign, spendPct, profitPct, growthStatus, appName, initialEROASCache, campaign.campaignId, campaign.sourceApp, campaignIdValue);
      tableData.push(campaignRow);
    });
  }
}

function createUnifiedRow(level, week, data, spendWoW, profitWoW, status, appName = '', initialEROASCache = null, identifier = '', displayName = '', campaignIdValue = '') {
  const headers = getUnifiedHeaders();
  const row = new Array(headers.length).fill('');
  
  row[0] = level;
  
  if (level === 'APP') {
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

function preloadAppsDbCache() {
  if (CURRENT_PROJECT !== 'TRICKY') return;
  
  if (TABLE_APPS_DB_CACHE.size > 0) return;
  
  try {
    const appsDb = new AppsDatabase('TRICKY');
    const cache = appsDb.loadFromCache();
    
    Object.keys(cache).forEach(bundleId => {
      const appInfo = cache[bundleId];
      let displayName = bundleId;
      
      if (appInfo && appInfo.publisher !== bundleId) {
        const publisher = appInfo.publisher || '';
        const appName = appInfo.appName || '';
        
        if (publisher && appName && publisher !== appName) {
          displayName = `${publisher} ${appName}`;
        } else if (publisher) {
          displayName = publisher;
        } else if (appName) {
          displayName = appName;
        }
      }
      
      TABLE_APPS_DB_CACHE.set(bundleId, {
        displayName: displayName,
        linkApp: appInfo ? appInfo.linkApp : ''
      });
    });
    
    console.log(`Apps Database preloaded: ${TABLE_APPS_DB_CACHE.size} entries`);
  } catch (e) {
    console.error('Error preloading Apps Database:', e);
  }
}

function getCachedDisplayName(bundleId, defaultName) {
  if (CURRENT_PROJECT !== 'TRICKY') {
    return { displayName: defaultName, isHyperlink: false };
  }
  
  if (TABLE_DISPLAY_NAME_CACHE.has(bundleId)) {
    return TABLE_DISPLAY_NAME_CACHE.get(bundleId);
  }
  
  let displayName = defaultName;
  let isHyperlink = false;
  const appsInfo = TABLE_APPS_DB_CACHE.get(bundleId);
  
  if (appsInfo && appsInfo.linkApp) {
    displayName = `=HYPERLINK("${appsInfo.linkApp}", "${defaultName}")`;
    isHyperlink = true;
  }
  
  const result = { displayName: displayName, isHyperlink: isHyperlink };
  TABLE_DISPLAY_NAME_CACHE.set(bundleId, result);
  return result;
}

function clearTableBuilderCaches() {
  WEEK_TOTALS_CACHE.clear();
  TABLE_APPS_DB_CACHE.clear();
  TABLE_DISPLAY_NAME_CACHE.clear();
}

function getUnifiedHeaders() {
  return [
    'Level', 'Week Range / Source App', 'ID', 'GEO',
    'Spend', 'Spend WoW %', 'Installs', 'CPI', 'ROAS D1→D3→D7→D30', 'IPM',
    'RR D-1', 'RR D-7', 'eARPU 365d', 'eROAS 365d', 'eROAS 730d', 'eProfit 730d', 'eProfit 730d WoW %', 'Growth Status', 'Comments'
  ];
}