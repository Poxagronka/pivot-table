const WEEK_TOTALS_CACHE = new Map();
const TABLE_APPS_DB_CACHE = new Map();
const TABLE_DISPLAY_NAME_CACHE = new Map();

function buildUnifiedTable(data, tableData, formatData, wow, initialEROASCache) {
  clearTableBuilderCaches();
  
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    const networkKeys = Object.keys(data).sort((a, b) => 
      data[a].networkName.localeCompare(data[b].networkName)
    );
    
    networkKeys.forEach((networkKey, networkIndex) => {
      const network = data[networkKey];
      
      formatData.push({ row: tableData.length + 1, type: 'NETWORK' });
      const emptyRow = new Array(getUnifiedHeaders().length).fill('');
      emptyRow[0] = 'NETWORK';
      emptyRow[1] = network.networkName;
      tableData.push(emptyRow);
      
      const weekKeys = Object.keys(network.weeks).sort();
      
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
    
  } else {
    preloadAppsDbCache();
    
    const appKeys = Object.keys(data).sort((a, b) => data[a].appName.localeCompare(data[b].appName));
    
    appKeys.forEach((appKey, appIndex) => {
      const app = data[appKey];
      
      formatData.push({ row: tableData.length + 1, type: 'APP' });
      const emptyRow = new Array(getUnifiedHeaders().length).fill('');
      emptyRow[0] = 'APP';
      emptyRow[1] = app.appName;
      tableData.push(emptyRow);

      const weekKeys = Object.keys(app.weeks).sort();
      
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
        
        addUnifiedSubRows(tableData, week, weekKey, wow, formatData, app.appName, initialEROASCache);
      });
    });
  }
}

function addUnifiedSubRows(tableData, week, weekKey, wow, formatData, appName = '', initialEROASCache = null) {
  if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
    const sourceAppKeys = Object.keys(week.sourceApps).sort((a, b) => {
      const totalSpendA = week.sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
      const totalSpendB = week.sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
      return totalSpendB - totalSpendA;
    });
    
    sourceAppKeys.forEach(sourceAppKey => {
      const sourceApp = week.sourceApps[sourceAppKey];
      const sourceAppTotals = getCachedWeekTotals(sourceApp.campaigns);
      
      const sourceAppWoWKey = `${sourceApp.sourceAppId}_${weekKey}`;
      const sourceAppWoW = wow.sourceAppWoW[sourceAppWoWKey] || {};
      
      const spendWoW = sourceAppWoW.spendChangePercent !== undefined ? `${sourceAppWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = sourceAppWoW.eProfitChangePercent !== undefined ? `${sourceAppWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = sourceAppWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'SOURCE_APP' });
      
      const cachedName = getCachedDisplayName(sourceApp.sourceAppId, sourceApp.sourceAppName);
      if (cachedName.isHyperlink) {
        formatData.push({ row: tableData.length + 1, type: 'HYPERLINK' });
      }
      
      const sourceAppRow = createUnifiedRow('SOURCE_APP', week, sourceAppTotals, spendWoW, profitWoW, status, appName, initialEROASCache, sourceApp.sourceAppId, cachedName.displayName);
      tableData.push(sourceAppRow);
      
      addCampaignRowsBatched(tableData, sourceApp.campaigns, week, weekKey, wow, formatData, appName, initialEROASCache);
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
      
      const networkWoWKey = `${network.networkId}_${weekKey}`;
      const networkWoW = wow.networkWoW[networkWoWKey] || {};
      
      const spendWoW = networkWoW.spendChangePercent !== undefined ? `${networkWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = networkWoW.eProfitChangePercent !== undefined ? `${networkWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = networkWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'NETWORK' });
      
      const networkRow = createUnifiedRow('NETWORK', week, networkTotals, spendWoW, profitWoW, status, appName, initialEROASCache, network.networkId, network.networkName);
      tableData.push(networkRow);
    });
  } else if (week.campaigns) {
    addCampaignRowsBatched(tableData, week.campaigns, week, weekKey, wow, formatData, appName, initialEROASCache);
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
      const status = campaignWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'CAMPAIGN' });
      if (CURRENT_PROJECT === 'TRICKY' || CURRENT_PROJECT === 'REGULAR') {
        formatData.push({ row: tableData.length + 1, type: 'HYPERLINK' });
      }
      
      const campaignRow = createUnifiedRow('CAMPAIGN', week, campaign, spendPct, profitPct, status, appName, initialEROASCache, campaignIdValue, campaign.campaignName);
      tableData.push(campaignRow);
    });
  }
}

function preloadAppsDbCache() {
  if (CURRENT_PROJECT !== 'TRICKY') return;
  
  try {
    const appsDb = new AppsDatabase('TRICKY');
    const cache = appsDb.loadFromCache();
    
    Object.keys(cache).forEach(bundleId => {
      const appInfo = cache[bundleId];
      TABLE_APPS_DB_CACHE.set(bundleId, {
        publisher: appInfo.publisher || bundleId,
        appName: appInfo.appName || '',
        linkApp: appInfo.linkApp || ''
      });
    });
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

function createUnifiedRow(level, weekData, metrics, spendWoW, profitWoW, status, contextName, initialEROASCache, id = '', name = '') {
  const weekRange = weekData.weekStart && weekData.weekEnd ? 
    getFormattedDateRange(weekData.weekStart, weekData.weekEnd) : '';
  
  let displayName = name || '';
  let geoValue = '';
  
  if (level === 'CAMPAIGN' && metrics.campaignName) {
    geoValue = extractGeoFromCampaign(metrics.campaignName);
  }
  
  let eroasDisplay = '';
  let eprofitDisplay = '';
  
  if (initialEROASCache && metrics.eRoasForecastD730 > 0) {
    if (level === 'WEEK') {
      eroasDisplay = initialEROASCache.formatEROASWithInitial('WEEK', contextName, weekRange, metrics.eRoasForecastD730);
      eprofitDisplay = initialEROASCache.formatProfitWithInitial('WEEK', contextName, weekRange, metrics.eProfitForecast);
    } else if (level === 'SOURCE_APP') {
      eroasDisplay = initialEROASCache.formatEROASWithInitial('SOURCE_APP', contextName, weekRange, metrics.eRoasForecastD730, id);
      eprofitDisplay = initialEROASCache.formatProfitWithInitial('SOURCE_APP', contextName, weekRange, metrics.eProfitForecast, id);
    } else if (level === 'CAMPAIGN') {
      eroasDisplay = initialEROASCache.formatEROASWithInitial('CAMPAIGN', contextName, weekRange, metrics.eRoasForecastD730, metrics.campaignId, id);
      eprofitDisplay = initialEROASCache.formatProfitWithInitial('CAMPAIGN', contextName, weekRange, metrics.eProfitForecast, metrics.campaignId, id);
    } else {
      eroasDisplay = `${Math.round(metrics.eRoasForecastD730)}%`;
      eprofitDisplay = `${Math.round(metrics.eProfitForecast)}$`;
    }
  } else {
    eroasDisplay = metrics.eRoasForecastD730 > 0 ? `${Math.round(metrics.eRoasForecastD730)}%` : '';
    eprofitDisplay = metrics.eProfitForecast > 0 ? `${Math.round(metrics.eProfitForecast)}$` : '';
  }
  
  const roasSequence = [
    metrics.roasD1 > 0 ? Math.round(metrics.roasD1) + '%' : '',
    metrics.roasD3 > 0 ? Math.round(metrics.roasD3) + '%' : '',
    metrics.roasD7 > 0 ? Math.round(metrics.roasD7) + '%' : '',
    metrics.roasD30 > 0 ? Math.round(metrics.roasD30) + '%' : ''
  ].filter(r => r).join('→');
  
  return [
    level,
    level === 'WEEK' ? weekRange : displayName,
    id,
    geoValue,
    metrics.spend > 0 ? Math.round(metrics.spend) : '',
    spendWoW,
    metrics.installs > 0 ? Math.round(metrics.installs) : '',
    metrics.cpi > 0 ? Math.round(metrics.cpi * 100) / 100 : '',
    roasSequence,
    metrics.ipm > 0 ? Math.round(metrics.ipm * 10) / 10 : '',
    metrics.rrD1 > 0 ? Math.round(metrics.rrD1 * 100) / 100 : '',
    metrics.rrD7 > 0 ? Math.round(metrics.rrD7 * 100) / 100 : '',
    metrics.eArpuForecast > 0 ? Math.round(metrics.eArpuForecast * 100) / 100 : '',
    metrics.eRoasForecast > 0 ? Math.round(metrics.eRoasForecast) + '%' : '',
    eroasDisplay,
    eprofitDisplay,
    profitWoW,
    status,
    ''
  ];
}

function createUnifiedRowGrouping(sheet, tableData, data) {
  if (tableData.length <= 2) return;
  
  try {
    const groupRanges = [];
    let currentGroupStart = null;
    let currentGroupType = null;
    
    for (let i = 1; i < tableData.length; i++) {
      const row = tableData[i];
      const level = row[0];
      
      if (level === 'APP' || level === 'NETWORK') {
        if (currentGroupStart !== null && currentGroupType !== level) {
          if (i - currentGroupStart > 1) {
            groupRanges.push({ start: currentGroupStart + 1, end: i, type: currentGroupType });
          }
        }
        currentGroupStart = i;
        currentGroupType = level;
      }
    }
    
    if (currentGroupStart !== null && tableData.length - currentGroupStart > 1) {
      groupRanges.push({ start: currentGroupStart + 1, end: tableData.length, type: currentGroupType });
    }
    
    groupRanges.forEach(range => {
      try {
        if (range.end - range.start > 1) {
          const group = sheet.getRange(range.start + 1, 1, range.end - range.start - 1, 1);
          group.shiftRowGroupDepth(1);
        }
      } catch (e) {
        console.error(`Error creating group for rows ${range.start}-${range.end}:`, e);
      }
    });
    
  } catch (e) {
    console.error('Error in createUnifiedRowGrouping:', e);
  }
}