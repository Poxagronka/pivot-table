function buildUnifiedTable(data, tableData, formatData, wow, initialEROASCache) {
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    const networkKeys = Object.keys(data).sort((a, b) => 
      data[a].networkName.localeCompare(data[b].networkName)
    );
    
    networkKeys.forEach(networkKey => {
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
        
        const weekTotals = calculateWeekTotals(allCampaigns);
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
          const appTotals = calculateWeekTotals(app.campaigns);
          
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
    const appKeys = Object.keys(data).sort((a, b) => data[a].appName.localeCompare(data[b].appName));
    
    appKeys.forEach(appKey => {
      const app = data[appKey];
      
      formatData.push({ row: tableData.length + 1, type: 'APP' });
      const emptyRow = new Array(getUnifiedHeaders().length).fill('');
      emptyRow[0] = 'APP';
      emptyRow[1] = app.appName;
      tableData.push(emptyRow);

      const weekKeys = Object.keys(app.weeks).sort();
      weekKeys.forEach(weekKey => {
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
        
        const weekTotals = calculateWeekTotals(allCampaigns);
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

function addUnifiedSubRows(tableData, week, weekKey, wow, formatData, appName, initialEROASCache) {
  if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
    const sourceAppKeys = Object.keys(week.sourceApps).sort((a, b) => {
      const totalSpendA = week.sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
      const totalSpendB = week.sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
      return totalSpendB - totalSpendA;
    });
    
    sourceAppKeys.forEach(sourceAppKey => {
      const sourceApp = week.sourceApps[sourceAppKey];
      const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
      
      const sourceAppWoWKey = `${sourceApp.sourceAppId}_${weekKey}`;
      const sourceAppWoW = wow.sourceAppWoW[sourceAppWoWKey] || {};
      
      const spendWoW = sourceAppWoW.spendChangePercent !== undefined ? `${sourceAppWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = sourceAppWoW.eProfitChangePercent !== undefined ? `${sourceAppWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = sourceAppWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'SOURCE_APP' });
      
      let sourceAppDisplayName = sourceApp.sourceAppName;
      if (CURRENT_PROJECT === 'TRICKY') {
        try {
          const appsDb = new AppsDatabase('TRICKY');
          const cache = appsDb.loadFromCache();
          const appInfo = cache[sourceApp.sourceAppId];
          if (appInfo && appInfo.linkApp) {
            sourceAppDisplayName = `=HYPERLINK("${appInfo.linkApp}", "${sourceApp.sourceAppName}")`;
            formatData.push({ row: tableData.length + 1, type: 'HYPERLINK' });
          }
        } catch (e) {
          console.log('Error getting store link for source app:', e);
        }
      }
      
      const sourceAppRow = createUnifiedRow('SOURCE_APP', week, sourceAppTotals, spendWoW, profitWoW, status, appName, initialEROASCache, sourceApp.sourceAppId, sourceAppDisplayName);
      tableData.push(sourceAppRow);
      
      addCampaignRows(tableData, sourceApp.campaigns, { weekStart: weekKey.split('-').join('/'), weekEnd: '' }, weekKey, wow, formatData, appName, initialEROASCache);
    });
  } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
    const networkKeys = Object.keys(week.networks).sort((a, b) => {
      const totalSpendA = week.networks[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
      const totalSpendB = week.networks[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
      return totalSpendB - totalSpendA;
    });
    
    networkKeys.forEach(networkKey => {
      const network = week.networks[networkKey];
      const networkTotals = calculateWeekTotals(network.campaigns);
      
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
    addCampaignRows(tableData, week.campaigns, week, weekKey, wow, formatData, appName, initialEROASCache);
  }
}

function addCampaignRows(tableData, campaigns, week, weekKey, wow, formatData, appName = '', initialEROASCache = null) {
  if (CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    return;
  }
  
  campaigns.sort((a, b) => b.spend - a.spend).forEach(campaign => {
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

function createUnifiedRow(level, week, data, spendWoW, profitWoW, status, appName = '', initialEROASCache = null, identifier = '', displayName = '', campaignIdValue = '') {
  const headers = getUnifiedHeaders();
  const row = new Array(headers.length).fill('');
  
  row[0] = level;
  
  if (level === 'WEEK') {
    row[1] = `${week.weekStart} - ${week.weekEnd}`;
    const combinedRoas = `${data.avgRoasD1.toFixed(0)}% → ${data.avgRoasD3.toFixed(0)}% → ${data.avgRoasD7.toFixed(0)}% → ${data.avgRoasD30.toFixed(0)}%`;
    
    let eROAS730Display = `${data.avgEROASD730.toFixed(0)}%`;
    if (initialEROASCache && appName) {
      const weekRange = `${week.weekStart} - ${week.weekEnd}`;
      eROAS730Display = initialEROASCache.formatEROASWithInitial('WEEK', appName, weekRange, data.avgEROASD730);
    }
    
    row[4] = data.totalSpend.toFixed(2); row[5] = spendWoW; row[6] = data.totalInstalls; row[7] = data.avgCpi.toFixed(3);
    row[8] = combinedRoas; row[9] = data.avgIpm.toFixed(1); row[10] = `${data.avgRrD1.toFixed(0)}%`; row[11] = `${data.avgRrD7.toFixed(0)}%`;
    row[12] = data.avgArpu.toFixed(3); row[13] = `${data.avgERoas.toFixed(0)}%`; row[14] = eROAS730Display;
    row[15] = data.totalProfit.toFixed(2); row[16] = profitWoW; row[17] = status;
  } else if (level === 'CAMPAIGN') {
    row[1] = data.sourceApp; row[2] = campaignIdValue; row[3] = data.geo;
    const combinedRoas = `${data.roasD1.toFixed(0)}% → ${data.roasD3.toFixed(0)}% → ${data.roasD7.toFixed(0)}% → ${data.roasD30.toFixed(0)}%`;
    
    let eROAS730Display = `${data.eRoasForecastD730.toFixed(0)}%`;
    if (initialEROASCache && appName) {
      const weekRange = `${week.weekStart} - ${week.weekEnd}`;
      eROAS730Display = initialEROASCache.formatEROASWithInitial('CAMPAIGN', appName, weekRange, data.eRoasForecastD730, data.campaignId, data.sourceApp);
    }
    
    row[4] = data.spend.toFixed(2); row[5] = spendWoW; row[6] = data.installs; row[7] = data.cpi ? data.cpi.toFixed(3) : '0.000';
    row[8] = combinedRoas; row[9] = data.ipm.toFixed(1); row[10] = `${data.rrD1.toFixed(0)}%`; row[11] = `${data.rrD7.toFixed(0)}%`;
    row[12] = data.eArpuForecast.toFixed(3); row[13] = `${data.eRoasForecast.toFixed(0)}%`; row[14] = eROAS730Display;
    row[15] = data.eProfitForecast.toFixed(2); row[16] = profitWoW; row[17] = status;
  } else {
    row[1] = displayName || identifier;
    const combinedRoas = `${data.avgRoasD1.toFixed(0)}% → ${data.avgRoasD3.toFixed(0)}% → ${data.avgRoasD7.toFixed(0)}% → ${data.avgRoasD30.toFixed(0)}%`;
    
    let eROAS730Display = `${data.avgEROASD730.toFixed(0)}%`;
    if (initialEROASCache && appName) {
      const weekRange = `${week.weekStart} - ${week.weekEnd}`;
      eROAS730Display = initialEROASCache.formatEROASWithInitial(level, appName, weekRange, data.avgEROASD730, identifier, displayName);
    }
    
    row[4] = data.totalSpend.toFixed(2); row[5] = spendWoW; row[6] = data.totalInstalls; row[7] = data.avgCpi.toFixed(3);
    row[8] = combinedRoas; row[9] = data.avgIpm.toFixed(1); row[10] = `${data.avgRrD1.toFixed(0)}%`; row[11] = `${data.avgRrD7.toFixed(0)}%`;
    row[12] = data.avgArpu.toFixed(3); row[13] = `${data.avgERoas.toFixed(0)}%`; row[14] = eROAS730Display;
    row[15] = data.totalProfit.toFixed(2); row[16] = profitWoW; row[17] = status;
  }
  
  row[18] = '';
  return row;
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

function getUnifiedHeaders() {
  return [
    'Level', 'Week Range / Source App', 'ID', 'GEO',
    'Spend', 'Spend WoW %', 'Installs', 'CPI', 'ROAS D1→D3→D7→D30', 'IPM',
    'RR D-1', 'RR D-7', 'eARPU 365d', 'eROAS 365d', 'eROAS 730d', 'eProfit 730d', 'eProfit 730d WoW %', 'Growth Status', 'Comments'
  ];
}