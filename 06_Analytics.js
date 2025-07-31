function calculateWeekTotals(campaigns) {
  if (!campaigns || campaigns.length === 0) {
    return {
      spend: 0, cpi: 0, installs: 0, ipm: 0, rrD1: 0, roasD1: 0,
      roasD3: 0, rrD7: 0, roasD7: 0, roasD30: 0, eArpuForecast: 0,
      eRoasForecast: 0, eProfitForecast: 0, eRoasForecastD730: 0
    };
  }
  
  const totals = {
    spend: 0, cpi: 0, installs: 0, ipm: 0, rrD1: 0, roasD1: 0,
    roasD3: 0, rrD7: 0, roasD7: 0, roasD30: 0, eArpuForecast: 0,
    eRoasForecast: 0, eProfitForecast: 0, eRoasForecastD730: 0
  };
  
  let totalSpend = 0;
  let totalInstalls = 0;
  let weightedMetrics = {
    cpi: 0, ipm: 0, rrD1: 0, roasD1: 0, roasD3: 0, rrD7: 0, roasD7: 0,
    roasD30: 0, eArpuForecast: 0, eRoasForecast: 0, eRoasForecastD730: 0
  };
  
  campaigns.forEach(campaign => {
    const spend = parseFloat(campaign.spend) || 0;
    const installs = parseInt(campaign.installs) || 0;
    
    totalSpend += spend;
    totalInstalls += installs;
    totals.eProfitForecast += parseFloat(campaign.eProfitForecast) || 0;
    
    if (spend > 0) {
      Object.keys(weightedMetrics).forEach(metric => {
        const value = parseFloat(campaign[metric]) || 0;
        weightedMetrics[metric] += value * spend;
      });
    }
  });
  
  totals.spend = totalSpend;
  totals.installs = totalInstalls;
  
  if (totalSpend > 0) {
    Object.keys(weightedMetrics).forEach(metric => {
      totals[metric] = weightedMetrics[metric] / totalSpend;
    });
  }
  
  return totals;
}

function calculateWeekOverWeekAnalytics(appData) {
  const apps = {};
  const appWeekWoW = {};
  
  Object.keys(appData).forEach(appKey => {
    const app = appData[appKey];
    const weeks = Object.keys(app.weeks).sort();
    const appAnalytics = { appName: app.appName, totalSpend: 0, totalProfit: 0, weeks: [] };
    
    weeks.forEach((weekKey, index) => {
      const currentWeek = app.weeks[weekKey];
      let allCampaigns = [];
      
      if (CURRENT_PROJECT === 'TRICKY' && currentWeek.sourceApps) {
        Object.values(currentWeek.sourceApps).forEach(sourceApp => {
          allCampaigns.push(...sourceApp.campaigns);
        });
      } else if (CURRENT_PROJECT === 'OVERALL' && currentWeek.networks) {
        Object.values(currentWeek.networks).forEach(network => {
          allCampaigns.push(...network.campaigns);
        });
      } else {
        allCampaigns = currentWeek.campaigns || [];
      }
      
      const weekTotals = calculateWeekTotals(allCampaigns);
      appAnalytics.totalSpend += weekTotals.spend;
      appAnalytics.totalProfit += weekTotals.eProfitForecast;
      
      if (index > 0) {
        const prevWeekKey = weeks[index - 1];
        const prevWeek = app.weeks[prevWeekKey];
        
        let prevAllCampaigns = [];
        if (CURRENT_PROJECT === 'TRICKY' && prevWeek.sourceApps) {
          Object.values(prevWeek.sourceApps).forEach(sourceApp => {
            prevAllCampaigns.push(...sourceApp.campaigns);
          });
        } else if (CURRENT_PROJECT === 'OVERALL' && prevWeek.networks) {
          Object.values(prevWeek.networks).forEach(network => {
            prevAllCampaigns.push(...network.campaigns);
          });
        } else {
          prevAllCampaigns = prevWeek.campaigns || [];
        }
        
        const prevWeekTotals = calculateWeekTotals(prevAllCampaigns);
        
        const spendChange = weekTotals.spend - prevWeekTotals.spend;
        const spendChangePercent = prevWeekTotals.spend > 0 ? (spendChange / prevWeekTotals.spend) * 100 : 0;
        
        const profitChange = weekTotals.eProfitForecast - prevWeekTotals.eProfitForecast;
        const profitChangePercent = prevWeekTotals.eProfitForecast > 0 ? (profitChange / prevWeekTotals.eProfitForecast) * 100 : 0;
        
        let growthStatus = 'Stable';
        if (spendChangePercent > 10 && profitChangePercent > 5) {
          growthStatus = 'Growing';
        } else if (spendChangePercent < -10 || profitChangePercent < -15) {
          growthStatus = 'Declining';
        }
        
        const appWeekKey = `${app.appName}_${weekKey}`;
        appWeekWoW[appWeekKey] = {
          spendChange,
          spendChangePercent,
          eProfitChange: profitChange,
          eProfitChangePercent: profitChangePercent,
          growthStatus
        };
      }
      
      appAnalytics.weeks.push({
        weekKey,
        weekStart: currentWeek.weekStart,
        weekEnd: currentWeek.weekEnd,
        ...weekTotals
      });
    });
    
    apps[appKey] = appAnalytics;
  });
  
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    const networkData = {};
    
    Object.values(appData).forEach(network => {
      const networkKey = `${network.networkName}_${network.networkId}`;
      const weeks = Object.keys(network.weeks).sort();
      
      weeks.forEach((weekKey, index) => {
        const week = network.weeks[weekKey];
        
        if (index > 0) {
          const prevWeekKey = weeks[index - 1];
          const prevWeek = network.weeks[prevWeekKey];
          
          Object.values(week.apps).forEach(app => {
            const appWeekKey = `${app.appName}_${weekKey}`;
            
            let currentCampaigns = app.campaigns || [];
            let prevCampaigns = prevWeek.apps[app.appId] ? prevWeek.apps[app.appId].campaigns || [] : [];
            
            const currentTotals = calculateWeekTotals(currentCampaigns);
            const prevTotals = calculateWeekTotals(prevCampaigns);
            
            const spendChange = currentTotals.spend - prevTotals.spend;
            const spendChangePercent = prevTotals.spend > 0 ? (spendChange / prevTotals.spend) * 100 : 0;
            
            const profitChange = currentTotals.eProfitForecast - prevTotals.eProfitForecast;
            const profitChangePercent = prevTotals.eProfitForecast > 0 ? (profitChange / prevTotals.eProfitForecast) * 100 : 0;
            
            let growthStatus = 'Stable';
            if (spendChangePercent > 10 && profitChangePercent > 5) {
              growthStatus = 'Growing';
            } else if (spendChangePercent < -10 || profitChangePercent < -15) {
              growthStatus = 'Declining';
            }
            
            appWeekWoW[appWeekKey] = {
              networkId: network.networkId,
              networkName: network.networkName,
              spendChange,
              spendChangePercent,
              eProfitChange: profitChange,
              eProfitChangePercent: profitChangePercent,
              growthStatus
            };
          });
        }
      });
    });
    
    return { apps, appWeekWoW, networkData };
  }
  
  if (CURRENT_PROJECT === 'OVERALL') {
    const networkData = {};
    const campaignData = {};
    
    Object.values(appData).forEach(app => {
      const weeks = Object.keys(app.weeks).sort();
      
      weeks.forEach((weekKey, index) => {
        const week = app.weeks[weekKey];
        
        if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
          Object.values(week.sourceApps).forEach(sourceApp => {
            sourceApp.campaigns.forEach(c => {
              if (c.campaignId) {
                campaignData[`${c.campaignId}_${week.weekStart}`] = {
                  weekStart: week.weekStart,
                  appName: app.appName,
                  campaignId: c.campaignId,
                  campaignName: c.campaignName,
                  spend: c.spend,
                  eRoasForecastD730: c.eRoasForecastD730 || 0,
                  eProfitForecast: c.eProfitForecast,
                  installs: c.installs,
                  cpi: c.cpi,
                  roasD1: c.roasD1 || 0,
                  roasD3: c.roasD3 || 0,
                  roasD7: c.roasD7 || 0,
                  roasD30: c.roasD30 || 0,
                  ipm: c.ipm || 0,
                  eArpuForecast: c.eArpuForecast || 0,
                  rrD1: c.rrD1 || 0,
                  rrD7: c.rrD7 || 0
                };
              }
            });
          });
        } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
          Object.values(week.networks).forEach(network => {
            allCampaigns.push(...network.campaigns);
            
            const networkKey = network.networkId;
            const networkSpend = network.campaigns.reduce((s, c) => s + c.spend, 0);
            const networkProfit = network.campaigns.reduce((s, c) => s + c.eProfitForecast, 0);
            
            if (!networkData[networkKey]) {
              networkData[networkKey] = {};
            }
            
            networkData[networkKey][week.weekStart] = {
              weekStart: week.weekStart,
              spend: networkSpend,
              profit: networkProfit,
              networkId: network.networkId,
              networkName: network.networkName
            };
          });
        } else if (CURRENT_PROJECT === 'OVERALL' && week.campaigns) {
          allCampaigns = week.campaigns || [];
        } else if (CURRENT_PROJECT === 'OVERALL') {
          allCampaigns = week.campaigns || [];
        } else {
          allCampaigns = week.campaigns || [];
          
          allCampaigns.forEach(c => {
            if (c.campaignId) {
              campaignData[`${c.campaignId}_${week.weekStart}`] = {
                weekStart: week.weekStart,
                appName: app.appName,
                campaignId: c.campaignId,
                campaignName: c.campaignName,
                spend: c.spend,
                eRoasForecastD730: c.eRoasForecastD730 || 0,
                eProfitForecast: c.eProfitForecast,
                installs: c.installs,
                cpi: c.cpi,
                roasD1: c.roasD1 || 0,
                roasD3: c.roasD3 || 0,
                roasD7: c.roasD7 || 0,
                roasD30: c.roasD30 || 0,
                ipm: c.ipm || 0,
                eArpuForecast: c.eArpuForecast || 0,
                rrD1: c.rrD1 || 0,
                rrD7: c.rrD7 || 0
              };
            }
          });
        }
      });
   });
    
    return { apps, appWeekWoW, networkData, campaignData };
  }
  
  return { apps, appWeekWoW };
}

let WEEK_TOTALS_CACHE = new Map();

function getCachedWeekTotals(campaigns) {
  if (!campaigns || campaigns.length === 0) {
    return {
      spend: 0, cpi: 0, installs: 0, ipm: 0, rrD1: 0, roasD1: 0,
      roasD3: 0, rrD7: 0, roasD7: 0, roasD30: 0, eArpuForecast: 0,
      eRoasForecast: 0, eProfitForecast: 0, eRoasForecastD730: 0
    };
  }
  
  const cacheKey = campaigns.map(c => `${c.campaignId}_${c.spend}_${c.installs}`).join('|');
  
  if (WEEK_TOTALS_CACHE.has(cacheKey)) {
    return WEEK_TOTALS_CACHE.get(cacheKey);
  }
  
  const totals = {
    spend: 0, cpi: 0, installs: 0, ipm: 0, rrD1: 0, roasD1: 0,
    roasD3: 0, rrD7: 0, roasD7: 0, roasD30: 0, eArpuForecast: 0,
    eRoasForecast: 0, eProfitForecast: 0, eRoasForecastD730: 0
  };
  
  let totalSpend = 0;
  let totalInstalls = 0;
  let weightedMetrics = {
    cpi: 0, ipm: 0, rrD1: 0, roasD1: 0, roasD3: 0, rrD7: 0, roasD7: 0,
    roasD30: 0, eArpuForecast: 0, eRoasForecast: 0, eRoasForecastD730: 0
  };
  
  campaigns.forEach(campaign => {
    const spend = parseFloat(campaign.spend) || 0;
    const installs = parseInt(campaign.installs) || 0;
    
    totalSpend += spend;
    totalInstalls += installs;
    totals.eProfitForecast += parseFloat(campaign.eProfitForecast) || 0;
    
    if (spend > 0) {
      Object.keys(weightedMetrics).forEach(metric => {
        const value = parseFloat(campaign[metric]) || 0;
        weightedMetrics[metric] += value * spend;
      });
    }
  });
  
  totals.spend = totalSpend;
  totals.installs = totalInstalls;
  
  if (totalSpend > 0) {
    Object.keys(weightedMetrics).forEach(metric => {
      totals[metric] = weightedMetrics[metric] / totalSpend;
    });
  }
  
  WEEK_TOTALS_CACHE.set(cacheKey, totals);
  return totals;
}

function updateProjectData(projectName) {
  const startTime = Date.now();
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return;
  }
  
  try {
    const cache = new CommentCache(projectName);
    cache.syncCommentsFromSheet();
  } catch (e) {
    console.error(`${projectName}: Failed to save comments:`, e);
  }
  
  let earliestDate = null;
  const range = `${config.SHEET_NAME}!A:B`;
  const response = Sheets.Spreadsheets.Values.get(config.SHEET_ID, range);
  const data = response.values || [];

  for (let i = 1; i < data.length; i++) {
    if (data[i] && data[i][0] === 'WEEK') {
      const weekRange = data[i][1];
      if (weekRange) {
        const [startStr] = weekRange.split(' - ');
        const startDate = new Date(startStr);
        if (!earliestDate || startDate < earliestDate) earliestDate = startDate;
      }
    }
  }
  
  if (!earliestDate) {
    return;
  }
  
  const today = new Date();
  const dayOfWeek = today.getDay();
  let endDate = new Date(today);
  
  if (dayOfWeek === 0) {
    endDate.setDate(today.getDate() - 1);
  } else {
    endDate.setDate(today.getDate() - dayOfWeek);
  }
  
  const dateRange = {
    from: formatDateForAPI(earliestDate),
    to: formatDateForAPI(endDate)
  };
  
  const raw = fetchProjectCampaignData(projectName, dateRange);
  
  if (!raw.data?.analytics?.richStats?.stats?.length) {
    return;
  }
  
  const processed = processProjectApiData(projectName, raw);
  
  if (Object.keys(processed).length === 0) {
    return;
  }
  
  clearProjectDataSilent(projectName);
  
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  let rowCount = 0;
  const recordCount = raw.data.analytics.richStats.stats.length;
  
  try {
    if (projectName === 'OVERALL') {
      rowCount = createOverallPivotTable(processed);
    } else if (projectName === 'INCENT_TRAFFIC') {
      rowCount = createIncentTrafficPivotTable(processed);
    } else {
      rowCount = createEnhancedPivotTable(processed);
    }
    const cache = new CommentCache(projectName);
    cache.applyCommentsToSheet();
  } finally {
    setCurrentProject(originalProject);
  }
  
  const totalTime = Date.now() - startTime;
  logInfo(projectName, recordCount, rowCount, totalTime);
}

function updateAllDataToCurrent() {
  const overallStartTime = Date.now();
  const ui = SpreadsheetApp.getUi();
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('No existing data found. Please create a report first.');
    return;
  }
  
  try {
    let earliestDate = null;

    const range = `${config.SHEET_NAME}!A:B`;
    const response = Sheets.Spreadsheets.Values.get(config.SHEET_ID, range);
    const data = response.values || [];

    for (let i = 1; i < data.length; i++) {
      if (data[i] && data[i][0] === 'WEEK') {
        const weekRange = data[i][1];
        if (weekRange) {
          const [startStr] = weekRange.split(' - ');
          const startDate = new Date(startStr);
          if (!earliestDate || startDate < earliestDate) earliestDate = startDate;
        }
      }
    }
    
    if (!earliestDate) {
      ui.alert('No week data found in the sheet.');
      return;
    }
    
    const today = new Date();
    const dayOfWeek = today.getDay();
    let endDate = new Date(today);
    if (dayOfWeek === 0) {
      endDate.setDate(today.getDate() - 1);
    } else {
      endDate.setDate(today.getDate() - dayOfWeek);
    }
    
    const dateRange = { from: formatDateForAPI(earliestDate), to: formatDateForAPI(endDate) };
    const raw = fetchCampaignData(dateRange);
    
    if (!raw.data?.analytics?.richStats?.stats?.length) {
      ui.alert('No data found for the date range.');
      return;
    }
    
    const processed = processApiData(raw);
    
    if (Object.keys(processed).length === 0) {
      ui.alert('No valid data to process.');
      return;
    }
    
    clearAllDataSilent();
    
    let rowCount = 0;
    const recordCount = raw.data.analytics.richStats.stats.length;
    
    if (CURRENT_PROJECT === 'OVERALL') {
      rowCount = createOverallPivotTable(processed);
    } else if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
      rowCount = createIncentTrafficPivotTable(processed);
    } else {
      rowCount = createEnhancedPivotTable(processed);
    }
    
    const cache = new CommentCache();
    cache.applyCommentsToSheet();
    
    const totalTime = Date.now() - overallStartTime;
    logInfo(CURRENT_PROJECT, recordCount, rowCount, totalTime);
    
    ui.alert('Success', `Successfully updated all data from ${dateRange.from} to ${dateRange.to}!`, ui.ButtonSet.OK);
  } catch (e) {
    console.error('Error updating data:', e);
    ui.alert('Error', 'Error updating data: ' + e.toString(), ui.ButtonSet.OK);
  }
}