function calculateWoWMetrics(appData) {
  if (!appData || typeof appData !== 'object') {
    return { campaignWoW: {}, appWeekWoW: {}, sourceAppWoW: {}, networkWoW: {} };
  }

  if (CURRENT_PROJECT === 'OVERALL') {
    return calculateOverallWoWMetricsOptimized(appData);
  }
  
  const campaignData = {};
  const appWeekData = {};
  const sourceAppData = {};

  Object.values(appData).forEach(app => {
    appWeekData[app.appName] = {};
    
    Object.values(app.weeks).forEach(week => {
      let allCampaigns = [];
      
      if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
        Object.values(week.sourceApps).forEach(sourceApp => {
          allCampaigns.push(...sourceApp.campaigns);
          
          const sourceAppKey = sourceApp.sourceAppId;
          const sourceAppSpend = sourceApp.campaigns.reduce((s, c) => s + c.spend, 0);
          const sourceAppProfit = sourceApp.campaigns.reduce((s, c) => s + c.eProfitForecast, 0);
          
          if (!sourceAppData[sourceAppKey]) {
            sourceAppData[sourceAppKey] = {};
          }
          
          sourceAppData[sourceAppKey][week.weekStart] = {
            weekStart: week.weekStart,
            spend: sourceAppSpend,
            profit: sourceAppProfit,
            sourceAppId: sourceApp.sourceAppId,
            sourceAppName: sourceApp.sourceAppName
          };
          
          sourceApp.campaigns.forEach(c => {
            if (c.campaignId) {
              campaignData[`${c.campaignId}_${week.weekStart}`] = {
                campaignId: c.campaignId,
                weekStart: week.weekStart,
                spend: c.spend,
                eProfitForecast: c.eProfitForecast
              };
            }
          });
        });
      } else {
        allCampaigns = week.campaigns || [];
        
        allCampaigns.forEach(c => {
          if (c.campaignId) {
            campaignData[`${c.campaignId}_${week.weekStart}`] = {
              campaignId: c.campaignId,
              weekStart: week.weekStart,
              spend: c.spend,
              eProfitForecast: c.eProfitForecast
            };
          }
        });
      }
      
      const spend = allCampaigns.reduce((s, c) => s + c.spend, 0);
      const profit = allCampaigns.reduce((s, c) => s + c.eProfitForecast, 0);
      appWeekData[app.appName][week.weekStart] = { weekStart: week.weekStart, spend, profit };
    });
  });

  const campaignWoW = {};
  const campaigns = {};
  Object.values(campaignData).forEach(d => {
    if (!campaigns[d.campaignId]) campaigns[d.campaignId] = [];
    campaigns[d.campaignId].push(d);
  });

  Object.keys(campaigns).forEach(campaignId => {
    campaigns[campaignId].sort((a, b) => new Date(a.weekStart) - new Date(b.weekStart));
    campaigns[campaignId].forEach((curr, i) => {
      const key = `${campaignId}_${curr.weekStart}`;
      campaignWoW[key] = { spendChangePercent: 0, eProfitChangePercent: 0, growthStatus: 'First Week' };
      
      if (i > 0) {
        const prev = campaigns[campaignId][i - 1];
        const spendPct = prev.spend ? ((curr.spend - prev.spend) / Math.abs(prev.spend)) * 100 : 0;
        const profitPct = prev.eProfitForecast ? ((curr.eProfitForecast - prev.eProfitForecast) / Math.abs(prev.eProfitForecast)) * 100 : 0;
        campaignWoW[key] = { 
          spendChangePercent: spendPct, 
          eProfitChangePercent: profitPct, 
          growthStatus: calculateGrowthStatusFast(prev, curr, spendPct, profitPct) 
        };
      }
    });
  });

  const appWeekWoW = {};
  Object.keys(appWeekData).forEach(appName => {
    const weeks = Object.values(appWeekData[appName]).sort((a, b) => new Date(a.weekStart) - new Date(b.weekStart));
    weeks.forEach((curr, i) => {
      const key = `${appName}_${curr.weekStart}`;
      appWeekWoW[key] = { spendChangePercent: 0, eProfitChangePercent: 0, growthStatus: 'First Week' };
      
      if (i > 0) {
        const prev = weeks[i - 1];
        const spendPct = prev.spend ? ((curr.spend - prev.spend) / Math.abs(prev.spend)) * 100 : 0;
        const profitPct = prev.profit ? ((curr.profit - prev.profit) / Math.abs(prev.profit)) * 100 : 0;
        appWeekWoW[key] = { 
          spendChangePercent: spendPct, 
          eProfitChangePercent: profitPct, 
          growthStatus: calculateGrowthStatusFast(prev, curr, spendPct, profitPct, 'profit') 
        };
      }
    });
  });

  const sourceAppWoW = {};
  if (CURRENT_PROJECT === 'TRICKY') {
    Object.keys(sourceAppData).forEach(bundleId => {
      const weeks = Object.values(sourceAppData[bundleId]).sort((a, b) => new Date(a.weekStart) - new Date(b.weekStart));
      weeks.forEach((curr, i) => {
        const key = `${bundleId}_${curr.weekStart}`;
        sourceAppWoW[key] = { spendChangePercent: 0, eProfitChangePercent: 0, growthStatus: 'First Week' };
        
        if (i > 0) {
          const prev = weeks[i - 1];
          const spendPct = prev.spend ? ((curr.spend - prev.spend) / Math.abs(prev.spend)) * 100 : 0;
          const profitPct = prev.profit ? ((curr.profit - prev.profit) / Math.abs(prev.profit)) * 100 : 0;
          sourceAppWoW[key] = { 
            spendChangePercent: spendPct, 
            eProfitChangePercent: profitPct, 
            growthStatus: calculateGrowthStatusFast(prev, curr, spendPct, profitPct, 'profit') 
          };
        }
      });
    });
  }

  return { campaignWoW, appWeekWoW, sourceAppWoW, networkWoW: {} };
}

function calculateOverallWoWMetricsOptimized(appData) {
  const appWeekData = {};
  const networkData = {};

  Object.values(appData).forEach(app => {
    appWeekData[app.appName] = {};
    
    Object.values(app.weeks).forEach(week => {
      let weekSpend = 0;
      let weekProfit = 0;
      
      if (week.networks) {
        Object.values(week.networks).forEach(network => {
          weekSpend += network.spend || 0;
          weekProfit += network.eProfitForecast || 0;
          
          if (!networkData[network.networkId]) {
            networkData[network.networkId] = {};
          }
          
          networkData[network.networkId][week.weekStart] = {
            weekStart: week.weekStart,
            spend: network.spend,
            profit: network.eProfitForecast
          };
        });
      }
      
      appWeekData[app.appName][week.weekStart] = { 
        weekStart: week.weekStart, 
        spend: weekSpend, 
        profit: weekProfit 
      };
    });
  });

  const appWeekWoW = {};
  Object.keys(appWeekData).forEach(appName => {
    const weeks = Object.values(appWeekData[appName]).sort((a, b) => new Date(a.weekStart) - new Date(b.weekStart));
    weeks.forEach((curr, i) => {
      const key = `${appName}_${curr.weekStart}`;
      appWeekWoW[key] = { spendChangePercent: 0, eProfitChangePercent: 0, growthStatus: 'First Week' };
      
      if (i > 0) {
        const prev = weeks[i - 1];
        const spendPct = prev.spend ? ((curr.spend - prev.spend) / Math.abs(prev.spend)) * 100 : 0;
        const profitPct = prev.profit ? ((curr.profit - prev.profit) / Math.abs(prev.profit)) * 100 : 0;
        appWeekWoW[key] = { 
          spendChangePercent: spendPct, 
          eProfitChangePercent: profitPct, 
          growthStatus: calculateGrowthStatusFast(prev, curr, spendPct, profitPct, 'profit') 
        };
      }
    });
  });

  const networkWoW = {};
  Object.keys(networkData).forEach(networkId => {
    const weeks = Object.values(networkData[networkId]).sort((a, b) => new Date(a.weekStart) - new Date(b.weekStart));
    weeks.forEach((curr, i) => {
      const key = `${networkId}_${curr.weekStart}`;
      networkWoW[key] = { spendChangePercent: 0, eProfitChangePercent: 0, growthStatus: 'First Week' };
      
      if (i > 0) {
        const prev = weeks[i - 1];
        const spendPct = prev.spend ? ((curr.spend - prev.spend) / Math.abs(prev.spend)) * 100 : 0;
        const profitPct = prev.profit ? ((curr.profit - prev.profit) / Math.abs(prev.profit)) * 100 : 0;
        networkWoW[key] = { 
          spendChangePercent: spendPct, 
          eProfitChangePercent: profitPct, 
          growthStatus: calculateGrowthStatusFast(prev, curr, spendPct, profitPct, 'profit') 
        };
      }
    });
  });

  return { campaignWoW: {}, appWeekWoW, sourceAppWoW: {}, networkWoW };
}

function calculateGrowthStatusFast(prev, curr, spendPct, profitPct, profitField = 'eProfitForecast') {
  const prevProfit = profitField === 'profit' ? prev.profit : prev.eProfitForecast;
  const currProfit = profitField === 'profit' ? curr.profit : curr.eProfitForecast;
  
  if (prevProfit < 0 && currProfit > 0) return '🟢 Healthy Growth';
  if (prevProfit > 0 && currProfit < 0) return '🔴 Inefficient Growth';
  if (profitPct <= -8) return '🔴 Inefficient Growth';
  if (spendPct >= 10 && profitPct >= 5) return '🟢 Healthy Growth';
  if (spendPct <= -5 && profitPct >= 8) return '🟢 Efficiency Improvement';
  if (spendPct <= -15) {
    if (profitPct >= 0) return '🔵 Scaling Down - Efficient';
    if (profitPct >= -10 && profitPct <= -1) return '🔵 Scaling Down - Moderate';
    if (profitPct <= -15) return '🔵 Scaling Down - Problematic';
    return '🔵 Scaling Down';
  }
  if (Math.abs(spendPct) <= 2 && Math.abs(profitPct) <= 2) return '⚪ Stable';
  return '⚪ Stable';
}

function calculateGrowthStatus(prev, curr, spendPct, profitPct, profitField = 'eProfitForecast') {
  return calculateGrowthStatusFast(prev, curr, spendPct, profitPct, profitField);
}

function calculateProjectGrowthStatus(projectName, prev, curr, spendPct, profitPct, profitField = 'eProfitForecast') {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    return calculateGrowthStatusFast(prev, curr, spendPct, profitPct, profitField);
  } finally {
    setCurrentProject(originalProject);
  }
}

function getGrowthStatusExplanation() {
  return `Growth Status Criteria for ${CURRENT_PROJECT}:

🟢 ПОЗИТИВНЫЕ: Healthy Growth (Spend ≥10% AND Profit ≥5%), Efficiency Improvement (спенд падает, профит растет), переход из убытка в прибыль
🔴 КРИТИЧЕСКИЕ: Inefficient Growth (Profit ≤-8%), переход из прибыли в убыток  
🟠 ПРЕДУПРЕЖДАЮЩИЕ: Declining Efficiency (спенд растет/стабилен, профит падает умеренно)
🔵 СОКРАЩЕНИЕ: Scaling Down (Spend ≤-15%) - Efficient/Moderate/Problematic
🟡 УМЕРЕННЫЕ: Moderate Growth/Decline, Minimal Growth, различные паттерны
⚪ СТАБИЛЬНЫЕ: Минимальные изменения в любую сторону`;
}

function getProjectGrowthStatusExplanation(projectName) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    return getGrowthStatusExplanation();
  } finally {
    setCurrentProject(originalProject);
  }
}

function analyzeGrowthScenario(spendPct, profitPct, projectName = CURRENT_PROJECT) {
  const mockPrev = { eProfitForecast: 100, spend: 100 };
  const mockCurr = { eProfitForecast: 100 + profitPct, spend: 100 + spendPct };
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  const status = calculateGrowthStatusFast(mockPrev, mockCurr, spendPct, profitPct);
  setCurrentProject(originalProject);
  return { spendPct, profitPct, projectName, status };
}

function generateReport(days) {
  console.log(`generateReport: start for ${CURRENT_PROJECT}, days: ${days}`);
  try {
    console.log('generateReport: getting date range');
    const dateRange = getDateRange(days);
    console.log(`generateReport: date range: ${dateRange.from} to ${dateRange.to}`);
    
    console.log('generateReport: fetching data');
    const raw = fetchCampaignData(dateRange);
    
    if (!raw.data?.analytics?.richStats?.stats?.length) {
      console.log('generateReport: no data found');
      SpreadsheetApp.getUi().alert('No data found for the specified period.');
      return;
    }
    
    console.log(`generateReport: processing ${raw.data.analytics.richStats.stats.length} records`);
    const processed = processApiData(raw);
    
    const processedCount = Object.keys(processed).length;
    console.log(`generateReport: processed ${processedCount} apps`);
    
    if (processedCount === 0) {
      console.log('generateReport: no valid data after processing');
      SpreadsheetApp.getUi().alert('No valid data to process.');
      return;
    }
    
    console.log('generateReport: clearing old data');
    clearAllDataSilent();
    
    console.log('generateReport: creating pivot table');
    if (CURRENT_PROJECT === 'OVERALL') {
      createOverallPivotTable(processed);
    } else {
      createEnhancedPivotTable(processed);
    }
    
    console.log('generateReport: applying comments');
    const cache = new CommentCache();
    cache.applyCommentsToSheet();
    
    console.log('generateReport: done');
  } catch (e) {
    console.error('generateReport error:', e);
    SpreadsheetApp.getUi().alert('Error', 'Error generating report: ' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function generateReportForDateRange(startDate, endDate) {
  console.log(`generateReportForDateRange: ${CURRENT_PROJECT}, ${startDate} to ${endDate}`);
  const ui = SpreadsheetApp.getUi();
  
  try {
    const dateRange = { from: startDate, to: endDate };
    console.log('generateReportForDateRange: fetching data');
    const raw = fetchCampaignData(dateRange);
    
    if (!raw.data?.analytics?.richStats?.stats?.length) {
      console.log('generateReportForDateRange: no data found');
      ui.alert('No Data', 'No data found for the selected date range.', ui.ButtonSet.OK);
      return;
    }
    
    console.log(`generateReportForDateRange: processing ${raw.data.analytics.richStats.stats.length} records`);
    const processed = processApiData(raw, true);
    
    const processedCount = Object.keys(processed).length;
    console.log(`generateReportForDateRange: processed ${processedCount} apps`);
    
    if (processedCount === 0) {
      console.log('generateReportForDateRange: no valid data after processing');
      ui.alert('No Valid Data', 'No valid data to process for the selected date range.', ui.ButtonSet.OK);
      return;
    }
    
    console.log('generateReportForDateRange: clearing old data');
    clearAllDataSilent();
    
    console.log('generateReportForDateRange: creating pivot table');
    if (CURRENT_PROJECT === 'OVERALL') {
      createOverallPivotTable(processed);
    } else {
      createEnhancedPivotTable(processed);
    }
    
    console.log('generateReportForDateRange: applying comments');
    const cache = new CommentCache();
    cache.applyCommentsToSheet();
    
    console.log('generateReportForDateRange: done');
    ui.alert('Success', `Report generated successfully!\n\nDate range: ${startDate} to ${endDate}`, ui.ButtonSet.OK);
  } catch (e) {
    console.error('generateReportForDateRange error:', e);
    ui.alert('Error', 'Error generating report:\n\n' + e.toString() + '\n\nPlease check:\n1. Your internet connection\n2. The API token is still valid\n3. Try a smaller date range', ui.ButtonSet.OK);
  }
}

function updateProjectData(projectName) {
  updateProjectDataOptimized(projectName);
}

function updateAllDataToCurrent() {
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
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'WEEK') {
        const weekRange = data[i][1];
        const [startStr] = weekRange.split(' - ');
        const startDate = new Date(startStr);
        if (!earliestDate || startDate < earliestDate) earliestDate = startDate;
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
    
    if (CURRENT_PROJECT === 'OVERALL') {
      createOverallPivotTable(processed);
    } else {
      createEnhancedPivotTable(processed);
    }
    
    const cache = new CommentCache();
    cache.applyCommentsToSheet();
    
    ui.alert('Success', `Successfully updated all data from ${dateRange.from} to ${dateRange.to}!`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error updating data: ' + e.toString(), ui.ButtonSet.OK);
  }
}