/**
 * Analytics Functions - ОБНОВЛЕНО: WoW аналитика с новыми ROAS метриками и поддержкой сетей для OVERALL
 */

function calculateWoWMetrics(appData) {
  if (!appData || typeof appData !== 'object') {
    console.error('Invalid appData provided to calculateWoWMetrics');
    return { campaignWoW: {}, appWeekWoW: {}, sourceAppWoW: {} };
  }

  try {
    const campaignData = {};
    const appWeekData = {};
    const sourceAppData = {};
    const networkData = {}; // Добавляем для OVERALL

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
                  campaignName: c.campaignName,
                  sourceApp: c.sourceApp,
                  weekStart: week.weekStart,
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
          // Обработка сеток для OVERALL
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
          // Обратная совместимость со старой структурой
          allCampaigns = week.campaigns || [];
        } else if (CURRENT_PROJECT === 'OVERALL') {
          allCampaigns = week.campaigns || [];
        } else {
          allCampaigns = week.campaigns || [];
          
          allCampaigns.forEach(c => {
            if (c.campaignId) {
              campaignData[`${c.campaignId}_${week.weekStart}`] = {
                campaignId: c.campaignId,
                campaignName: c.campaignName,
                sourceApp: c.sourceApp,
                weekStart: week.weekStart,
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
        
        const spend = allCampaigns.reduce((s, c) => s + c.spend, 0);
        const profit = allCampaigns.reduce((s, c) => s + c.eProfitForecast, 0);
        appWeekData[app.appName][week.weekStart] = { weekStart: week.weekStart, spend, profit };
      });
    });

    const campaignWoW = {};
    
    // Обработка WoW для сеток в OVERALL
    if (CURRENT_PROJECT === 'OVERALL') {
      const networks = {};
      Object.values(networkData).forEach(d => {
        Object.values(d).forEach(weekData => {
          if (!networks[weekData.networkId]) networks[weekData.networkId] = [];
          networks[weekData.networkId].push(weekData);
        });
      });

      Object.keys(networks).forEach(networkId => {
        networks[networkId].sort((a, b) => new Date(a.weekStart) - new Date(b.weekStart));
        networks[networkId].forEach((curr, i) => {
          const key = `${networkId}_${curr.weekStart}`;
          campaignWoW[key] = { spendChangePercent: 0, eProfitChangePercent: 0, growthStatus: 'First Week' };
          
          if (i > 0) {
            const prev = networks[networkId][i - 1];
            const spendPct = prev.spend ? ((curr.spend - prev.spend) / Math.abs(prev.spend)) * 100 : 0;
            const profitPct = prev.profit ? ((curr.profit - prev.profit) / Math.abs(prev.profit)) * 100 : 0;
            campaignWoW[key] = { 
              spendChangePercent: spendPct, 
              eProfitChangePercent: profitPct, 
              growthStatus: calculateGrowthStatus(prev, curr, spendPct, profitPct, 'profit') 
            };
          }
        });
      });
    } else if (CURRENT_PROJECT !== 'OVERALL') {
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
              growthStatus: calculateGrowthStatus(prev, curr, spendPct, profitPct) 
            };
          }
        });
      });
    }

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
            growthStatus: calculateGrowthStatus(prev, curr, spendPct, profitPct, 'profit') 
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
              growthStatus: calculateGrowthStatus(prev, curr, spendPct, profitPct, 'profit') 
            };
          }
        });
      });
    }

    return { campaignWoW, appWeekWoW, sourceAppWoW };
  } catch (e) {
    console.error('Error calculating WoW metrics:', e);
    return { campaignWoW: {}, appWeekWoW: {}, sourceAppWoW: {} };
  }
}

function calculateGrowthStatus(prev, curr, spendPct, profitPct, profitField = 'eProfitForecast') {
  const prevProfit = profitField === 'profit' ? prev.profit : prev.eProfitForecast;
  const currProfit = profitField === 'profit' ? curr.profit : curr.eProfitForecast;
  
  const t = {
    healthyGrowth: { minSpendChange: 10, minProfitChange: 5 },
    efficiencyImprovement: { maxSpendDecline: -5, minProfitGrowth: 8 },
    inefficientGrowth: { minSpendChange: 0, maxProfitChange: -8 },
    decliningEfficiency: { minSpendStable: -2, maxSpendGrowth: 10, maxProfitDecline: -4, minProfitDecline: -7 },
    scalingDown: { maxSpendChange: -15, efficient: { minProfitChange: 0 }, moderate: { maxProfitDecline: -10, minProfitDecline: -1 }, problematic: { maxProfitDecline: -15 } },
    moderateGrowthSpend: 3, moderateGrowthProfit: 2,
    minimalGrowth: { maxSpendChange: 2, maxProfitChange: 1 },
    moderateDecline: { maxSpendDecline: -3, maxProfitDecline: -3, spendOptimizationRatio: 1.5, efficiencyDropRatio: 1.5, proportionalRatio: 1.3 },
    stable: { maxAbsoluteChange: 2 }
  };
  
  if (prevProfit < 0 && currProfit > 0) return '🟢 Healthy Growth';
  if (prevProfit > 0 && currProfit < 0) return '🔴 Inefficient Growth';
  if (profitPct <= t.inefficientGrowth.maxProfitChange) return '🔴 Inefficient Growth';
  if (spendPct >= t.healthyGrowth.minSpendChange && profitPct >= t.healthyGrowth.minProfitChange) return '🟢 Healthy Growth';
  if (spendPct <= t.efficiencyImprovement.maxSpendDecline && profitPct >= t.efficiencyImprovement.minProfitGrowth) return '🟢 Efficiency Improvement';
  if (spendPct <= t.efficiencyImprovement.maxSpendDecline && profitPct > 0 && profitPct < t.efficiencyImprovement.minProfitGrowth) return '🟡 Efficiency Improvement';
  
  if (spendPct <= t.scalingDown.maxSpendChange) {
    if (profitPct >= t.scalingDown.efficient.minProfitChange) return '🔵 Scaling Down - Efficient';
    if (profitPct >= t.scalingDown.moderate.minProfitDecline && profitPct <= t.scalingDown.moderate.maxProfitDecline) return '🔵 Scaling Down - Moderate';
    if (profitPct <= t.scalingDown.problematic.maxProfitDecline) return '🔵 Scaling Down - Problematic';
    return '🔵 Scaling Down';
  }
  
  if (spendPct >= t.decliningEfficiency.minSpendStable && spendPct <= t.decliningEfficiency.maxSpendGrowth && profitPct >= t.decliningEfficiency.maxProfitDecline && profitPct <= t.decliningEfficiency.minProfitDecline) return '🟠 Declining Efficiency';
  
  if (spendPct < 0 && profitPct < 0 && spendPct >= t.moderateDecline.maxSpendDecline && profitPct >= t.moderateDecline.maxProfitDecline) {
    const spendDeclineAbs = Math.abs(spendPct);
    const profitDeclineAbs = Math.abs(profitPct);
    if (spendDeclineAbs >= profitDeclineAbs * t.moderateDecline.spendOptimizationRatio) return '🟡 Moderate Decline - Spend Optimization';
    if (profitDeclineAbs >= spendDeclineAbs * t.moderateDecline.efficiencyDropRatio) return '🟡 Moderate Decline - Efficiency Drop';
    return '🟡 Moderate Decline - Proportional';
  }
  
  if (spendPct >= t.moderateGrowthSpend && profitPct >= t.moderateGrowthProfit && (spendPct < t.healthyGrowth.minSpendChange || profitPct < t.healthyGrowth.minProfitChange)) return '🟡 Moderate Growth';
  if (spendPct > 0 && profitPct > 0) {
    if (spendPct <= t.minimalGrowth.maxSpendChange && profitPct <= t.minimalGrowth.maxProfitChange) return '🟡 Minimal Growth';
    if (spendPct < t.moderateGrowthSpend || profitPct < t.moderateGrowthProfit) return '🟡 Minimal Growth';
  }
  if (spendPct < 0 && profitPct < 0 && spendPct >= t.scalingDown.maxSpendChange && profitPct >= t.inefficientGrowth.maxProfitChange) return '🟡 Moderate Decline';
  if (spendPct > 0 && spendPct <= 15 && profitPct < 0 && profitPct >= -10) return '🟠 Declining Efficiency';
  if (Math.abs(spendPct) <= 5 && profitPct < -2 && profitPct >= -12) return '🟠 Declining Efficiency';
  if (Math.abs(spendPct) <= t.stable.maxAbsoluteChange && Math.abs(profitPct) <= t.stable.maxAbsoluteChange) return '⚪ Stable';
  if (Math.abs(spendPct) <= 10 && Math.abs(profitPct) <= 10) return '⚪ Stable';
  return '⚪ Stable';
}

function calculateProjectGrowthStatus(projectName, prev, curr, spendPct, profitPct, profitField = 'eProfitForecast') {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    return calculateGrowthStatus(prev, curr, spendPct, profitPct, profitField);
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
  const status = calculateGrowthStatus(mockPrev, mockCurr, spendPct, profitPct);
  setCurrentProject(originalProject);
  return { spendPct, profitPct, projectName, status };
}

function generateReport(days) {
  try {
    const dateRange = getDateRange(days);
    const raw = fetchCampaignData(dateRange);
    
    if (!raw.data?.analytics?.richStats?.stats?.length) {
      SpreadsheetApp.getUi().alert('No data found for the specified period.');
      return;
    }
    
    const processed = processApiData(raw);
    if (Object.keys(processed).length === 0) {
      SpreadsheetApp.getUi().alert('No valid data to process.');
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
  } catch (e) {
    console.error('Error generating report:', e);
    SpreadsheetApp.getUi().alert('Error', 'Error generating report: ' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function generateReportForDateRange(startDate, endDate) {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const dateRange = { from: startDate, to: endDate };
    const raw = fetchCampaignData(dateRange);
    
    if (!raw.data?.analytics?.richStats?.stats?.length) {
      ui.alert('No Data', 'No data found for the selected date range.', ui.ButtonSet.OK);
      return;
    }
    
    const processed = processApiData(raw, true);
    if (Object.keys(processed).length === 0) {
      ui.alert('No Valid Data', 'No valid data to process for the selected date range.', ui.ButtonSet.OK);
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
    
    ui.alert('Success', `Report generated successfully!\n\nDate range: ${startDate} to ${endDate}`, ui.ButtonSet.OK);
  } catch (e) {
    console.error('Error generating report for date range:', e);
    ui.alert('Error', 'Error generating report:\n\n' + e.toString() + '\n\nPlease check:\n1. Your internet connection\n2. The API token is still valid\n3. Try a smaller date range', ui.ButtonSet.OK);
  }
}

function updateProjectData(projectName) {
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log(`${projectName}: No existing data to update`);
    return;
  }
  
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
    console.log(`${projectName}: No week data found`);
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
  
  console.log(`${projectName}: Fetching data from ${dateRange.from} to ${dateRange.to}`);
  
  const raw = fetchProjectCampaignData(projectName, dateRange);
  
  if (!raw.data?.analytics?.richStats?.stats?.length) {
    console.log(`${projectName}: No data returned from API`);
    return;
  }
  
  const processed = processProjectApiData(projectName, raw);
  
  if (Object.keys(processed).length === 0) {
    console.log(`${projectName}: No valid data to process`);
    return;
  }
  
  clearProjectDataSilent(projectName);
  
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    if (projectName === 'OVERALL') {
      createOverallPivotTable(processed);
    } else {
      createEnhancedPivotTable(processed);
    }
    const cache = new CommentCache(projectName);
    cache.applyCommentsToSheet();
  } finally {
    setCurrentProject(originalProject);
  }
  
  console.log(`${projectName}: Update completed`);
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
    console.error('Error updating data:', e);
    ui.alert('Error', 'Error updating data: ' + e.toString(), ui.ButtonSet.OK);
  }
}