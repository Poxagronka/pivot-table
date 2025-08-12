// Unified WoW calculation
function calculateWoWMetrics(appData) {
  if (!appData || typeof appData !== 'object') {
    console.error('Invalid appData provided to calculateWoWMetrics');
    return { campaignWoW: {}, appWeekWoW: {}, sourceAppWoW: {} };
  }

  try {
    const results = {
      campaignWoW: {},
      appWeekWoW: {},
      sourceAppWoW: {},
      networkWoW: {},
      weekWoW: {},
      appWoW: {}
    };
    
    // For INCENT_TRAFFIC, structure is different
    if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
      return calculateIncentTrafficWoWMetrics(appData);
    }
    
    // Standard processing
    const dataCollections = {
      campaign: {},
      appWeek: {},
      sourceApp: {},
      network: {}
    };
    
    // Collect all data
    Object.values(appData).forEach(app => {
      dataCollections.appWeek[app.appName] = {};
      
      Object.values(app.weeks).forEach(week => {
        let allCampaigns = [];
        
        if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
          // TRICKY: Process source apps
          Object.values(week.sourceApps).forEach(sourceApp => {
            allCampaigns.push(...sourceApp.campaigns);
            
            const sourceAppKey = sourceApp.sourceAppId;
            if (!dataCollections.sourceApp[sourceAppKey]) {
              dataCollections.sourceApp[sourceAppKey] = {};
            }
            
            dataCollections.sourceApp[sourceAppKey][week.weekStart] = {
              weekStart: week.weekStart,
              spend: sourceApp.campaigns.reduce((s, c) => s + c.spend, 0),
              profit: sourceApp.campaigns.reduce((s, c) => s + c.eProfitForecast, 0)
            };
            
            sourceApp.campaigns.forEach(c => {
              if (c.campaignId) {
                dataCollections.campaign[`${c.campaignId}_${week.weekStart}`] = {
                  campaignId: c.campaignId,
                  weekStart: week.weekStart,
                  spend: c.spend,
                  eProfitForecast: c.eProfitForecast
                };
              }
            });
          });
        } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
          // OVERALL: Process networks
          Object.values(week.networks).forEach(network => {
            allCampaigns.push(...network.campaigns);
            
            const networkKey = network.networkId;
            if (!dataCollections.network[networkKey]) {
              dataCollections.network[networkKey] = {};
            }
            
            dataCollections.network[networkKey][week.weekStart] = {
              weekStart: week.weekStart,
              spend: network.campaigns.reduce((s, c) => s + c.spend, 0),
              profit: network.campaigns.reduce((s, c) => s + c.eProfitForecast, 0)
            };
          });
        } else {
          // Regular projects
          allCampaigns = week.campaigns || [];
          
          allCampaigns.forEach(c => {
            if (c.campaignId) {
              dataCollections.campaign[`${c.campaignId}_${week.weekStart}`] = {
                campaignId: c.campaignId,
                weekStart: week.weekStart,
                spend: c.spend,
                eProfitForecast: c.eProfitForecast
              };
            }
          });
        }
        
        // App week data
        const spend = allCampaigns.reduce((s, c) => s + c.spend, 0);
        const profit = allCampaigns.reduce((s, c) => s + c.eProfitForecast, 0);
        dataCollections.appWeek[app.appName][week.weekStart] = { weekStart: week.weekStart, spend, profit };
      });
    });
    
    // Calculate WoW for all collections
    const calculateCollectionWoW = (collection, resultKey, profitField = 'eProfitForecast') => {
      Object.keys(collection).forEach(entityId => {
        const weeks = Object.values(collection[entityId]).sort((a, b) => new Date(a.weekStart) - new Date(b.weekStart));
        
        weeks.forEach((curr, i) => {
          const key = `${entityId}_${curr.weekStart}`;
          
          if (i === 0) {
            results[resultKey][key] = { spendChangePercent: 0, eProfitChangePercent: 0, growthStatus: 'First Week' };
          } else {
            const prev = weeks[i - 1];
            const spendPct = prev.spend ? ((curr.spend - prev.spend) / Math.abs(prev.spend)) * 100 : 0;
            const profitValue = profitField === 'profit' ? curr.profit : curr.eProfitForecast;
            const prevProfitValue = profitField === 'profit' ? prev.profit : prev.eProfitForecast;
            const profitPct = prevProfitValue ? ((profitValue - prevProfitValue) / Math.abs(prevProfitValue)) * 100 : 0;
            
            results[resultKey][key] = {
              spendChangePercent: spendPct,
              eProfitChangePercent: profitPct,
              growthStatus: calculateGrowthStatus(prev, curr, spendPct, profitPct, profitField)
            };
          }
        });
      });
    };
    
    // Calculate for each type
    calculateCollectionWoW(dataCollections.campaign, 'campaignWoW');
    calculateCollectionWoW(dataCollections.appWeek, 'appWeekWoW', 'profit');
    
    if (CURRENT_PROJECT === 'TRICKY') {
      calculateCollectionWoW(dataCollections.sourceApp, 'sourceAppWoW', 'profit');
    } else if (CURRENT_PROJECT === 'OVERALL') {
      calculateCollectionWoW(dataCollections.network, 'campaignWoW', 'profit');
    }
    
    return { 
      campaignWoW: results.campaignWoW, 
      appWeekWoW: results.appWeekWoW, 
      sourceAppWoW: results.sourceAppWoW 
    };
    
  } catch (e) {
    console.error('Error calculating WoW metrics:', e);
    return { campaignWoW: {}, appWeekWoW: {}, sourceAppWoW: {} };
  }
}

function calculateIncentTrafficWoWMetrics(networkData) {
  const results = { weekWoW: {}, appWoW: {}, networkWoW: {} };
  
  Object.keys(networkData).forEach(networkKey => {
    const network = networkData[networkKey];
    const weeks = Object.values(network.weeks).sort((a, b) => new Date(a.weekStart) - new Date(b.weekStart));
    
    const appHistory = {};
    
    weeks.forEach((week, i) => {
      const weekKey = `${networkKey}_${week.weekStart}`;
      const allCampaigns = [];
      Object.values(week.apps).forEach(app => allCampaigns.push(...app.campaigns));
      
      const spend = allCampaigns.reduce((s, c) => s + c.spend, 0);
      const profit = allCampaigns.reduce((s, c) => s + c.eProfitForecast, 0);
      
      if (i === 0) {
        results.weekWoW[weekKey] = { spendChangePercent: 0, eProfitChangePercent: 0, growthStatus: 'First Week' };
      } else {
        const prevWeek = weeks[i - 1];
        const prevCampaigns = [];
        Object.values(prevWeek.apps).forEach(app => prevCampaigns.push(...app.campaigns));
        const prevSpend = prevCampaigns.reduce((s, c) => s + c.spend, 0);
        const prevProfit = prevCampaigns.reduce((s, c) => s + c.eProfitForecast, 0);
        
        const spendPct = prevSpend ? ((spend - prevSpend) / Math.abs(prevSpend)) * 100 : 0;
        const profitPct = prevProfit ? ((profit - prevProfit) / Math.abs(prevProfit)) * 100 : 0;
        
        results.weekWoW[weekKey] = {
          spendChangePercent: spendPct,
          eProfitChangePercent: profitPct,
          growthStatus: calculateGrowthStatus({ spend: prevSpend, profit: prevProfit }, { spend, profit }, spendPct, profitPct, 'profit')
        };
      }
      
      // App level WoW
      Object.keys(week.apps).forEach(appId => {
        const appData = week.apps[appId];
        const appKey = `${networkKey}_${week.weekStart}_${appId}`;
        const appSpend = appData.campaigns.reduce((s, c) => s + c.spend, 0);
        const appProfit = appData.campaigns.reduce((s, c) => s + c.eProfitForecast, 0);
        
        if (appHistory[appId]?.length > 0) {
          const prevAppData = appHistory[appId][appHistory[appId].length - 1];
          const spendPct = prevAppData.spend ? ((appSpend - prevAppData.spend) / Math.abs(prevAppData.spend)) * 100 : 0;
          const profitPct = prevAppData.profit ? ((appProfit - prevAppData.profit) / Math.abs(prevAppData.profit)) * 100 : 0;
          
          results.appWoW[appKey] = {
            spendChangePercent: spendPct,
            eProfitChangePercent: profitPct,
            growthStatus: calculateGrowthStatus({ spend: prevAppData.spend, profit: prevAppData.profit }, { spend: appSpend, profit: appProfit }, spendPct, profitPct, 'profit')
          };
        } else {
          results.appWoW[appKey] = { spendChangePercent: 0, eProfitChangePercent: 0, growthStatus: 'First Week' };
        }
        
        if (!appHistory[appId]) appHistory[appId] = [];
        appHistory[appId].push({ weekStart: week.weekStart, spend: appSpend, profit: appProfit });
      });
    });
  });
  
  return results;
}

// Simplified growth status calculation
function calculateGrowthStatus(prev, curr, spendPct, profitPct, profitField = 'eProfitForecast') {
  const prevProfit = profitField === 'profit' ? prev.profit : prev.eProfitForecast;
  const currProfit = profitField === 'profit' ? curr.profit : curr.eProfitForecast;
  
  const t = getGrowthThresholds(CURRENT_PROJECT);
  
  // Special cases
  if (prevProfit < 0 && currProfit > 0) return '🟢 Healthy Growth';
  if (prevProfit > 0 && currProfit < 0) return '🔴 Inefficient Growth';
  
  // Check thresholds
  const checks = [
    { condition: profitPct <= t.inefficientGrowth.maxProfitChange, status: '🔴 Inefficient Growth' },
    { condition: spendPct >= t.healthyGrowth.minSpendChange && profitPct >= t.healthyGrowth.minProfitChange, status: '🟢 Healthy Growth' },
    { condition: spendPct <= t.efficiencyImprovement.maxSpendDecline && profitPct >= t.efficiencyImprovement.minProfitGrowth, status: '🟢 Efficiency Improvement' },
    { condition: spendPct <= t.scalingDown.maxSpendChange, status: getScalingDownStatus(profitPct, t.scalingDown) },
    { condition: Math.abs(spendPct) <= t.stable.maxAbsoluteChange && Math.abs(profitPct) <= t.stable.maxAbsoluteChange, status: '⚪ Stable' }
  ];
  
  for (const check of checks) {
    if (check.condition) return check.status;
  }
  
  // Default cases
  if (spendPct > 0 && profitPct > 0) {
    if (spendPct >= t.moderateGrowthSpend && profitPct >= t.moderateGrowthProfit) return '🟡 Moderate Growth';
    return '🟡 Minimal Growth';
  }
  
  if (spendPct < 0 && profitPct < 0) return '🟡 Moderate Decline';
  
  return '⚪ Stable';
}

function getScalingDownStatus(profitPct, scalingConfig) {
  if (profitPct >= scalingConfig.efficient.minProfitChange) return '🔵 Scaling Down - Efficient';
  if (profitPct >= scalingConfig.moderate.minProfitDecline && profitPct <= scalingConfig.moderate.maxProfitDecline) return '🔵 Scaling Down - Moderate';
  if (profitPct <= scalingConfig.problematic.maxProfitDecline) return '🔵 Scaling Down - Problematic';
  return '🔵 Scaling Down';
}

// Unified report generation
function generateReport(days) {
  executeReport(() => getDateRange(days), days);
}

function generateReportForDateRange(startDate, endDate) {
  executeReport(() => ({ from: startDate, to: endDate }), `${startDate} to ${endDate}`);
}

function executeReport(getDateRangeFn, description) {
  try {
    const dateRange = getDateRangeFn();
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
    
    // Create appropriate table based on project
    const tableCreators = {
      'OVERALL': createOverallPivotTable,
      'INCENT_TRAFFIC': createIncentTrafficPivotTable
    };
    
    const createTable = tableCreators[CURRENT_PROJECT] || createEnhancedPivotTable;
    createTable(processed);
    
    const cache = new CommentCache();
    cache.applyCommentsToSheet();
    
    console.log(`Report generated for ${description}`);
  } catch (e) {
    console.error('Error generating report:', e);
    SpreadsheetApp.getUi().alert('Error', 'Error generating report: ' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Unified update functions
function updateProjectData(projectName) {
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log(`${projectName}: No existing data to update`);
    return;
  }
  
  // Save comments before update
  try {
    const cache = new CommentCache(projectName);
    cache.syncCommentsFromSheet();
    console.log(`${projectName}: Comments saved`);
  } catch (e) {
    console.error(`${projectName}: Failed to save comments:`, e);
  }
  
  // Find earliest date using Sheets API v4
  const range = `${config.SHEET_NAME}!A:B`;
  const response = Sheets.Spreadsheets.Values.get(config.SHEET_ID, range);
  const data = response.values || [];
  
  let earliestDate = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i]?.[0] === 'WEEK' && data[i]?.[1]) {
      const [startStr] = data[i][1].split(' - ');
      const startDate = new Date(startStr);
      if (!earliestDate || startDate < earliestDate) earliestDate = startDate;
    }
  }
  
  if (!earliestDate) {
    console.log(`${projectName}: No week data found`);
    return;
  }
  
  // Calculate end date (last complete week)
  const today = new Date();
  const dayOfWeek = today.getDay();
  const endDate = new Date(today);
  endDate.setDate(today.getDate() - (dayOfWeek === 0 ? 1 : dayOfWeek));
  
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
    const tableCreators = {
      'OVERALL': createOverallPivotTable,
      'INCENT_TRAFFIC': createIncentTrafficPivotTable
    };
    
    const createTable = tableCreators[projectName] || createEnhancedPivotTable;
    createTable(processed);
    
    const cache = new CommentCache(projectName);
    cache.applyCommentsToSheet();
  } finally {
    setCurrentProject(originalProject);
  }
  
  console.log(`${projectName}: Update completed`);
}

function updateAllDataToCurrent() {
  updateProjectData(CURRENT_PROJECT);
}

// Legacy functions for compatibility (keep signatures!)
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