/**
 * Sheet Formatting and Table Creation - ОПТИМИЗИРОВАНО: двухфазная группировка с батчингом по сущностям
 */

function createEnhancedPivotTable(appData) { createUnifiedPivotTable(appData); }
function createOverallPivotTable(appData) { createUnifiedPivotTable(appData); }
function createIncentTrafficPivotTable(networkData) { createUnifiedPivotTable(networkData); }

function createUnifiedPivotTable(data) {
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);
  else sheet.clear();

  if (!data || Object.keys(data).length === 0) {
    console.log(`${CURRENT_PROJECT}: No data to display`);
    const headers = getUnifiedHeaders();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const initialEROASCache = new InitialEROASCache();
  initialEROASCache.recordInitialValuesFromData(data);

  const wow = CURRENT_PROJECT === 'INCENT_TRAFFIC' ? 
    calculateIncentTrafficWoWMetrics(data) : 
    calculateWoWMetrics(data);
  
  const headers = getUnifiedHeaders();
  const tableData = [headers];
  const formatData = [];

  buildUnifiedTable(data, tableData, formatData, wow, initialEROASCache);

  const range = sheet.getRange(1, 1, tableData.length, headers.length);
  range.setValues(tableData);
  
  applyEnhancedFormatting(sheet, tableData.length, headers.length, formatData, data);
  createUnifiedRowGrouping(sheet, tableData, data);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

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

function createUnifiedRowGrouping(sheet, tableData, data) {
  try {
    console.log('Starting two-phase unified row grouping...');
    const startTime = new Date().getTime();
    
    const sheetId = sheet.getSheetId();
    const spreadsheetId = sheet.getParent().getId();
    
    if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
      const sortedNetworks = Object.keys(data).sort((a, b) => 
        data[a].networkName.localeCompare(data[b].networkName)
      );
      
      let networkIndex = 0;
      for (const networkKey of sortedNetworks) {
        console.log(`Processing network ${networkIndex + 1}/${sortedNetworks.length}: ${data[networkKey].networkName}`);
        
        processEntityGroups(spreadsheetId, sheetId, data, networkKey, 'network');
        
        networkIndex++;
        if (networkIndex < sortedNetworks.length) {
          console.log('Waiting 5 seconds before next network...');
          Utilities.sleep(5000);
        }
      }
    } else {
      const sortedApps = Object.keys(data).sort((a, b) => 
        data[a].appName.localeCompare(data[b].appName)
      );
      
      let appIndex = 0;
      for (const appKey of sortedApps) {
        console.log(`Processing app ${appIndex + 1}/${sortedApps.length}: ${data[appKey].appName}`);
        
        processEntityGroups(spreadsheetId, sheetId, data, appKey, 'app');
        
        appIndex++;
        if (appIndex < sortedApps.length) {
          console.log('Waiting 5 seconds before next app...');
          Utilities.sleep(5000);
        }
      }
    }
    
    const endTime = new Date().getTime();
    console.log(`Two-phase unified row grouping completed in ${(endTime - startTime)/1000}s`);
    
  } catch (e) {
    console.error('Error in unified row grouping:', e);
  }
}

function processEntityGroups(spreadsheetId, sheetId, data, entityKey, entityType) {
  try {
    console.log(`Phase 1: Creating groups for ${entityType} ${entityKey}`);
    const createRequests = buildCreateGroupsForEntity(data, entityKey, entityType, sheetId);
    
    if (createRequests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({
        requests: createRequests
      }, spreadsheetId);
      console.log(`Created ${createRequests.length} groups`);
      
      Utilities.sleep(1000);
    }
    
    console.log(`Phase 2: Collapsing groups for ${entityType} ${entityKey}`);
    const collapseRequests = buildCollapseGroupsForEntity(data, entityKey, entityType, sheetId);
    
    if (collapseRequests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({
        requests: collapseRequests
      }, spreadsheetId);
      console.log(`Collapsed ${collapseRequests.length} groups`);
    }
    
  } catch (e) {
    console.error(`Error processing ${entityType} ${entityKey}:`, e);
  }
}

function buildCreateGroupsForEntity(data, entityKey, entityType, sheetId) {
  const groupRequests = [];
  let rowPointer = calculateRowPointer(data, entityKey, entityType);
  
  if (entityType === 'network') {
    const network = data[entityKey];
    const networkStartRow = rowPointer;
    rowPointer++;
    let networkTotalRows = 0;

    const sortedWeeks = Object.keys(network.weeks).sort();
    
    sortedWeeks.forEach(weekKey => {
      const week = network.weeks[weekKey];
      const weekStartRow = rowPointer;
      rowPointer++;
      
      const appCount = Object.keys(week.apps).length;
      rowPointer += appCount;
      networkTotalRows += 1 + appCount;
      
      if (appCount > 0) {
        groupRequests.push({
          addDimensionGroup: {
            range: {
              sheetId: sheetId,
              dimension: "ROWS",
              startIndex: weekStartRow,
              endIndex: weekStartRow + appCount
            }
          }
        });
      }
    });

    if (networkTotalRows > 0) {
      groupRequests.push({
        addDimensionGroup: {
          range: {
            sheetId: sheetId,
            dimension: "ROWS",
            startIndex: networkStartRow,
            endIndex: networkStartRow + networkTotalRows
          }
        }
      });
    }
  } else {
    const app = data[entityKey];
    const appStartRow = rowPointer;
    rowPointer++;
    let appTotalRows = 0;

    const sortedWeeks = Object.keys(app.weeks).sort();
    
    sortedWeeks.forEach(weekKey => {
      const week = app.weeks[weekKey];
      const weekStartRow = rowPointer;
      rowPointer++;
      let weekContentRows = 0;

      if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
        const sourceAppKeys = Object.keys(week.sourceApps).sort((a, b) => {
          const spendA = week.sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
          const spendB = week.sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
          return spendB - spendA;
        });
        
        sourceAppKeys.forEach(sourceAppKey => {
          const sourceApp = week.sourceApps[sourceAppKey];
          const sourceAppStartRow = rowPointer;
          rowPointer++;
          
          const campaignCount = sourceApp.campaigns.length;
          rowPointer += campaignCount;
          weekContentRows += 1 + campaignCount;
          
          if (campaignCount > 0) {
            groupRequests.push({
              addDimensionGroup: {
                range: {
                  sheetId: sheetId,
                  dimension: "ROWS",
                  startIndex: sourceAppStartRow,
                  endIndex: sourceAppStartRow + campaignCount
                }
              }
            });
          }
        });
        
      } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
        const networkCount = Object.keys(week.networks).length;
        rowPointer += networkCount;
        weekContentRows = networkCount;
        
        if (networkCount > 0) {
          groupRequests.push({
            addDimensionGroup: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: weekStartRow,
                endIndex: weekStartRow + networkCount
              }
            }
          });
        }
        
      } else if (CURRENT_PROJECT !== 'OVERALL' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
        const campaignCount = week.campaigns ? week.campaigns.length : 0;
        rowPointer += campaignCount;
        weekContentRows = campaignCount;
        
        if (campaignCount > 0) {
          groupRequests.push({
            addDimensionGroup: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: weekStartRow,
                endIndex: weekStartRow + campaignCount
              }
            }
          });
        }
      }
      
      if (weekContentRows > 0) {
        groupRequests.push({
          addDimensionGroup: {
            range: {
              sheetId: sheetId,
              dimension: "ROWS",
              startIndex: weekStartRow,
              endIndex: weekStartRow + weekContentRows
            }
          }
        });
        
        appTotalRows += 1 + weekContentRows;
      } else {
        appTotalRows += 1;
      }
    });

    if (appTotalRows > 0) {
      groupRequests.push({
        addDimensionGroup: {
          range: {
            sheetId: sheetId,
            dimension: "ROWS",
            startIndex: appStartRow,
            endIndex: appStartRow + appTotalRows
          }
        }
      });
    }
  }
  
  return groupRequests;
}

function buildCollapseGroupsForEntity(data, entityKey, entityType, sheetId) {
  const collapseRequests = [];
  let rowPointer = calculateRowPointer(data, entityKey, entityType);
  
  if (entityType === 'network') {
    const network = data[entityKey];
    const networkStartRow = rowPointer;
    rowPointer++;
    let networkTotalRows = 0;

    const sortedWeeks = Object.keys(network.weeks).sort();
    
    sortedWeeks.forEach(weekKey => {
      const week = network.weeks[weekKey];
      const weekStartRow = rowPointer;
      rowPointer++;
      
      const appCount = Object.keys(week.apps).length;
      rowPointer += appCount;
      networkTotalRows += 1 + appCount;
      
      if (appCount > 0) {
        collapseRequests.push({
          updateDimensionGroup: {
            dimensionGroup: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: weekStartRow,
                endIndex: weekStartRow + appCount
              },
              depth: 2,
              collapsed: true
            },
            fields: "collapsed"
          }
        });
      }
    });

    if (networkTotalRows > 0) {
      collapseRequests.push({
        updateDimensionGroup: {
          dimensionGroup: {
            range: {
              sheetId: sheetId,
              dimension: "ROWS",
              startIndex: networkStartRow,
              endIndex: networkStartRow + networkTotalRows
            },
            depth: 1,
            collapsed: true
          },
          fields: "collapsed"
        }
      });
    }
  } else {
    const app = data[entityKey];
    const appStartRow = rowPointer;
    rowPointer++;
    let appTotalRows = 0;

    const sortedWeeks = Object.keys(app.weeks).sort();
    
    sortedWeeks.forEach(weekKey => {
      const week = app.weeks[weekKey];
      const weekStartRow = rowPointer;
      rowPointer++;
      let weekContentRows = 0;

      if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
        const sourceAppKeys = Object.keys(week.sourceApps).sort((a, b) => {
          const spendA = week.sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
          const spendB = week.sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
          return spendB - spendA;
        });
        
        sourceAppKeys.forEach(sourceAppKey => {
          const sourceApp = week.sourceApps[sourceAppKey];
          const sourceAppStartRow = rowPointer;
          rowPointer++;
          
          const campaignCount = sourceApp.campaigns.length;
          rowPointer += campaignCount;
          weekContentRows += 1 + campaignCount;
          
          if (campaignCount > 0) {
            collapseRequests.push({
              updateDimensionGroup: {
                dimensionGroup: {
                  range: {
                    sheetId: sheetId,
                    dimension: "ROWS",
                    startIndex: sourceAppStartRow,
                    endIndex: sourceAppStartRow + campaignCount
                  },
                  depth: 3,
                  collapsed: true
                },
                fields: "collapsed"
              }
            });
          }
        });
        
      } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
        const networkCount = Object.keys(week.networks).length;
        rowPointer += networkCount;
        weekContentRows = networkCount;
        
        if (networkCount > 0) {
          collapseRequests.push({
            updateDimensionGroup: {
              dimensionGroup: {
                range: {
                  sheetId: sheetId,
                  dimension: "ROWS",
                  startIndex: weekStartRow,
                  endIndex: weekStartRow + networkCount
                },
                depth: 2,
                collapsed: true
              },
              fields: "collapsed"
            }
          });
        }
        
      } else if (CURRENT_PROJECT !== 'OVERALL' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
        const campaignCount = week.campaigns ? week.campaigns.length : 0;
        rowPointer += campaignCount;
        weekContentRows = campaignCount;
        
        if (campaignCount > 0) {
          collapseRequests.push({
            updateDimensionGroup: {
              dimensionGroup: {
                range: {
                  sheetId: sheetId,
                  dimension: "ROWS",
                  startIndex: weekStartRow,
                  endIndex: weekStartRow + campaignCount
                },
                depth: 2,
                collapsed: true
              },
              fields: "collapsed"
            }
          });
        }
      }
      
      if (weekContentRows > 0) {
        collapseRequests.push({
          updateDimensionGroup: {
            dimensionGroup: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: weekStartRow,
                endIndex: weekStartRow + weekContentRows
              },
              depth: 2,
              collapsed: true
            },
            fields: "collapsed"
          }
        });
        
        appTotalRows += 1 + weekContentRows;
      } else {
        appTotalRows += 1;
      }
    });

    if (appTotalRows > 0) {
      collapseRequests.push({
        updateDimensionGroup: {
          dimensionGroup: {
            range: {
              sheetId: sheetId,
              dimension: "ROWS",
              startIndex: appStartRow,
              endIndex: appStartRow + appTotalRows
            },
            depth: 1,
            collapsed: true
          },
          fields: "collapsed"
        }
      });
    }
  }
  
  return collapseRequests;
}

function calculateRowPointer(data, entityKey, entityType) {
  let rowPointer = 2;
  
  if (entityType === 'network') {
    const sortedNetworks = Object.keys(data).sort((a, b) => 
      data[a].networkName.localeCompare(data[b].networkName)
    );
    
    for (const networkKey of sortedNetworks) {
      if (networkKey === entityKey) break;
      
      const network = data[networkKey];
      rowPointer++;
      
      Object.keys(network.weeks).forEach(weekKey => {
        const week = network.weeks[weekKey];
        rowPointer++; 
        rowPointer += Object.keys(week.apps).length;
      });
    }
  } else {
    const sortedApps = Object.keys(data).sort((a, b) => 
      data[a].appName.localeCompare(data[b].appName)
    );
    
    for (const appKey of sortedApps) {
      if (appKey === entityKey) break;
      
      const app = data[appKey];
      rowPointer++;
      
      Object.keys(app.weeks).forEach(weekKey => {
        const week = app.weeks[weekKey];
        rowPointer++;
        
        if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
          Object.keys(week.sourceApps).forEach(sourceAppKey => {
            const sourceApp = week.sourceApps[sourceAppKey];
            rowPointer++;
            rowPointer += sourceApp.campaigns.length;
          });
        } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
          rowPointer += Object.keys(week.networks).length;
        } else if (CURRENT_PROJECT !== 'OVERALL' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
          rowPointer += week.campaigns ? week.campaigns.length : 0;
        }
      });
    }
  }
  
  return rowPointer;
}

function getUnifiedHeaders() {
  return [
    'Level', 'Week Range / Source App', 'ID', 'GEO',
    'Spend', 'Spend WoW %', 'Installs', 'CPI', 'ROAS D1→D3→D7→D30', 'IPM',
    'RR D-1', 'RR D-7', 'eARPU 365d', 'eROAS 365d', 'eROAS 730d', 'eProfit 730d', 'eProfit 730d WoW %', 'Growth Status', 'Comments'
  ];
}

function applyEnhancedFormatting(sheet, numRows, numCols, formatData, appData) {
  const config = getCurrentConfig();
  
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground(COLORS.HEADER.background)
    .setFontColor(COLORS.HEADER.fontColor)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setWrap(true);

  const columnWidths = TABLE_CONFIG.COLUMN_WIDTHS;
  columnWidths.forEach(col => sheet.setColumnWidth(col.c, col.w));

  if (numRows > 1) {
    const allDataRange = sheet.getRange(2, 1, numRows - 1, numCols);
    allDataRange.setVerticalAlignment('middle');
    
    const roasRange = sheet.getRange(2, 9, numRows - 1, 1);
    roasRange.setWrap(true).setHorizontalAlignment('center');
    
    const commentsRange = sheet.getRange(2, numCols, numRows - 1, 1);
    commentsRange.setWrap(true).setHorizontalAlignment('left');
    
    const growthStatusRange = sheet.getRange(2, numCols - 1, numRows - 1, 1);
    growthStatusRange.setWrap(true).setHorizontalAlignment('left');

    const eroasRange = sheet.getRange(2, 15, numRows - 1, 1);
    eroasRange.setHorizontalAlignment('right');
  }

  const rowTypeMap = { app: [], week: [], sourceApp: [], campaign: [], hyperlink: [], network: [] };
  formatData.forEach(item => {
    if (item.type === 'APP') rowTypeMap.app.push(item.row);
    if (item.type === 'WEEK') rowTypeMap.week.push(item.row);
    if (item.type === 'SOURCE_APP') rowTypeMap.sourceApp.push(item.row);
    if (item.type === 'CAMPAIGN') rowTypeMap.campaign.push(item.row);
    if (item.type === 'NETWORK') rowTypeMap.network.push(item.row);
    if (item.type === 'HYPERLINK') rowTypeMap.hyperlink.push(item.row);
  });

  rowTypeMap.app.forEach(r => {
    if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
      sheet.getRange(r, 1, 1, numCols)
           .setBackground(COLORS.CAMPAIGN_ROW.background)
           .setFontWeight('normal')
           .setFontSize(9);
    } else {
      sheet.getRange(r, 1, 1, numCols)
           .setBackground(COLORS.APP_ROW.background)
           .setFontColor(COLORS.APP_ROW.fontColor)
           .setFontWeight('bold')
           .setFontSize(10);
    }
  });

  rowTypeMap.week.forEach(r =>
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.WEEK_ROW.background)
         .setFontSize(10)
  );

  rowTypeMap.sourceApp.forEach(r =>
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.SOURCE_APP_ROW.background)
         .setFontSize(10)
  );

  rowTypeMap.campaign.forEach(r =>
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.CAMPAIGN_ROW.background)
         .setFontSize(9)
  );

  rowTypeMap.network.forEach(r => {
    if (CURRENT_PROJECT === 'OVERALL') {
      sheet.getRange(r, 1, 1, numCols)
           .setBackground(COLORS.CAMPAIGN_ROW.background)
           .setFontWeight('normal')
           .setFontSize(9);
    } else {
      sheet.getRange(r, 1, 1, numCols)
           .setBackground(COLORS.APP_ROW.background)
           .setFontColor(COLORS.APP_ROW.fontColor)
           .setFontWeight('bold')
           .setFontSize(10);
    }
  });

  if (rowTypeMap.hyperlink.length > 0 && CURRENT_PROJECT === 'TRICKY') {
    rowTypeMap.hyperlink.forEach(r => {
      const linkCell = sheet.getRange(r, 2);
      linkCell.setFontColor('#000000').setFontLine('none');
    });
  }

  if (numRows > 1) {
    sheet.getRange(2, 5, numRows - 1, 1).setNumberFormat('$0');
    sheet.getRange(2, 8, numRows - 1, 1).setNumberFormat('$0.0');
    sheet.getRange(2, 10, numRows - 1, 1).setNumberFormat('0.0');
    sheet.getRange(2, 13, numRows - 1, 1).setNumberFormat('$0.0');
    sheet.getRange(2, 16, numRows - 1, 1).setNumberFormat('$0');
  }

  applyConditionalFormatting(sheet, numRows, appData);
  applyEROASRichTextFormatting(sheet, numRows);
  
  sheet.hideColumns(1);
  sheet.hideColumns(13, 1);
  sheet.hideColumns(14, 1);
}

function applyEROASRichTextFormatting(sheet, numRows) {
  if (numRows <= 1) return;
  
  const eroasColumn = 15;
  const range = sheet.getRange(2, eroasColumn, numRows - 1, 1);
  const values = range.getValues();
  
  const richTextValues = values.map(row => {
    const cellValue = row[0];
    if (!cellValue || typeof cellValue !== 'string' || !cellValue.includes('→')) {
      return SpreadsheetApp.newRichTextValue().setText(cellValue || '').build();
    }
    
    const arrowIndex = cellValue.indexOf('→');
    if (arrowIndex === -1) {
      return SpreadsheetApp.newRichTextValue().setText(cellValue).build();
    }
    
    const beforeArrow = cellValue.substring(0, arrowIndex);
    
    const richTextBuilder = SpreadsheetApp.newRichTextValue()
    .setText(cellValue)
    .setTextStyle(0, beforeArrow.length, SpreadsheetApp.newTextStyle()
    .setForegroundColor('#808080')
    .setFontSize(9)
    .build());
    
    return richTextBuilder.build();
  });
  
  range.setRichTextValues(richTextValues.map(rtv => [rtv]));
}

function applyConditionalFormatting(sheet, numRows, appData) {
  const rules = [];
  
  if (numRows > 1) {
    const spendRange = sheet.getRange(2, 6, numRows - 1, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('%').whenNumberGreaterThan(0)
        .setBackground(COLORS.POSITIVE.background)
        .setFontColor(COLORS.POSITIVE.fontColor)
        .setRanges([spendRange]).build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('%').whenNumberLessThan(0)
        .setBackground(COLORS.NEGATIVE.background)
        .setFontColor(COLORS.NEGATIVE.fontColor)
        .setRanges([spendRange]).build()
    );

    const eroasColumn = 15;
    const eroasRange = sheet.getRange(2, eroasColumn, numRows - 1, 1);
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const level = data[i][0];
      let appName = '';
      let targetEROAS = 150;
      
      if (level === 'APP') {
        appName = data[i][1];
        targetEROAS = getTargetEROAS(CURRENT_PROJECT, appName);
      } else {
        for (let j = i - 1; j >= 1; j--) {
          if (data[j][0] === 'APP') {
            appName = data[j][1];
            targetEROAS = getTargetEROAS(CURRENT_PROJECT, appName);
            break;
          }
        }
      }
      
      const cellRange = sheet.getRange(i + 1, eroasColumn, 1, 1);
      const columnLetter = String.fromCharCode(64 + eroasColumn);
      const cellAddress = `${columnLetter}${i + 1}`;
      
      const extractValueFormula = `IF(ISERROR(SEARCH("→",${cellAddress})), VALUE(SUBSTITUTE(${cellAddress},"%","")), VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(${cellAddress},"→",REPT(" ",100)),100)),"%","")))`;
      
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${cellAddress})), ${extractValueFormula} >= ${targetEROAS})`)
          .setBackground(COLORS.POSITIVE.background)
          .setFontColor(COLORS.POSITIVE.fontColor)
          .setRanges([cellRange]).build()
      );
      
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${cellAddress})), ${extractValueFormula} >= 120, ${extractValueFormula} < ${targetEROAS})`)
          .setBackground(COLORS.WARNING.background)
          .setFontColor(COLORS.WARNING.fontColor)
          .setRanges([cellRange]).build()
      );
      
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${cellAddress})), ${extractValueFormula} < 120)`)
          .setBackground(COLORS.NEGATIVE.background)
          .setFontColor(COLORS.NEGATIVE.fontColor)
          .setRanges([cellRange]).build()
      );
    }

    const profitColumn = 17;
    const profitRange = sheet.getRange(2, profitColumn, numRows - 1, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('%').whenNumberGreaterThan(0)
        .setBackground(COLORS.POSITIVE.background)
        .setFontColor(COLORS.POSITIVE.fontColor)
        .setRanges([profitRange]).build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('%').whenNumberLessThan(0)
        .setBackground(COLORS.NEGATIVE.background)
        .setFontColor(COLORS.NEGATIVE.fontColor)
        .setRanges([profitRange]).build()
    );

    const growthColumn = 18;
    const growthRange = sheet.getRange(2, growthColumn, numRows - 1, 1);
    const statusColors = {
      "🟢 Healthy Growth": { background: "#d4edda", fontColor: "#155724" },
      "🟢 Efficiency Improvement": { background: "#d1f2eb", fontColor: "#0c5460" },
      "🔴 Inefficient Growth": { background: "#f8d7da", fontColor: "#721c24" },
      "🟠 Declining Efficiency": { background: "#ff9800", fontColor: "white" },
      "🔵 Scaling Down": { background: "#cce7ff", fontColor: "#004085" },
      "🔵 Scaling Down - Efficient": { background: "#b8e6b8", fontColor: "#2d5a2d" },
      "🔵 Scaling Down - Moderate": { background: "#d1ecf1", fontColor: "#0c5460" },
      "🔵 Scaling Down - Problematic": { background: "#ffcc99", fontColor: "#cc5500" },
      "🟡 Moderate Growth": { background: "#fff3cd", fontColor: "#856404" },
      "🟡 Moderate Decline - Efficiency Drop": { background: "#ffe0cc", fontColor: "#cc6600" },
      "🟡 Moderate Decline - Spend Optimization": { background: "#e6f3ff", fontColor: "#0066cc" },
      "🟡 Moderate Decline - Proportional": { background: "#f0f0f0", fontColor: "#666666" },
      "🟡 Efficiency Improvement": { background: "#e8f5e8", fontColor: "#2d5a2d" },
      "🟡 Minimal Growth": { background: "#fff8e1", fontColor: "#f57f17" },
      "🟡 Moderate Decline": { background: "#fff3cd", fontColor: "#856404" },
      "⚪ Stable": { background: "#f5f5f5", fontColor: "#616161" },
      "First Week": { background: "#e0e0e0", fontColor: "#757575" }
    };

    Object.entries(statusColors).forEach(([status, colors]) => {
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains(status)
          .setBackground(colors.background)
          .setFontColor(colors.fontColor)
          .setRanges([growthRange]).build()
      );
    });
  }
  
  sheet.setConditionalFormatRules(rules);
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

function createProjectPivotTable(projectName, appData) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    createUnifiedPivotTable(appData);
  } finally {
    setCurrentProject(originalProject);
  }
}