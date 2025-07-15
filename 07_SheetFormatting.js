function createEnhancedPivotTable(appData) {
  if (CURRENT_PROJECT === 'OVERALL' && Object.keys(appData).length > 10) {
    createProgressivePivotTable(appData);
  } else {
    createStandardPivotTable(appData);
  }
}

function createProgressivePivotTable(appData) {
  console.log('createProgressivePivotTable: start');
  const props = PropertiesService.getScriptProperties();
  const progressKey = `${CURRENT_PROJECT}_PROGRESS`;
  const dataKey = `${CURRENT_PROJECT}_DATA`;
  
  let progress = props.getProperty(progressKey);
  let savedData = props.getProperty(dataKey);
  
  if (!progress || !savedData) {
    console.log('createProgressivePivotTable: initializing new generation');
    const wow = calculateWoWMetrics(appData);
    const state = {
      project: CURRENT_PROJECT,
      appData: appData,
      wow: wow,
      currentAppIndex: 0,
      currentRow: 2,
      formatData: [],
      groupData: [],
      stage: 'DATA'
    };
    props.setProperty(progressKey, JSON.stringify(state));
    props.setProperty(dataKey, JSON.stringify(appData));
  }
  
  const state = JSON.parse(props.getProperty(progressKey));
  
  switch (state.stage) {
    case 'DATA':
      writeDataProgressive(state);
      break;
    case 'FORMAT':
      applyFormattingProgressive(state);
      break;
    case 'GROUPS':
      applyGroupingProgressive(state);
      break;
    case 'COMPLETE':
      console.log('createProgressivePivotTable: already complete');
      clearProgress(state.project);
      break;
  }
}

function writeDataProgressive(state) {
  console.log('writeDataProgressive: start');
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);
  
  const headers = getUnifiedHeaders();
  
  if (state.currentRow === 2) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    TABLE_CONFIG.COLUMN_WIDTHS.forEach(col => sheet.setColumnWidth(col.c, col.w));
  }
  
  const appKeys = Object.keys(state.appData).sort((a, b) => state.appData[a].appName.localeCompare(state.appData[b].appName));
  const ROWS_PER_EXECUTION = 150;
  let rowsWritten = 0;
  let tableData = [];
  
  const startTime = new Date().getTime();
  const MAX_EXECUTION_TIME = 50000;
  
  for (let i = state.currentAppIndex; i < appKeys.length; i++) {
    const appKey = appKeys[i];
    const app = state.appData[appKey];
    
    const appRowNum = state.currentRow + tableData.length;
    state.formatData.push({ row: appRowNum, type: 'APP' });
    tableData.push(['APP', app.appName, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    
    const weekKeys = Object.keys(app.weeks).sort();
    
    for (const weekKey of weekKeys) {
      const week = app.weeks[weekKey];
      const weekRowNum = state.currentRow + tableData.length;
      state.formatData.push({ row: weekRowNum, type: 'WEEK' });
      
      if (state.project === 'OVERALL' && week.networks) {
        const allNetworks = Object.values(week.networks || {});
        const weekTotals = calculateNetworkTotals(allNetworks);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = state.wow.appWeekWoW[appWeekKey] || {};
        
        tableData.push(createWeekRow(week, weekTotals, weekWoW.spendChangePercent ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '', weekWoW.eProfitChangePercent ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '', weekWoW.growthStatus || ''));
        
        const networkKeys = Object.keys(week.networks || {}).sort((a, b) => (week.networks[b].spend || 0) - (week.networks[a].spend || 0));
        
        for (const networkKey of networkKeys) {
          const network = week.networks[networkKey];
          const networkWoWKey = `${network.networkId}_${weekKey}`;
          const networkWoW = state.wow.networkWoW[networkWoWKey] || {};
          
          state.formatData.push({ row: state.currentRow + tableData.length, type: 'NETWORK' });
          tableData.push(createNetworkRow(network, networkWoW.spendChangePercent ? `${networkWoW.spendChangePercent.toFixed(0)}%` : '', networkWoW.eProfitChangePercent ? `${networkWoW.eProfitChangePercent.toFixed(0)}%` : '', networkWoW.growthStatus || ''));
        }
        
        if (networkKeys.length > 0) {
          state.groupData.push({
            type: 'week_group',
            startRow: weekRowNum + 1,
            count: networkKeys.length
          });
        }
        
      } else if (state.project === 'TRICKY' && week.sourceApps) {
        const allCampaigns = [];
        Object.values(week.sourceApps || {}).forEach(sourceApp => {
          allCampaigns.push(...sourceApp.campaigns);
        });
        
        const weekTotals = calculateWeekTotals(allCampaigns);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = state.wow.appWeekWoW[appWeekKey] || {};
        
        tableData.push(createWeekRow(week, weekTotals, weekWoW.spendChangePercent ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '', weekWoW.eProfitChangePercent ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '', weekWoW.growthStatus || ''));
        
        const sourceAppKeys = Object.keys(week.sourceApps).sort((a, b) => {
          const totalSpendA = week.sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
          const totalSpendB = week.sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
          return totalSpendB - totalSpendA;
        });
        
        let weekContentRows = 0;
        
        sourceAppKeys.forEach(sourceAppKey => {
          const sourceApp = week.sourceApps[sourceAppKey];
          const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
          const sourceAppWoWKey = `${sourceApp.sourceAppId}_${weekKey}`;
          const sourceAppWoW = state.wow.sourceAppWoW[sourceAppWoWKey] || {};
          
          const sourceAppRowNum = state.currentRow + tableData.length;
          state.formatData.push({ row: sourceAppRowNum, type: 'SOURCE_APP' });
          
          tableData.push(createSourceAppRow(sourceApp.sourceAppName, sourceAppTotals, sourceAppWoW.spendChangePercent ? `${sourceAppWoW.spendChangePercent.toFixed(0)}%` : '', sourceAppWoW.eProfitChangePercent ? `${sourceAppWoW.eProfitChangePercent.toFixed(0)}%` : '', sourceAppWoW.growthStatus || ''));
          
          sourceApp.campaigns.sort((a, b) => b.spend - a.spend).forEach(campaign => {
            const campaignIdValue = state.project === 'TRICKY' || state.project === 'REGULAR' 
              ? `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`
              : campaign.campaignId;
            
            const key = `${campaign.campaignId}_${weekKey}`;
            const campaignWoW = state.wow.campaignWoW[key] || {};
            
            state.formatData.push({ row: state.currentRow + tableData.length, type: 'CAMPAIGN' });
            tableData.push(createCampaignRow(campaign, campaignIdValue, campaignWoW.spendChangePercent ? `${campaignWoW.spendChangePercent.toFixed(0)}%` : '', campaignWoW.eProfitChangePercent ? `${campaignWoW.eProfitChangePercent.toFixed(0)}%` : '', campaignWoW.growthStatus || ''));
          });
          
          if (sourceApp.campaigns.length > 0) {
            state.groupData.push({
              type: 'campaign_group',
              startRow: sourceAppRowNum + 1,
              count: sourceApp.campaigns.length
            });
          }
          
          weekContentRows += 1 + sourceApp.campaigns.length;
        });
        
        if (weekContentRows > 0) {
          state.groupData.push({
            type: 'week_group',
            startRow: weekRowNum + 1,
            count: weekContentRows
          });
        }
        
      } else {
        const weekTotals = calculateWeekTotals(week.campaigns || []);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = state.wow.appWeekWoW[appWeekKey] || {};
        
        tableData.push(createWeekRow(week, weekTotals, weekWoW.spendChangePercent ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '', weekWoW.eProfitChangePercent ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '', weekWoW.growthStatus || ''));
        
        if (week.campaigns && week.campaigns.length > 0) {
          week.campaigns.sort((a, b) => b.spend - a.spend).forEach(campaign => {
            const campaignIdValue = state.project === 'TRICKY' || state.project === 'REGULAR' 
              ? `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`
              : campaign.campaignId;
            
            const key = `${campaign.campaignId}_${weekKey}`;
            const campaignWoW = state.wow.campaignWoW[key] || {};
            
            state.formatData.push({ row: state.currentRow + tableData.length, type: 'CAMPAIGN' });
            tableData.push(createCampaignRow(campaign, campaignIdValue, campaignWoW.spendChangePercent ? `${campaignWoW.spendChangePercent.toFixed(0)}%` : '', campaignWoW.eProfitChangePercent ? `${campaignWoW.eProfitChangePercent.toFixed(0)}%` : '', campaignWoW.growthStatus || ''));
          });
          
          state.groupData.push({
            type: 'week_group',
            startRow: weekRowNum + 1,
            count: week.campaigns.length
          });
        }
      }
    }
    
    const appContentRows = tableData.length - (appRowNum - state.currentRow + 1);
    if (appContentRows > 0) {
      state.groupData.push({
        type: 'app_group',
        startRow: appRowNum + 1,
        count: appContentRows
      });
    }
    
    rowsWritten = tableData.length;
    state.currentAppIndex = i + 1;
    
    if (rowsWritten >= ROWS_PER_EXECUTION || (new Date().getTime() - startTime) > MAX_EXECUTION_TIME) {
      console.log(`writeDataProgressive: writing batch of ${rowsWritten} rows`);
      break;
    }
  }
  
  if (tableData.length > 0) {
    sheet.getRange(state.currentRow, 1, tableData.length, headers.length).setValues(tableData);
    state.currentRow += tableData.length;
  }
  
  if (state.currentAppIndex >= appKeys.length) {
    console.log('writeDataProgressive: all data written, moving to format stage');
    state.stage = 'FORMAT';
    state.totalRows = state.currentRow - 1;
    state.formatIndex = 0;
  }
  
  PropertiesService.getScriptProperties().setProperty(`${state.project}_PROGRESS`, JSON.stringify(state));
  
  if (state.stage === 'DATA') {
    console.log('writeDataProgressive: scheduling next batch');
    Utilities.sleep(1000);
    writeDataProgressive(state);
  } else {
    applyFormattingProgressive(state);
  }
}

function applyFormattingProgressive(state) {
  console.log('applyFormattingProgressive: start');
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  const numRows = state.totalRows;
  const numCols = 19;
  
  const startTime = new Date().getTime();
  const MAX_EXECUTION_TIME = 45000;
  
  if (!state.formatIndex || state.formatIndex === 0) {
    console.log('applyFormattingProgressive: applying header and basic formatting');
    const headerRange = sheet.getRange(1, 1, 1, numCols);
    headerRange.setBackground(COLORS.HEADER.background).setFontColor(COLORS.HEADER.fontColor).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    
    if (numRows > 1) {
      sheet.getRange(2, 1, numRows - 1, numCols).setVerticalAlignment('middle');
      sheet.getRange(2, numCols, numRows - 1, 1).setWrap(true).setHorizontalAlignment('left');
      sheet.getRange(2, numCols - 1, numRows - 1, 1).setWrap(true).setHorizontalAlignment('left');
    }
    
    state.formatIndex = 1;
  }
  
  const formatBatch = 50;
  let batchesProcessed = 0;
  
  for (let i = state.formatIndex - 1; i < state.formatData.length; i += formatBatch) {
    const batch = state.formatData.slice(i, i + formatBatch);
    
    batch.forEach(item => {
      if (item.row <= numRows + 1) {
        try {
          if (item.type === 'APP') {
            sheet.getRange(item.row, 1, 1, numCols).setBackground(COLORS.APP_ROW.background).setFontWeight('bold').setFontSize(10);
          } else if (item.type === 'WEEK') {
            sheet.getRange(item.row, 1, 1, numCols).setBackground(COLORS.WEEK_ROW.background);
          } else if (item.type === 'NETWORK') {
            sheet.getRange(item.row, 1, 1, numCols).setBackground(COLORS.NETWORK_ROW.background).setFontSize(9);
          } else if (item.type === 'SOURCE_APP') {
            sheet.getRange(item.row, 1, 1, numCols).setBackground(COLORS.SOURCE_APP_ROW.background).setFontSize(9);
          } else if (item.type === 'CAMPAIGN') {
            sheet.getRange(item.row, 1, 1, numCols).setBackground(COLORS.CAMPAIGN_ROW.background).setFontSize(9);
          }
        } catch (e) {}
      }
    });
    
    state.formatIndex = i + formatBatch + 1;
    batchesProcessed++;
    
    if ((new Date().getTime() - startTime) > MAX_EXECUTION_TIME) {
      console.log(`applyFormattingProgressive: processed ${batchesProcessed} batches, pausing`);
      break;
    }
  }
  
  if (state.formatIndex >= state.formatData.length) {
    console.log('applyFormattingProgressive: applying number formats');
    
    if (numRows > 1) {
      sheet.getRange(2, 5, numRows - 1, 1).setNumberFormat('$0.00');
      sheet.getRange(2, 8, numRows - 1, 1).setNumberFormat('$0.000');
      sheet.getRange(2, 9, numRows - 1, 1).setNumberFormat('0.00');
      sheet.getRange(2, 10, numRows - 1, 1).setNumberFormat('0.0');
      sheet.getRange(2, 13, numRows - 1, 1).setNumberFormat('$0.000');
      sheet.getRange(2, 16, numRows - 1, 1).setNumberFormat('$0.00');
    }
    
    applyConditionalFormattingProgressive(sheet, numRows);
    sheet.hideColumns(1);
    
    state.stage = 'GROUPS';
    state.groupIndex = 0;
  }
  
  PropertiesService.getScriptProperties().setProperty(`${state.project}_PROGRESS`, JSON.stringify(state));
  
  if (state.stage === 'FORMAT') {
    console.log('applyFormattingProgressive: scheduling next batch');
    Utilities.sleep(1000);
    applyFormattingProgressive(state);
  } else {
    applyGroupingProgressive(state);
  }
}

function applyGroupingProgressive(state) {
  console.log('applyGroupingProgressive: start');
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  const startTime = new Date().getTime();
  const MAX_EXECUTION_TIME = 45000;
  const GROUPS_PER_BATCH = 10;
  
  let groupsProcessed = 0;
  
  for (let i = state.groupIndex || 0; i < state.groupData.length; i++) {
    const group = state.groupData[i];
    
    try {
      sheet.getRange(group.startRow, 1, group.count, 1).shiftRowGroupDepth(1);
      SpreadsheetApp.flush();
      Utilities.sleep(100);
      
      sheet.getRange(group.startRow, 1, group.count, 1).collapseGroups();
      SpreadsheetApp.flush();
      Utilities.sleep(100);
    } catch (e) {
      console.error(`Error applying group ${i}:`, e);
    }
    
    groupsProcessed++;
    state.groupIndex = i + 1;
    
    if (groupsProcessed >= GROUPS_PER_BATCH || (new Date().getTime() - startTime) > MAX_EXECUTION_TIME) {
      console.log(`applyGroupingProgressive: processed ${groupsProcessed} groups, pausing`);
      break;
    }
  }
  
  if (state.groupIndex >= state.groupData.length) {
    console.log('applyGroupingProgressive: all groups applied');
    state.stage = 'COMPLETE';
  }
  
  PropertiesService.getScriptProperties().setProperty(`${state.project}_PROGRESS`, JSON.stringify(state));
  
  if (state.stage === 'GROUPS') {
    console.log('applyGroupingProgressive: scheduling next batch');
    Utilities.sleep(1000);
    applyGroupingProgressive(state);
  } else {
    console.log('applyGroupingProgressive: complete');
    clearProgress(state.project);
    
    const cache = new CommentCache();
    cache.applyCommentsToSheet();
  }
}

function clearProgress(projectName) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(`${projectName}_PROGRESS`);
  props.deleteProperty(`${projectName}_DATA`);
}

function createStandardPivotTable(appData) {
  console.log('createStandardPivotTable: start');
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);

  const wow = calculateWoWMetrics(appData);
  const headers = getUnifiedHeaders();
  const tableData = [headers];
  const formatData = [];

  const appKeys = Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName));
  
  appKeys.forEach(appKey => {
    const app = appData[appKey];
    
    formatData.push({ row: tableData.length + 1, type: 'APP' });
    tableData.push(['APP', app.appName, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);

    const weekKeys = Object.keys(app.weeks).sort();
    
    weekKeys.forEach(weekKey => {
      const week = app.weeks[weekKey];
      
      formatData.push({ row: tableData.length + 1, type: 'WEEK' });
      
      if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
        const allNetworks = Object.values(week.networks || {});
        const weekTotals = calculateNetworkTotals(allNetworks);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = wow.appWeekWoW[appWeekKey] || {};
        
        tableData.push(createWeekRow(week, weekTotals, weekWoW.spendChangePercent ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '', weekWoW.eProfitChangePercent ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '', weekWoW.growthStatus || ''));
        
        addNetworkRows(tableData, week.networks, weekKey, wow, formatData);
        
      } else if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
        const allCampaigns = [];
        Object.values(week.sourceApps || {}).forEach(sourceApp => {
          allCampaigns.push(...sourceApp.campaigns);
        });
        
        const weekTotals = calculateWeekTotals(allCampaigns);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = wow.appWeekWoW[appWeekKey] || {};
        
        tableData.push(createWeekRow(week, weekTotals, weekWoW.spendChangePercent ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '', weekWoW.eProfitChangePercent ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '', weekWoW.growthStatus || ''));
        
        addSourceAppRows(tableData, week.sourceApps, weekKey, wow, formatData);
        
      } else {
        const weekTotals = calculateWeekTotals(week.campaigns || []);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = wow.appWeekWoW[appWeekKey] || {};
        
        tableData.push(createWeekRow(week, weekTotals, weekWoW.spendChangePercent ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '', weekWoW.eProfitChangePercent ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '', weekWoW.growthStatus || ''));
        
        addCampaignRows(tableData, week.campaigns || [], week, weekKey, wow, formatData);
      }
    });
  });

  sheet.getRange(1, 1, tableData.length, headers.length).setValues(tableData);
  applyStandardFormatting(sheet, tableData.length, headers.length, formatData, appData);
  createStandardRowGrouping(sheet, tableData, appData);
  sheet.setFrozenRows(1);
}

function applyStandardFormatting(sheet, numRows, numCols, formatData, appData) {
  applyBasicFormatting(sheet, numRows, numCols);
  applyRowFormatting(sheet, formatData, numCols);
  applyNumberFormats(sheet, numRows);
  applyConditionalFormatting(sheet, numRows, appData);
  sheet.hideColumns(1);
}

function applyConditionalFormattingProgressive(sheet, numRows) {
  const rules = [];
  
  if (numRows > 1) {
    const spendRange = sheet.getRange(2, 6, numRows - 1, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule().whenTextContains('%').whenNumberGreaterThan(0).setBackground(COLORS.POSITIVE.background).setFontColor(COLORS.POSITIVE.fontColor).setRanges([spendRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextContains('%').whenNumberLessThan(0).setBackground(COLORS.NEGATIVE.background).setFontColor(COLORS.NEGATIVE.fontColor).setRanges([spendRange]).build()
    );

    const profitColumn = 17;
    const profitRange = sheet.getRange(2, profitColumn, numRows - 1, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule().whenTextContains('%').whenNumberGreaterThan(0).setBackground(COLORS.POSITIVE.background).setFontColor(COLORS.POSITIVE.fontColor).setRanges([profitRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextContains('%').whenNumberLessThan(0).setBackground(COLORS.NEGATIVE.background).setFontColor(COLORS.NEGATIVE.fontColor).setRanges([profitRange]).build()
    );
  }
  
  sheet.setConditionalFormatRules(rules);
}

function addNetworkRows(tableData, networks, weekKey, wow, formatData) {
  if (!networks) return;
  
  const networkKeys = Object.keys(networks).sort((a, b) => networks[b].spend - networks[a].spend);
  
  networkKeys.forEach(networkKey => {
    const network = networks[networkKey];
    
    const networkWoWKey = `${network.networkId}_${weekKey}`;
    const networkWoW = wow.networkWoW[networkWoWKey] || {};
    
    formatData.push({ row: tableData.length + 1, type: 'NETWORK' });
    
    tableData.push(createNetworkRow(network, networkWoW.spendChangePercent ? `${networkWoW.spendChangePercent.toFixed(0)}%` : '', networkWoW.eProfitChangePercent ? `${networkWoW.eProfitChangePercent.toFixed(0)}%` : '', networkWoW.growthStatus || ''));
  });
}

function addSourceAppRows(tableData, sourceApps, weekKey, wow, formatData) {
  const sourceAppKeys = Object.keys(sourceApps).sort((a, b) => {
    const totalSpendA = sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
    const totalSpendB = sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
    return totalSpendB - totalSpendA;
  });
  
  sourceAppKeys.forEach(sourceAppKey => {
    const sourceApp = sourceApps[sourceAppKey];
    const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
    
    const sourceAppWoWKey = `${sourceApp.sourceAppId}_${weekKey}`;
    const sourceAppWoW = wow.sourceAppWoW[sourceAppWoWKey] || {};
    
    formatData.push({ row: tableData.length + 1, type: 'SOURCE_APP' });
    tableData.push(createSourceAppRow(sourceApp.sourceAppName, sourceAppTotals, sourceAppWoW.spendChangePercent ? `${sourceAppWoW.spendChangePercent.toFixed(0)}%` : '', sourceAppWoW.eProfitChangePercent ? `${sourceAppWoW.eProfitChangePercent.toFixed(0)}%` : '', sourceAppWoW.growthStatus || ''));
    
    addCampaignRows(tableData, sourceApp.campaigns, { weekStart: weekKey }, weekKey, wow, formatData);
  });
}

function createNetworkRow(network, spendWoW, profitWoW, status) {
  const installs = network.installs || 0;
  const spend = network.spend || 0;
  const cpi = installs > 0 ? spend / installs : 0;
  const rrD1 = network.rrD1 || 0;
  const roas = network.roas || 0;
  const rrD7 = network.rrD7 || 0;
  const ipm = network.ipm || 0;
  const eArpuForecast = network.eArpuForecast || 0;
  const eRoasForecast = network.eRoasForecast || 0;
  const eRoasForecastD730 = network.eRoasForecastD730 || 0;
  const eProfitForecast = network.eProfitForecast || 0;
  
  return [
    'NETWORK', network.networkName, '', 'ALL',
    spend.toFixed(2), spendWoW, installs, cpi.toFixed(3),
    roas.toFixed(2), ipm.toFixed(1), `${rrD1.toFixed(1)}%`, `${rrD7.toFixed(1)}%`,
    eArpuForecast.toFixed(3), `${eRoasForecast.toFixed(0)}%`, `${eRoasForecastD730.toFixed(0)}%`,
    eProfitForecast.toFixed(2), profitWoW, status, ''
  ];
}

function calculateNetworkTotals(networks) {
  const totalSpend = networks.reduce((s, n) => s + (n.spend || 0), 0);
  const totalInstalls = networks.reduce((s, n) => s + (n.installs || 0), 0);
  const avgCpi = totalInstalls ? totalSpend / totalInstalls : 0;
  const avgRoas = networks.length ? networks.reduce((s, n) => s + (n.roas || 0), 0) / networks.length : 0;
  const avgRrD1 = networks.length ? networks.reduce((s, n) => s + (n.rrD1 || 0), 0) / networks.length : 0;
  const avgRrD7 = networks.length ? networks.reduce((s, n) => s + (n.rrD7 || 0), 0) / networks.length : 0;
  const avgIpm = networks.length ? networks.reduce((s, n) => s + (n.ipm || 0), 0) / networks.length : 0;
  const avgArpu = networks.length ? networks.reduce((s, n) => s + (n.eArpuForecast || 0), 0) / networks.length : 0;
  
  const validForEROAS = networks.filter(n => n.eRoasForecast >= 1 && n.eRoasForecast <= 1000 && n.spend > 0);
  
  let avgERoas = 0;
  if (validForEROAS.length > 0) {
    const totalWeightedEROAS = validForEROAS.reduce((sum, n) => sum + (n.eRoasForecast * n.spend), 0);
    const totalSpendForEROAS = validForEROAS.reduce((sum, n) => sum + n.spend, 0);
    avgERoas = totalSpendForEROAS > 0 ? totalWeightedEROAS / totalSpendForEROAS : 0;
  }
  
  const validForEROASD730 = networks.filter(n => n.eRoasForecastD730 >= 1 && n.eRoasForecastD730 <= 1000 && n.spend > 0);
  
  let avgEROASD730 = 0;
  if (validForEROASD730.length > 0) {
    const totalWeightedEROASD730 = validForEROASD730.reduce((sum, n) => sum + (n.eRoasForecastD730 * n.spend), 0);
    const totalSpendForEROASD730 = validForEROASD730.reduce((sum, n) => sum + n.spend, 0);
    avgEROASD730 = totalSpendForEROASD730 > 0 ? totalWeightedEROASD730 / totalSpendForEROASD730 : 0;
  }
  
  const totalProfit = networks.reduce((s, n) => s + (n.eProfitForecast || 0), 0);

  return {
    totalSpend, totalInstalls, avgCpi, avgRoas, avgIpm, avgRrD1, avgRrD7,
    avgArpu, avgERoas, avgEROASD730, totalProfit
  };
}

function createSourceAppRow(sourceAppDisplayName, totals, spendWoW, profitWoW, status) {
  return [
    'SOURCE_APP', sourceAppDisplayName, '', '',
    totals.totalSpend.toFixed(2), spendWoW, totals.totalInstalls, totals.avgCpi.toFixed(3),
    totals.avgRoas.toFixed(2), totals.avgIpm.toFixed(1), `${totals.avgRrD1.toFixed(1)}%`, `${totals.avgRrD7.toFixed(1)}%`,
    totals.avgArpu.toFixed(3), `${totals.avgERoas.toFixed(0)}%`, `${totals.avgEROASD730.toFixed(0)}%`,
    totals.totalProfit.toFixed(2), profitWoW, status, ''
  ];
}

function getUnifiedHeaders() {
  return ['Level','Week Range / Source App','ID','GEO','Spend','Spend WoW %','Installs','CPI','ROAS D-1','IPM','RR D-1','RR D-7','eARPU 365d','eROAS 365d','eROAS 730d','eProfit 730d','eProfit 730d WoW %','Growth Status','Comments'];
}

function createWeekRow(week, weekTotals, spendWoW, profitWoW, status) {
  return [
    'WEEK', `${week.weekStart} - ${week.weekEnd}`, '', '',
    weekTotals.totalSpend.toFixed(2), spendWoW, weekTotals.totalInstalls, weekTotals.avgCpi.toFixed(3),
    weekTotals.avgRoas.toFixed(2), weekTotals.avgIpm.toFixed(1), `${weekTotals.avgRrD1.toFixed(1)}%`, `${weekTotals.avgRrD7.toFixed(1)}%`,
    weekTotals.avgArpu.toFixed(3), `${weekTotals.avgERoas.toFixed(0)}%`, `${weekTotals.avgEROASD730.toFixed(0)}%`,
    weekTotals.totalProfit.toFixed(2), profitWoW, status, ''
  ];
}

function applyBasicFormatting(sheet, numRows, numCols) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setBackground(COLORS.HEADER.background).setFontColor(COLORS.HEADER.fontColor).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(10).setWrap(true);

  TABLE_CONFIG.COLUMN_WIDTHS.forEach(col => sheet.setColumnWidth(col.c, col.w));

  if (numRows > 1) {
    sheet.getRange(2, 1, numRows - 1, numCols).setVerticalAlignment('middle');
    sheet.getRange(2, numCols, numRows - 1, 1).setWrap(true).setHorizontalAlignment('left');
    sheet.getRange(2, numCols - 1, numRows - 1, 1).setWrap(true).setHorizontalAlignment('left');
  }
}

function applyRowFormatting(sheet, formatData, numCols) {
  const rowsByType = { app: [], week: [], network: [], sourceApp: [], campaign: [] };
  
  formatData.forEach(item => {
    if (item.type === 'APP') rowsByType.app.push(item.row);
    else if (item.type === 'WEEK') rowsByType.week.push(item.row);
    else if (item.type === 'NETWORK') rowsByType.network.push(item.row);
    else if (item.type === 'SOURCE_APP') rowsByType.sourceApp.push(item.row);
    else if (item.type === 'CAMPAIGN') rowsByType.campaign.push(item.row);
  });

  const applyFormat = (rows, color, size = 10, weight = 'normal') => {
    rows.forEach(r => sheet.getRange(r, 1, 1, numCols).setBackground(color.background).setFontColor(color.fontColor || 'black').setFontWeight(weight).setFontSize(size));
  };

  applyFormat(rowsByType.app, COLORS.APP_ROW, 10, 'bold');
  applyFormat(rowsByType.week, COLORS.WEEK_ROW);
  applyFormat(rowsByType.network, COLORS.NETWORK_ROW, 9);
  applyFormat(rowsByType.sourceApp, COLORS.SOURCE_APP_ROW, 9);
  applyFormat(rowsByType.campaign, COLORS.CAMPAIGN_ROW, 9);
}

function applyNumberFormats(sheet, numRows) {
  if (numRows > 1) {
    sheet.getRange(2, 5, numRows - 1, 1).setNumberFormat('$0.00');
    sheet.getRange(2, 8, numRows - 1, 1).setNumberFormat('$0.000');
    sheet.getRange(2, 9, numRows - 1, 1).setNumberFormat('0.00');
    sheet.getRange(2, 10, numRows - 1, 1).setNumberFormat('0.0');
    sheet.getRange(2, 13, numRows - 1, 1).setNumberFormat('$0.000');
    sheet.getRange(2, 16, numRows - 1, 1).setNumberFormat('$0.00');
  }
}

function applyConditionalFormatting(sheet, numRows, appData) {
  const rules = [];
  
  if (numRows > 1) {
    const spendRange = sheet.getRange(2, 6, numRows - 1, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule().whenTextContains('%').whenNumberGreaterThan(0).setBackground(COLORS.POSITIVE.background).setFontColor(COLORS.POSITIVE.fontColor).setRanges([spendRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextContains('%').whenNumberLessThan(0).setBackground(COLORS.NEGATIVE.background).setFontColor(COLORS.NEGATIVE.fontColor).setRanges([spendRange]).build()
    );

    const eroasColumn = 15;
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
      const colLetter = String.fromCharCode(64 + eroasColumn);
      
      rules.push(
        SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(NOT(ISBLANK(${colLetter}${i + 1})), VALUE(SUBSTITUTE(${colLetter}${i + 1},"%","")) >= ${targetEROAS})`).setBackground(COLORS.POSITIVE.background).setFontColor(COLORS.POSITIVE.fontColor).setRanges([cellRange]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(NOT(ISBLANK(${colLetter}${i + 1})), VALUE(SUBSTITUTE(${colLetter}${i + 1},"%","")) >= 120, VALUE(SUBSTITUTE(${colLetter}${i + 1},"%","")) < ${targetEROAS})`).setBackground(COLORS.WARNING.background).setFontColor(COLORS.WARNING.fontColor).setRanges([cellRange]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(NOT(ISBLANK(${colLetter}${i + 1})), VALUE(SUBSTITUTE(${colLetter}${i + 1},"%","")) < 120)`).setBackground(COLORS.NEGATIVE.background).setFontColor(COLORS.NEGATIVE.fontColor).setRanges([cellRange]).build()
      );
    }

    const profitColumn = 17;
    const profitRange = sheet.getRange(2, profitColumn, numRows - 1, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule().whenTextContains('%').whenNumberGreaterThan(0).setBackground(COLORS.POSITIVE.background).setFontColor(COLORS.POSITIVE.fontColor).setRanges([profitRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextContains('%').whenNumberLessThan(0).setBackground(COLORS.NEGATIVE.background).setFontColor(COLORS.NEGATIVE.fontColor).setRanges([profitRange]).build()
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
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains(status).setBackground(colors.background).setFontColor(colors.fontColor).setRanges([growthRange]).build());
    });
  }
  
  sheet.setConditionalFormatRules(rules);
}

function calculateWeekTotals(campaigns) {
  const totalSpend = campaigns.reduce((s, c) => s + c.spend, 0);
  const totalInstalls = campaigns.reduce((s, c) => s + c.installs, 0);
  const avgCpi = totalInstalls ? totalSpend / totalInstalls : 0;
  const avgRoas = campaigns.length ? campaigns.reduce((s, c) => s + c.roas, 0) / campaigns.length : 0;
  const avgIpm = campaigns.length ? campaigns.reduce((s, c) => s + c.ipm, 0) / campaigns.length : 0;
  const avgRrD1 = campaigns.length ? campaigns.reduce((s, c) => s + c.rrD1, 0) / campaigns.length : 0;
  const avgRrD7 = campaigns.length ? campaigns.reduce((s, c) => s + c.rrD7, 0) / campaigns.length : 0;
  const avgArpu = campaigns.length ? campaigns.reduce((s, c) => s + c.eArpuForecast, 0) / campaigns.length : 0;
  
  const validForEROAS = campaigns.filter(c => c.eRoasForecast >= 1 && c.eRoasForecast <= 1000 && c.spend > 0);
  
  let avgERoas = 0;
  if (validForEROAS.length > 0) {
    const totalWeightedEROAS = validForEROAS.reduce((sum, c) => sum + (c.eRoasForecast * c.spend), 0);
    const totalSpendForEROAS = validForEROAS.reduce((sum, c) => sum + c.spend, 0);
    avgERoas = totalSpendForEROAS > 0 ? totalWeightedEROAS / totalSpendForEROAS : 0;
  }
  
  const validForEROASD730 = campaigns.filter(c => c.eRoasForecastD730 >= 1 && c.eRoasForecastD730 <= 1000 && c.spend > 0);
  
  let avgEROASD730 = 0;
  if (validForEROASD730.length > 0) {
    const totalWeightedEROASD730 = validForEROASD730.reduce((sum, c) => sum + (c.eRoasForecastD730 * c.spend), 0);
    const totalSpendForEROASD730 = validForEROASD730.reduce((sum, c) => sum + c.spend, 0);
    avgEROASD730 = totalSpendForEROASD730 > 0 ? totalWeightedEROASD730 / totalSpendForEROASD730 : 0;
  }
  
  const totalProfit = campaigns.reduce((s, c) => s + c.eProfitForecast, 0);

  return {
    totalSpend, totalInstalls, avgCpi, avgRoas, avgIpm, avgRrD1, avgRrD7,
    avgArpu, avgERoas, avgEROASD730, totalProfit
  };
}

function addCampaignRows(tableData, campaigns, week, weekKey, wow, formatData) {
  if (CURRENT_PROJECT === 'OVERALL') return;
  
  campaigns.sort((a, b) => b.spend - a.spend).forEach(campaign => {
    let campaignIdValue;
    if (CURRENT_PROJECT === 'TRICKY' || CURRENT_PROJECT === 'REGULAR') {
      campaignIdValue = `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`;
    } else {
      campaignIdValue = campaign.campaignId;
    }
    
    const key = `${campaign.campaignId}_${weekKey}`;
    const campaignWoW = wow.campaignWoW[key] || {};
    
    formatData.push({ row: tableData.length + 1, type: 'CAMPAIGN' });
    
    tableData.push(createCampaignRow(campaign, campaignIdValue, campaignWoW.spendChangePercent ? `${campaignWoW.spendChangePercent.toFixed(0)}%` : '', campaignWoW.eProfitChangePercent ? `${campaignWoW.eProfitChangePercent.toFixed(0)}%` : '', campaignWoW.growthStatus || ''));
  });
}

function createCampaignRow(campaign, campaignIdValue, spendPct, profitPct, growthStatus) {
  return [
    'CAMPAIGN', campaign.sourceApp, campaignIdValue, campaign.geo,
    campaign.spend.toFixed(2), spendPct, campaign.installs, campaign.cpi ? campaign.cpi.toFixed(3) : '0.000',
    campaign.roas.toFixed(2), campaign.ipm.toFixed(1), `${campaign.rrD1.toFixed(1)}%`, `${campaign.rrD7.toFixed(1)}%`,
    campaign.eArpuForecast.toFixed(3), `${campaign.eRoasForecast.toFixed(0)}%`, `${campaign.eRoasForecastD730.toFixed(0)}%`,
    campaign.eProfitForecast.toFixed(2), profitPct, growthStatus, ''
  ];
}

function createStandardRowGrouping(sheet, tableData, appData) {
  try {
    let rowPointer = 2;
    const sortedApps = Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName));
    const groupOperations = [];

    sortedApps.forEach(appKey => {
      const app = appData[appKey];
      const appStartRow = rowPointer;
      rowPointer++;

      const sortedWeeks = Object.keys(app.weeks).sort();
      
      sortedWeeks.forEach(weekKey => {
        const week = app.weeks[weekKey];
        const weekStartRow = rowPointer;
        rowPointer++;

        let weekContentRows = 0;

        if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
          const networkCount = Object.keys(week.networks).length;
          rowPointer += networkCount;
          weekContentRows = networkCount;
          
          if (networkCount > 0) {
            groupOperations.push({
              type: 'week_group',
              startRow: weekStartRow + 1,
              count: networkCount
            });
          }
          
        } else if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
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
              groupOperations.push({
                type: 'campaign_group',
                startRow: sourceAppStartRow + 1,
                count: campaignCount
              });
            }
          });
          
          if (weekContentRows > 0) {
            groupOperations.push({
              type: 'week_group',
              startRow: weekStartRow + 1,
              count: weekContentRows
            });
          }
          
        } else if (week.campaigns) {
          const campaignCount = week.campaigns.length;
          rowPointer += campaignCount;
          weekContentRows = campaignCount;
          
          if (campaignCount > 0) {
            groupOperations.push({
              type: 'week_group',
              startRow: weekStartRow + 1,
              count: campaignCount
            });
          }
        }
      });

      const appContentRows = rowPointer - appStartRow - 1;
      if (appContentRows > 0) {
        groupOperations.push({
          type: 'app_group',
          startRow: appStartRow + 1,
          count: appContentRows
        });
      }
    });

    groupOperations.forEach(op => {
      try {
        sheet.getRange(op.startRow, 1, op.count, 1).shiftRowGroupDepth(1);
        sheet.getRange(op.startRow, 1, op.count, 1).collapseGroups();
      } catch (e) {}
    });
  } catch (e) {}
}

function createOverallPivotTable(appData) {
  createEnhancedPivotTable(appData);
}

function createProjectPivotTable(projectName, appData) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    createEnhancedPivotTable(appData);
  } finally {
    setCurrentProject(originalProject);
  }
}