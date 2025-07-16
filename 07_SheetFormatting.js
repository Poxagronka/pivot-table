function createEnhancedPivotTable(appData) {
  if (CURRENT_PROJECT === 'TRICKY') {
    createTrickyOptimizedPivotTable(appData);
  } else if (CURRENT_PROJECT === 'OVERALL') {
    createOverallPivotTable(appData);
  } else {
    createStandardPivotTable(appData);
  }
}

function createTrickyOptimizedPivotTable(appData) {
  console.log('Creating TRICKY optimized pivot table...');
  const config = getCurrentConfig();
  const spreadsheetId = config.SHEET_ID;
  const sheetName = config.SHEET_NAME;
  
  const sheetId = getOrCreateSheetAdvanced(spreadsheetId, sheetName);
  const wow = calculateWoWMetrics(appData);
  const { tableData, formatData, hyperlinkData, groupData } = buildTrickyTableData(appData, wow);
  
  writeDataAdvanced(spreadsheetId, sheetId, tableData);
  applyBatchFormattingAdvanced(spreadsheetId, sheetId, formatData, hyperlinkData, tableData.length);
  applyTrickyGroupingAdvanced(spreadsheetId, sheetId, groupData);
  
  console.log('TRICKY optimized pivot table completed');
}

function createStandardPivotTable(appData) {
  console.log('Creating standard pivot table...');
  const config = getCurrentConfig();
  const spreadsheetId = config.SHEET_ID;
  const sheetName = config.SHEET_NAME;
  
  const sheetId = getOrCreateSheetAdvanced(spreadsheetId, sheetName);
  const wow = calculateWoWMetrics(appData);
  const { tableData, formatData, hyperlinkData, groupData } = buildStandardTableData(appData, wow);
  
  writeDataAdvanced(spreadsheetId, sheetId, tableData);
  applyBatchFormattingAdvanced(spreadsheetId, sheetId, formatData, hyperlinkData, tableData.length);
  applyStandardGroupingAdvanced(spreadsheetId, sheetId, groupData);
  
  console.log('Standard pivot table completed');
}

function createOverallPivotTable(appData) {
  console.log('Creating overall pivot table...');
  const config = getCurrentConfig();
  const spreadsheetId = config.SHEET_ID;
  const sheetName = config.SHEET_NAME;
  
  const sheetId = getOrCreateSheetAdvanced(spreadsheetId, sheetName);
  const wow = calculateWoWMetrics(appData);
  const { tableData, formatData, groupData } = buildOverallTableData(appData, wow);
  
  writeDataAdvanced(spreadsheetId, sheetId, tableData);
  applyBatchFormattingAdvanced(spreadsheetId, sheetId, formatData, null, tableData.length);
  applyOverallGroupingAdvanced(spreadsheetId, sheetId, groupData);
  
  console.log('Overall pivot table completed');
}

function getOrCreateSheetAdvanced(spreadsheetId, sheetName) {
  console.log(`Getting/creating sheet: ${sheetName}`);
  
  try {
    const spreadsheet = Sheets.Spreadsheets.get(spreadsheetId);
    const existingSheet = spreadsheet.sheets.find(sheet => sheet.properties.title === sheetName);
    
    if (existingSheet) {
      console.log(`Clearing existing sheet: ${sheetName}`);
      Sheets.Spreadsheets.batchUpdate({
        requests: [{
          updateCells: {
            range: { sheetId: existingSheet.properties.sheetId },
            fields: 'userEnteredValue,userEnteredFormat'
          }
        }]
      }, spreadsheetId);
      return existingSheet.properties.sheetId;
    }
    
    console.log(`Creating new sheet: ${sheetName}`);
    const response = Sheets.Spreadsheets.batchUpdate({
      requests: [{
        addSheet: {
          properties: {
            title: sheetName,
            gridProperties: { rowCount: 1000, columnCount: 20 }
          }
        }
      }]
    }, spreadsheetId);
    
    return response.replies[0].addSheet.properties.sheetId;
  } catch (e) {
    console.error('Error in getOrCreateSheetAdvanced:', e);
    throw e;
  }
}

function writeDataAdvanced(spreadsheetId, sheetId, tableData) {
  console.log(`Writing ${tableData.length} rows to sheet...`);
  
  const requests = [{
    updateCells: {
      range: {
        sheetId: sheetId,
        startRowIndex: 0,
        endRowIndex: tableData.length,
        startColumnIndex: 0,
        endColumnIndex: tableData[0].length
      },
      rows: tableData.map(row => ({
        values: row.map(cell => ({
          userEnteredValue: { stringValue: cell.toString() }
        }))
      })),
      fields: 'userEnteredValue'
    }
  }];
  
  Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
  console.log('Data written successfully');
}

function applyBatchFormattingAdvanced(spreadsheetId, sheetId, formatData, hyperlinkData, numRows) {
  console.log('Applying batch formatting...');
  const requests = [];
  const headers = getUnifiedHeaders();
  const numCols = headers.length;
  
  requests.push(...getHeaderFormatRequestsAdvanced(sheetId, numCols));
  requests.push(...getColumnFormatRequestsAdvanced(sheetId, numRows, numCols));
  requests.push(...getRowFormatRequestsAdvanced(sheetId, formatData, numCols));
  
  if (hyperlinkData && hyperlinkData.length > 0) {
    requests.push(...getHyperlinkFormatRequestsAdvanced(sheetId, hyperlinkData));
  }
  
  requests.push({
    updateDimensionProperties: {
      range: { sheetId: sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
      properties: { hiddenByUser: true },
      fields: 'hiddenByUser'
    }
  });
  
  requests.push({
    updateSheetProperties: {
      properties: { sheetId: sheetId, gridProperties: { frozenRowCount: 1 } },
      fields: 'gridProperties.frozenRowCount'
    }
  });
  
  if (requests.length > 0) {
    console.log(`Executing ${requests.length} format requests...`);
    const batchSize = 100;
    for (let i = 0; i < requests.length; i += batchSize) {
      const batch = requests.slice(i, i + batchSize);
      Sheets.Spreadsheets.batchUpdate({ requests: batch }, spreadsheetId);
      
      if (i + batchSize < requests.length) {
        Utilities.sleep(1000);
      }
    }
  }
  
  console.log('Formatting applied successfully');
}

function getHeaderFormatRequestsAdvanced(sheetId, numCols) {
  return [{
    repeatCell: {
      range: { sheetId: sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: numCols },
      cell: {
        userEnteredFormat: {
          backgroundColor: { red: 0.258, green: 0.522, blue: 0.957 },
          textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
          verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP'
        }
      },
      fields: 'userEnteredFormat'
    }
  }];
}

function getColumnFormatRequestsAdvanced(sheetId, numRows, numCols) {
  const requests = [];
  const columnWidths = TABLE_CONFIG.COLUMN_WIDTHS;
  
  columnWidths.forEach(col => {
    requests.push({
      updateDimensionProperties: {
        range: { sheetId: sheetId, dimension: 'COLUMNS', startIndex: col.c - 1, endIndex: col.c },
        properties: { pixelSize: col.w },
        fields: 'pixelSize'
      }
    });
  });
  
  if (numRows > 1) {
    requests.push({
      repeatCell: {
        range: { sheetId: sheetId, startRowIndex: 1, endRowIndex: numRows, startColumnIndex: 4, endColumnIndex: 5 },
        cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$0.00' } } },
        fields: 'userEnteredFormat.numberFormat'
      }
    });
    
    requests.push({
      repeatCell: {
        range: { sheetId: sheetId, startRowIndex: 1, endRowIndex: numRows, startColumnIndex: 7, endColumnIndex: 8 },
        cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$0.000' } } },
        fields: 'userEnteredFormat.numberFormat'
      }
    });
    
    requests.push({
      repeatCell: {
        range: { sheetId: sheetId, startRowIndex: 1, endRowIndex: numRows, startColumnIndex: 15, endColumnIndex: 16 },
        cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$0.00' } } },
        fields: 'userEnteredFormat.numberFormat'
      }
    });
  }
  
  return requests;
}

function getRowFormatRequestsAdvanced(sheetId, formatData, numCols) {
  const requests = [];
  const rowsByType = { APP: [], WEEK: [], SOURCE_APP: [], CAMPAIGN: [] };
  
  formatData.forEach(item => {
    if (rowsByType[item.type]) {
      rowsByType[item.type].push(item.row - 1);
    }
  });
  
  const formatConfigs = {
    APP: { bg: COLORS.APP_ROW.background, fg: COLORS.APP_ROW.fontColor, bold: true, size: 10 },
    WEEK: { bg: COLORS.WEEK_ROW.background, fg: null, bold: false, size: 10 },
    SOURCE_APP: { bg: COLORS.SOURCE_APP_ROW.background, fg: null, bold: false, size: 9 },
    CAMPAIGN: { bg: COLORS.CAMPAIGN_ROW.background, fg: null, bold: false, size: 9 }
  };
  
  Object.entries(rowsByType).forEach(([type, rows]) => {
    const config = formatConfigs[type];
    
    if (rows.length > 0) {
      const ranges = rows.map(rowIndex => ({
        sheetId: sheetId,
        startRowIndex: rowIndex,
        endRowIndex: rowIndex + 1,
        startColumnIndex: 0,
        endColumnIndex: numCols
      }));
      
      requests.push({
        repeatCell: {
          ranges: ranges,
          cell: {
            userEnteredFormat: {
              backgroundColor: hexToRgb(config.bg),
              textFormat: {
                foregroundColor: config.fg ? hexToRgb(config.fg) : { red: 0, green: 0, blue: 0 },
                bold: config.bold,
                fontSize: config.size
              },
              verticalAlignment: 'MIDDLE'
            }
          },
          fields: 'userEnteredFormat'
        }
      });
    }
  });
  
  return requests;
}

function getHyperlinkFormatRequestsAdvanced(sheetId, hyperlinkData) {
  return hyperlinkData.map(link => ({
    repeatCell: {
      range: { 
        sheetId: sheetId,
        startRowIndex: link.row - 1,
        endRowIndex: link.row,
        startColumnIndex: link.col - 1,
        endColumnIndex: link.col
      },
      cell: {
        userEnteredFormat: {
          textFormat: { foregroundColor: { red: 0, green: 0, blue: 0 }, underline: false }
        }
      },
      fields: 'userEnteredFormat.textFormat'
    }
  }));
}

function applyTrickyGroupingAdvanced(spreadsheetId, sheetId, groupData) {
  console.log('Applying TRICKY grouping...');
  const requests = [];
  
  groupData.sourceApps.forEach(group => {
    requests.push({
      addDimensionGroup: {
        range: { sheetId: sheetId, dimension: 'ROWS', startIndex: group.startRow, endIndex: group.startRow + group.count }
      }
    });
  });
  
  groupData.weeks.forEach(group => {
    requests.push({
      addDimensionGroup: {
        range: { sheetId: sheetId, dimension: 'ROWS', startIndex: group.startRow, endIndex: group.startRow + group.count }
      }
    });
  });
  
  groupData.apps.forEach(group => {
    requests.push({
      addDimensionGroup: {
        range: { sheetId: sheetId, dimension: 'ROWS', startIndex: group.startRow, endIndex: group.startRow + group.count }
      }
    });
  });
  
  if (requests.length > 0) {
    console.log(`Creating ${requests.length} groups...`);
    const batchSize = 50;
    for (let i = 0; i < requests.length; i += batchSize) {
      const batch = requests.slice(i, i + batchSize);
      Sheets.Spreadsheets.batchUpdate({ requests: batch }, spreadsheetId);
      
      if (i + batchSize < requests.length) {
        Utilities.sleep(500);
      }
    }
  }
  
  console.log('TRICKY grouping applied successfully');
}

function applyStandardGroupingAdvanced(spreadsheetId, sheetId, groupData) {
  console.log('Applying standard grouping...');
  const requests = [];
  
  groupData.weeks.forEach(group => {
    requests.push({
      addDimensionGroup: {
        range: { sheetId: sheetId, dimension: 'ROWS', startIndex: group.startRow, endIndex: group.startRow + group.count }
      }
    });
  });
  
  groupData.apps.forEach(group => {
    requests.push({
      addDimensionGroup: {
        range: { sheetId: sheetId, dimension: 'ROWS', startIndex: group.startRow, endIndex: group.startRow + group.count }
      }
    });
  });
  
  if (requests.length > 0) {
    console.log(`Creating ${requests.length} groups...`);
    const batchSize = 50;
    for (let i = 0; i < requests.length; i += batchSize) {
      const batch = requests.slice(i, i + batchSize);
      Sheets.Spreadsheets.batchUpdate({ requests: batch }, spreadsheetId);
      
      if (i + batchSize < requests.length) {
        Utilities.sleep(500);
      }
    }
  }
  
  console.log('Standard grouping applied successfully');
}

function applyOverallGroupingAdvanced(spreadsheetId, sheetId, groupData) {
  console.log('Applying overall grouping...');
  const requests = [];
  
  groupData.apps.forEach(group => {
    requests.push({
      addDimensionGroup: {
        range: { sheetId: sheetId, dimension: 'ROWS', startIndex: group.startRow, endIndex: group.startRow + group.count }
      }
    });
  });
  
  if (requests.length > 0) {
    console.log(`Creating ${requests.length} groups...`);
    const batchSize = 50;
    for (let i = 0; i < requests.length; i += batchSize) {
      const batch = requests.slice(i, i + batchSize);
      Sheets.Spreadsheets.batchUpdate({ requests: batch }, spreadsheetId);
      
      if (i + batchSize < requests.length) {
        Utilities.sleep(500);
      }
    }
  }
  
  console.log('Overall grouping applied successfully');
}

function buildTrickyTableData(appData, wow) {
  const headers = getUnifiedHeaders();
  const tableData = [headers];
  const formatData = [];
  const hyperlinkData = [];
  const groupData = { apps: [], weeks: [], sourceApps: [] };
  
  const cache = initTrickyOptimizedCache();
  
  Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName)).forEach(appKey => {
    const app = appData[appKey];
    const appStartRow = tableData.length;
    
    const emptyRow = new Array(headers.length).fill('');
    emptyRow[0] = 'APP';
    emptyRow[1] = app.appName;
    tableData.push(emptyRow);
    formatData.push({ row: tableData.length, type: 'APP' });
    
    let appContentRows = 0;
    
    Object.keys(app.weeks).sort().forEach(weekKey => {
      const week = app.weeks[weekKey];
      const weekStartRow = tableData.length;
      
      const allCampaigns = [];
      Object.values(week.sourceApps || {}).forEach(sourceApp => {
        allCampaigns.push(...sourceApp.campaigns);
      });
      
      const weekTotals = calculateWeekTotals(allCampaigns);
      const appWeekKey = `${app.appName}_${weekKey}`;
      const weekWoW = wow.appWeekWoW[appWeekKey] || {};
      
      const weekRow = createWeekRow(week, weekTotals, weekWoW);
      tableData.push(weekRow);
      formatData.push({ row: tableData.length, type: 'WEEK' });
      
      let weekContentRows = 0;
      
      if (week.sourceApps) {
        Object.keys(week.sourceApps).sort((a, b) => {
          const spendA = week.sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
          const spendB = week.sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
          return spendB - spendA;
        }).forEach(sourceAppKey => {
          const sourceApp = week.sourceApps[sourceAppKey];
          const sourceAppStartRow = tableData.length;
          const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
          
          const sourceAppWoWKey = `${sourceApp.sourceAppId}_${weekKey}`;
          const sourceAppWoW = wow.sourceAppWoW[sourceAppWoWKey] || {};
          
          let sourceAppDisplayName = sourceApp.sourceAppName;
          const appInfo = cache?.appsDbCache[sourceApp.sourceAppId];
          if (appInfo && appInfo.linkApp) {
            sourceAppDisplayName = `=HYPERLINK("${appInfo.linkApp}", "${sourceApp.sourceAppName}")`;
            hyperlinkData.push({ row: tableData.length + 1, col: 2 });
          }
          
          const sourceAppRow = createSourceAppRow(sourceAppDisplayName, sourceAppTotals, sourceAppWoW);
          tableData.push(sourceAppRow);
          formatData.push({ row: tableData.length, type: 'SOURCE_APP' });
          
          sourceApp.campaigns.sort((a, b) => b.spend - a.spend).forEach(campaign => {
            const campaignIdValue = `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`;
            const key = `${campaign.campaignId}_${weekKey}`;
            const campaignWoW = wow.campaignWoW[key] || {};
            
            const campaignRow = createCampaignRow(campaign, campaignIdValue, campaignWoW);
            tableData.push(campaignRow);
            formatData.push({ row: tableData.length, type: 'CAMPAIGN' });
          });
          
          const campaignCount = sourceApp.campaigns.length;
          weekContentRows += 1 + campaignCount;
          
          if (campaignCount > 0) {
            groupData.sourceApps.push({
              startRow: sourceAppStartRow + 1,
              count: campaignCount
            });
          }
        });
        
        if (weekContentRows > 0) {
          groupData.weeks.push({
            startRow: weekStartRow + 1,
            count: weekContentRows
          });
        }
      }
      
      appContentRows += 1 + weekContentRows;
    });
    
    if (appContentRows > 0) {
      groupData.apps.push({
        startRow: appStartRow + 1,
        count: appContentRows
      });
    }
  });
  
  return { tableData, formatData, hyperlinkData, groupData };
}

function buildStandardTableData(appData, wow) {
  const headers = getUnifiedHeaders();
  const tableData = [headers];
  const formatData = [];
  const hyperlinkData = [];
  const groupData = { apps: [], weeks: [] };
  
  Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName)).forEach(appKey => {
    const app = appData[appKey];
    const appStartRow = tableData.length;
    
    const emptyRow = new Array(headers.length).fill('');
    emptyRow[0] = 'APP';
    emptyRow[1] = app.appName;
    tableData.push(emptyRow);
    formatData.push({ row: tableData.length, type: 'APP' });
    
    let appContentRows = 0;
    
    Object.keys(app.weeks).sort().forEach(weekKey => {
      const week = app.weeks[weekKey];
      const weekStartRow = tableData.length;
      
      const weekTotals = calculateWeekTotals(week.campaigns);
      const appWeekKey = `${app.appName}_${weekKey}`;
      const weekWoW = wow.appWeekWoW[appWeekKey] || {};
      
      const weekRow = createWeekRow(week, weekTotals, weekWoW);
      tableData.push(weekRow);
      formatData.push({ row: tableData.length, type: 'WEEK' });
      
      const campaignCount = week.campaigns ? week.campaigns.length : 0;
      
      if (week.campaigns) {
        week.campaigns.sort((a, b) => b.spend - a.spend).forEach(campaign => {
          const campaignIdValue = (CURRENT_PROJECT === 'TRICKY' || CURRENT_PROJECT === 'REGULAR') ?
            `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")` :
            campaign.campaignId;
          
          const key = `${campaign.campaignId}_${weekKey}`;
          const campaignWoW = wow.campaignWoW[key] || {};
          
          const campaignRow = createCampaignRow(campaign, campaignIdValue, campaignWoW);
          tableData.push(campaignRow);
          formatData.push({ row: tableData.length, type: 'CAMPAIGN' });
        });
      }
      
      appContentRows += 1 + campaignCount;
      
      if (campaignCount > 0) {
        groupData.weeks.push({
          startRow: weekStartRow + 1,
          count: campaignCount
        });
      }
    });
    
    if (appContentRows > 0) {
      groupData.apps.push({
        startRow: appStartRow + 1,
        count: appContentRows
      });
    }
  });
  
  return { tableData, formatData, hyperlinkData, groupData };
}

function buildOverallTableData(appData, wow) {
  const headers = getUnifiedHeaders();
  const tableData = [headers];
  const formatData = [];
  const groupData = { apps: [], weeks: [] };
  
  Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName)).forEach(appKey => {
    const app = appData[appKey];
    const appStartRow = tableData.length;
    
    const emptyRow = new Array(headers.length).fill('');
    emptyRow[0] = 'APP';
    emptyRow[1] = app.appName;
    tableData.push(emptyRow);
    formatData.push({ row: tableData.length, type: 'APP' });
    
    const weekCount = Object.keys(app.weeks).length;
    
    Object.keys(app.weeks).sort().forEach(weekKey => {
      const week = app.weeks[weekKey];
      const weekTotals = calculateWeekTotals(week.campaigns);
      const appWeekKey = `${app.appName}_${weekKey}`;
      const weekWoW = wow.appWeekWoW[appWeekKey] || {};
      
      const weekRow = createWeekRow(week, weekTotals, weekWoW);
      tableData.push(weekRow);
      formatData.push({ row: tableData.length, type: 'WEEK' });
    });
    
    if (weekCount > 0) {
      groupData.apps.push({
        startRow: appStartRow + 1,
        count: weekCount
      });
    }
  });
  
  return { tableData, formatData, groupData };
}

function createWeekRow(week, weekTotals, weekWoW) {
  const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
  const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
  const status = weekWoW.growthStatus || '';
  
  return [
    'WEEK', `${week.weekStart} - ${week.weekEnd}`, '', '',
    weekTotals.totalSpend.toFixed(2), spendWoW, weekTotals.totalInstalls, weekTotals.avgCpi.toFixed(3),
    weekTotals.avgRoas.toFixed(2), weekTotals.avgIpm.toFixed(1), `${weekTotals.avgRrD1.toFixed(1)}%`, `${weekTotals.avgRrD7.toFixed(1)}%`,
    weekTotals.avgArpu.toFixed(3), `${weekTotals.avgERoas.toFixed(0)}%`, `${weekTotals.avgEROASD730.toFixed(0)}%`,
    weekTotals.totalProfit.toFixed(2), profitWoW, status, ''
  ];
}

function createSourceAppRow(sourceAppDisplayName, totals, sourceAppWoW) {
  const spendWoW = sourceAppWoW.spendChangePercent !== undefined ? `${sourceAppWoW.spendChangePercent.toFixed(0)}%` : '';
  const profitWoW = sourceAppWoW.eProfitChangePercent !== undefined ? `${sourceAppWoW.eProfitChangePercent.toFixed(0)}%` : '';
  const status = sourceAppWoW.growthStatus || '';
  
  return [
    'SOURCE_APP', sourceAppDisplayName, '', '',
    totals.totalSpend.toFixed(2), spendWoW, totals.totalInstalls, totals.avgCpi.toFixed(3),
    totals.avgRoas.toFixed(2), totals.avgIpm.toFixed(1), `${totals.avgRrD1.toFixed(1)}%`, `${totals.avgRrD7.toFixed(1)}%`,
    totals.avgArpu.toFixed(3), `${totals.avgERoas.toFixed(0)}%`, `${totals.avgEROASD730.toFixed(0)}%`,
    totals.totalProfit.toFixed(2), profitWoW, status, ''
  ];
}

function createCampaignRow(campaign, campaignIdValue, campaignWoW) {
  const spendPct = campaignWoW.spendChangePercent !== undefined ? `${campaignWoW.spendChangePercent.toFixed(0)}%` : '';
  const profitPct = campaignWoW.eProfitChangePercent !== undefined ? `${campaignWoW.eProfitChangePercent.toFixed(0)}%` : '';
  const growthStatus = campaignWoW.growthStatus || '';
  
  return [
    'CAMPAIGN', campaign.sourceApp, campaignIdValue, campaign.geo,
    campaign.spend.toFixed(2), spendPct, campaign.installs, campaign.cpi ? campaign.cpi.toFixed(3) : '0.000',
    campaign.roas.toFixed(2), campaign.ipm.toFixed(1), `${campaign.rrD1.toFixed(1)}%`, `${campaign.rrD7.toFixed(1)}%`,
    campaign.eArpuForecast.toFixed(3), `${campaign.eRoasForecast.toFixed(0)}%`, `${campaign.eRoasForecastD730.toFixed(0)}%`,
    campaign.eProfitForecast.toFixed(2), profitPct, growthStatus, ''
  ];
}

function getUnifiedHeaders() {
  return [
    'Level', 'Week Range / Source App', 'ID', 'GEO',
    'Spend', 'Spend WoW %', 'Installs', 'CPI', 'ROAS D-1', 'IPM',
    'RR D-1', 'RR D-7', 'eARPU 365d', 'eROAS 365d', 'eROAS 730d', 'eProfit 730d', 'eProfit 730d WoW %', 'Growth Status', 'Comments'
  ];
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

function hexToRgb(hex) {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    red: parseInt(result[1], 16) / 255,
    green: parseInt(result[2], 16) / 255,
    blue: parseInt(result[3], 16) / 255
  } : { red: 0, green: 0, blue: 0 };
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