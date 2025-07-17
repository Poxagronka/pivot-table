/**
 * Sheet Formatting and Table Creation - ОБНОВЛЕНО: унифицированные метрики + исправленная ширина столбцов
 */

function createEnhancedPivotTable(appData) {
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);
  else sheet.clear();

  const wow = calculateWoWMetrics(appData);
  const headers = getUnifiedHeaders();
  const tableData = [headers];
  const formatData = [];

  const appKeys = Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName));
  appKeys.forEach(appKey => {
    const app = appData[appKey];
    
    formatData.push({ row: tableData.length + 1, type: 'APP' });
    const emptyRow = new Array(headers.length).fill('');
    emptyRow[0] = 'APP';
    emptyRow[1] = app.appName;
    tableData.push(emptyRow);

    const weekKeys = Object.keys(app.weeks).sort();
    weekKeys.forEach(weekKey => {
      const week = app.weeks[weekKey];
      
      formatData.push({ row: tableData.length + 1, type: 'WEEK' });
      
      if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
        const allCampaigns = [];
        Object.values(week.sourceApps).forEach(sourceApp => {
          allCampaigns.push(...sourceApp.campaigns);
        });
        
        const weekTotals = calculateWeekTotals(allCampaigns);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = wow.appWeekWoW[appWeekKey] || {};
        
        const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
        const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
        const status = weekWoW.growthStatus || '';
        
        const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status);
        tableData.push(weekRow);
        
        addSourceAppRows(tableData, week.sourceApps, weekKey, wow, formatData);
        
      } else {
        const weekTotals = calculateWeekTotals(week.campaigns);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = wow.appWeekWoW[appWeekKey] || {};
        
        const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
        const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
        const status = weekWoW.growthStatus || '';
        
        const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status);
        tableData.push(weekRow);
        
        addCampaignRows(tableData, week.campaigns, week, weekKey, wow, formatData);
      }
    });
  });

  const range = sheet.getRange(1, 1, tableData.length, headers.length);
  range.setValues(tableData);
  
  applyEnhancedFormatting(sheet, tableData.length, headers.length, formatData, appData);
  createRowGrouping(sheet, tableData, appData);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2); // Заморозить первые 2 столбца (Level скрыт, Week Range / Source App видимый)
}

function createOverallPivotTable(appData) {
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);
  else sheet.clear();

  const wow = calculateWoWMetrics(appData);
  const headers = getUnifiedHeaders();
  const tableData = [headers];
  const formatData = [];

  const appKeys = Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName));
  appKeys.forEach(appKey => {
    const app = appData[appKey];
    
    formatData.push({ row: tableData.length + 1, type: 'APP' });
    const emptyRow = new Array(headers.length).fill('');
    emptyRow[0] = 'APP';
    emptyRow[1] = app.appName;
    tableData.push(emptyRow);

    const weekKeys = Object.keys(app.weeks).sort();
    weekKeys.forEach(weekKey => {
      const week = app.weeks[weekKey];
      
      const weekTotals = calculateWeekTotals(week.campaigns);
      const appWeekKey = `${app.appName}_${weekKey}`;
      const weekWoW = wow.appWeekWoW[appWeekKey] || {};
      
      const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = weekWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'WEEK' });
      const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status);
      tableData.push(weekRow);
    });
  });

  const range = sheet.getRange(1, 1, tableData.length, headers.length);
  range.setValues(tableData);
  
  applyEnhancedFormatting(sheet, tableData.length, headers.length, formatData, appData);
  createOverallRowGrouping(sheet, tableData, appData);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2); // Заморозить первые 2 столбца (Level скрыт, Week Range / Source App видимый)
}

function createOverallRowGrouping(sheet, tableData, appData) {
  const numCols = getUnifiedHeaders().length;

  try {
    let rowPointer = 2;
    const sortedApps = Object.keys(appData).sort((a, b) => 
      appData[a].appName.localeCompare(appData[b].appName)
    );

    sortedApps.forEach(appKey => {
      const app = appData[appKey];
      const appStartRow = rowPointer;
      rowPointer++;

      const sortedWeeks = Object.keys(app.weeks).sort();
      const weekCount = sortedWeeks.length;
      rowPointer += weekCount;

      if (weekCount > 0) {
        try {
          sheet.getRange(appStartRow + 1, 1, weekCount, numCols).shiftRowGroupDepth(1);
          sheet.getRange(appStartRow + 1, 1, weekCount, 1).collapseGroups();
        } catch (e) {
          console.log('Error grouping weeks under app:', e);
        }
      }
    });
    
    console.log('Overall row grouping completed successfully');
    
  } catch (e) {
    console.error('Error in createOverallRowGrouping:', e);
  }
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
    
    const sourceAppRow = createSourceAppRow(sourceAppDisplayName, sourceAppTotals, spendWoW, profitWoW, status);
    tableData.push(sourceAppRow);
    
    addCampaignRows(tableData, sourceApp.campaigns, { weekStart: weekKey.split('-').join('/'), weekEnd: '' }, weekKey, wow, formatData);
  });
}

function createSourceAppRow(sourceAppDisplayName, totals, spendWoW, profitWoW, status) {
  return [
    'SOURCE_APP', sourceAppDisplayName, '', '',
    totals.totalSpend.toFixed(2), spendWoW, totals.totalInstalls, totals.avgCpi.toFixed(3),
    totals.avgRoas.toFixed(2), totals.avgIpm.toFixed(1), `${totals.avgRrD1.toFixed(0)}%`, `${totals.avgRrD7.toFixed(0)}%`,
    totals.avgArpu.toFixed(3), `${totals.avgERoas.toFixed(0)}%`, `${totals.avgEROASD730.toFixed(0)}%`,
    totals.totalProfit.toFixed(2), profitWoW, status, ''
  ];
}

function getUnifiedHeaders() {
  return [
    'Level', 'Week Range / Source App', 'ID', 'GEO',
    'Spend', 'Spend WoW %', 'Installs', 'CPI', 'ROAS D-1', 'IPM',
    'RR D-1', 'RR D-7', 'eARPU 365d', 'eROAS 365d', 'eROAS 730d', 'eProfit 730d', 'eProfit 730d WoW %', 'Growth Status', 'Comments'
  ];
}

function createWeekRow(week, weekTotals, spendWoW, profitWoW, status) {
  return [
    'WEEK', `${week.weekStart} - ${week.weekEnd}`, '', '',
    weekTotals.totalSpend.toFixed(2), spendWoW, weekTotals.totalInstalls, weekTotals.avgCpi.toFixed(3),
    weekTotals.avgRoas.toFixed(2), weekTotals.avgIpm.toFixed(1), `${weekTotals.avgRrD1.toFixed(0)}%`, `${weekTotals.avgRrD7.toFixed(0)}%`,
    weekTotals.avgArpu.toFixed(3), `${weekTotals.avgERoas.toFixed(0)}%`, `${weekTotals.avgEROASD730.toFixed(0)}%`,
    weekTotals.totalProfit.toFixed(2), profitWoW, status, ''
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
    
    const commentsRange = sheet.getRange(2, numCols, numRows - 1, 1);
    commentsRange.setWrap(true).setHorizontalAlignment('left');
    
    const growthStatusRange = sheet.getRange(2, numCols - 1, numRows - 1, 1);
    growthStatusRange.setWrap(true).setHorizontalAlignment('left');
  }

  const appRows = [], weekRows = [], sourceAppRows = [], campaignRows = [], hyperlinkRows = [];
  formatData.forEach(item => {
    if (item.type === 'APP') appRows.push(item.row);
    if (item.type === 'WEEK') weekRows.push(item.row);
    if (item.type === 'SOURCE_APP') sourceAppRows.push(item.row);
    if (item.type === 'CAMPAIGN') campaignRows.push(item.row);
    if (item.type === 'HYPERLINK') hyperlinkRows.push(item.row);
  });

  appRows.forEach(r =>
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.APP_ROW.background)
         .setFontColor(COLORS.APP_ROW.fontColor)
         .setFontWeight('bold')
         .setFontSize(10)
  );

  weekRows.forEach(r =>
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.WEEK_ROW.background)
         .setFontSize(10)
  );

  sourceAppRows.forEach(r =>
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.SOURCE_APP_ROW.background)
         .setFontSize(9)
  );

  campaignRows.forEach(r =>
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.CAMPAIGN_ROW.background)
         .setFontSize(9)
  );

  if (hyperlinkRows.length > 0 && CURRENT_PROJECT === 'TRICKY') {
    hyperlinkRows.forEach(r => {
      const linkCell = sheet.getRange(r, 2);
      linkCell.setFontColor('#000000').setFontLine('none');
    });
  }

  if (numRows > 1) {
  sheet.getRange(2, 5, numRows - 1, 1).setNumberFormat('$0');        // Spend - до целого
  sheet.getRange(2, 8, numRows - 1, 1).setNumberFormat('$0.0');      // CPI - 1 знак после точки
  sheet.getRange(2, 9, numRows - 1, 1).setNumberFormat('0.0');       // ROAS D-1 - 1 знак после точки
  sheet.getRange(2, 10, numRows - 1, 1).setNumberFormat('0.0');      // IPM - без изменений
  sheet.getRange(2, 13, numRows - 1, 1).setNumberFormat('$0.0');     // eARPU 365d - 1 знак после точки
  sheet.getRange(2, 16, numRows - 1, 1).setNumberFormat('$0');       // eProfit 730d - до целого
}

  applyConditionalFormatting(sheet, numRows, appData);
  sheet.hideColumns(1);
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
      
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${String.fromCharCode(64 + eroasColumn)}${i + 1})), VALUE(SUBSTITUTE(${String.fromCharCode(64 + eroasColumn)}${i + 1},"%","")) >= ${targetEROAS})`)
          .setBackground(COLORS.POSITIVE.background)
          .setFontColor(COLORS.POSITIVE.fontColor)
          .setRanges([cellRange]).build()
      );
      
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${String.fromCharCode(64 + eroasColumn)}${i + 1})), VALUE(SUBSTITUTE(${String.fromCharCode(64 + eroasColumn)}${i + 1},"%","")) >= 120, VALUE(SUBSTITUTE(${String.fromCharCode(64 + eroasColumn)}${i + 1},"%","")) < ${targetEROAS})`)
          .setBackground(COLORS.WARNING.background)
          .setFontColor(COLORS.WARNING.fontColor)
          .setRanges([cellRange]).build()
      );
      
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${String.fromCharCode(64 + eroasColumn)}${i + 1})), VALUE(SUBSTITUTE(${String.fromCharCode(64 + eroasColumn)}${i + 1},"%","")) < 120)`)
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
  const avgRoas = campaigns.length ? campaigns.reduce((s, c) => s + c.roas, 0) / campaigns.length : 0;
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
    totalSpend, totalInstalls, avgCpi, avgRoas, avgIpm, avgRrD1, avgRrD7,
    avgArpu, avgERoas, avgEROASD730, totalProfit
  };
}

function addCampaignRows(tableData, campaigns, week, weekKey, wow, formatData) {
  if (CURRENT_PROJECT === 'OVERALL') {
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
    
    const campaignRow = createCampaignRow(campaign, campaignIdValue, spendPct, profitPct, growthStatus);
    tableData.push(campaignRow);
  });
}

function createCampaignRow(campaign, campaignIdValue, spendPct, profitPct, growthStatus) {
  return [
    'CAMPAIGN', campaign.sourceApp, campaignIdValue, campaign.geo,
    campaign.spend.toFixed(2), spendPct, campaign.installs, campaign.cpi ? campaign.cpi.toFixed(3) : '0.000',
    campaign.roas.toFixed(2), campaign.ipm.toFixed(1), `${campaign.rrD1.toFixed(0)}%`, `${campaign.rrD7.toFixed(0)}%`,
    campaign.eArpuForecast.toFixed(3), `${campaign.eRoasForecast.toFixed(0)}%`, `${campaign.eRoasForecastD730.toFixed(0)}%`,
    campaign.eProfitForecast.toFixed(2), profitPct, growthStatus, ''
  ];
}

function createRowGrouping(sheet, tableData, appData) {
  const numCols = getUnifiedHeaders().length;

  try {
    let rowPointer = 2;

    const sortedApps = Object.keys(appData).sort((a, b) => 
      appData[a].appName.localeCompare(appData[b].appName)
    );

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
              try {
                sheet.getRange(sourceAppStartRow + 1, 1, campaignCount, numCols).shiftRowGroupDepth(1);
                sheet.getRange(sourceAppStartRow + 1, 1, campaignCount, 1).collapseGroups();
              } catch (e) {
                console.log('Error grouping campaigns under source app:', e);
              }
            }
          });
          
          if (weekContentRows > 0) {
            try {
              sheet.getRange(weekStartRow + 1, 1, weekContentRows, numCols).shiftRowGroupDepth(1);
              sheet.getRange(weekStartRow + 1, 1, weekContentRows, 1).collapseGroups();
            } catch (e) {
              console.log('Error grouping week content:', e);
            }
          }
          
        } else if (CURRENT_PROJECT !== 'OVERALL') {
          const campaignCount = week.campaigns ? week.campaigns.length : 0;
          rowPointer += campaignCount;
          weekContentRows = campaignCount;
          
          if (campaignCount > 0) {
            try {
              sheet.getRange(weekStartRow + 1, 1, campaignCount, numCols).shiftRowGroupDepth(1);
              sheet.getRange(weekStartRow + 1, 1, campaignCount, 1).collapseGroups();
            } catch (e) {
              console.log('Error grouping campaigns under week:', e);
            }
          }
        }
      });

      const appContentRows = rowPointer - appStartRow - 1;
      if (appContentRows > 0) {
        try {
          sheet.getRange(appStartRow + 1, 1, appContentRows, numCols).shiftRowGroupDepth(1);
          sheet.getRange(appStartRow + 1, 1, appContentRows, 1).collapseGroups();
        } catch (e) {
          console.log('Error grouping app content:', e);
        }
      }
    });
    
    console.log('Row grouping completed successfully');
    
  } catch (e) {
    console.error('Error in createRowGrouping:', e);
  }
}

function createProjectPivotTable(projectName, appData) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    if (projectName === 'OVERALL') {
      createOverallPivotTable(appData);
    } else {
      createEnhancedPivotTable(appData);
    }
  } finally {
    setCurrentProject(originalProject);
  }
}