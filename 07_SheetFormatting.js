/**
 * Sheet Formatting and Table Creation - Multi Project Support
 * ОБНОВЛЕНО: Добавлена поддержка Applovin (использует формат Google Ads)
 */

/**
 * Create enhanced pivot table with week-level WoW analysis
 */
function createEnhancedPivotTable(appData) {
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);
  else sheet.clear();

  // ВАЖНО: Сначала вычисляем WoW метрики
  const wow = calculateWoWMetrics(appData);

  const headers = getProjectHeaders();
  const tableData = [headers];
  const formatData = [];

  // Process each app
  const appKeys = Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName));
  appKeys.forEach(appKey => {
    const app = appData[appKey];
    
    // APP row
    formatData.push({ row: tableData.length + 1, type: 'APP' });
    const emptyRow = new Array(headers.length).fill('');
    emptyRow[0] = 'APP';
    emptyRow[1] = app.appName;
    tableData.push(emptyRow);

    // Process weeks for this app
    const weekKeys = Object.keys(app.weeks).sort();
    weekKeys.forEach(weekKey => {
      const week = app.weeks[weekKey];
      const campaigns = week.campaigns;
      
      // Calculate week totals
      const weekTotals = calculateWeekTotals(campaigns);
      
      // ИСПРАВЛЕНО: Получаем WoW данные для недели
      const appWeekKey = `${app.appName}_${weekKey}`;
      const weekWoW = wow.appWeekWoW[appWeekKey] || {};
      
      const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = weekWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'WEEK' });
      
      const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status);
      tableData.push(weekRow);
      
      // Add campaign rows
      addCampaignRows(tableData, campaigns, week, weekKey, wow, formatData);
    });
  });

  // Write data to sheet
  const range = sheet.getRange(1, 1, tableData.length, headers.length);
  range.setValues(tableData);
  
  // Apply formatting and grouping
  applyEnhancedFormatting(sheet, tableData.length, headers.length, formatData);
  createRowGrouping(sheet, tableData, appData);
  sheet.setFrozenRows(1);
}

/**
 * Get project-specific headers
 */
function getProjectHeaders() {
  if (CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN') {
    return [
      'Level', 'Week Range / Source App', 'ID', 'GEO',
      'Spend', 'Spend WoW %', 'Installs', 'CPI', 'ROAS D-1', 'RR D-1',
      'RR D-7', 'eROAS 365d', 'eProfit 730d', 'eProfit 730d WoW %', 'Growth Status', 'Comments'
    ];
  } else {
    return [
      'Level', 'Week Range / Source App', 'ID', 'GEO',
      'Spend', 'Spend WoW %', 'Installs', 'CPI', 'ROAS D-1', 'IPM',
      'eARPU 365d', 'eROAS 365d', 'eProfit 730d', 'eProfit 730d WoW %', 'Growth Status', 'Comments'
    ];
  }
}

/**
 * Create week row based on project type
 */
function createWeekRow(week, weekTotals, spendWoW, profitWoW, status) {
  if (CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN') {
    return [
      'WEEK',
      `${week.weekStart} - ${week.weekEnd}`,
      '', '',
      weekTotals.totalSpend.toFixed(2),
      spendWoW,
      weekTotals.totalInstalls,
      weekTotals.avgCpi.toFixed(3),
      weekTotals.avgRoas.toFixed(2),
      `${weekTotals.avgRrD1.toFixed(1)}%`,
      `${weekTotals.avgRrD7.toFixed(1)}%`,
      `${weekTotals.avgERoas.toFixed(0)}%`,
      weekTotals.totalProfit.toFixed(2),
      profitWoW,
      status,
      ''
    ];
  } else {
    return [
      'WEEK',
      `${week.weekStart} - ${week.weekEnd}`,
      '', '',
      weekTotals.totalSpend.toFixed(2),
      spendWoW,
      weekTotals.totalInstalls,
      weekTotals.avgCpi.toFixed(3),
      weekTotals.avgRoas.toFixed(2),
      weekTotals.avgIpm.toFixed(1),
      weekTotals.avgArpu.toFixed(3),
      `${weekTotals.avgERoas.toFixed(0)}%`,
      weekTotals.totalProfit.toFixed(2),
      profitWoW,
      status,
      ''
    ];
  }
}

/**
 * Enhanced formatting function with text wrap for headers and comments
 */
function applyEnhancedFormatting(sheet, numRows, numCols, formatData) {
  const config = getCurrentConfig();
  
  // Header formatting with text wrap
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground(COLORS.HEADER.background)
    .setFontColor(COLORS.HEADER.fontColor)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setWrap(true);

  // Set column widths based on project
  const columnWidths = getProjectColumnWidths();
  columnWidths.forEach(col => {
    sheet.setColumnWidth(col.c, col.w);
  });

  // Text wrap для колонки комментариев И Growth Status
  if (numRows > 1) {
    // Comments column (last column)
    const commentsRange = sheet.getRange(2, numCols, numRows - 1, 1);
    commentsRange
      .setWrap(true)
      .setVerticalAlignment('top')
      .setHorizontalAlignment('left');
    
    // Growth Status column (second to last)
    const growthStatusRange = sheet.getRange(2, numCols - 1, numRows - 1, 1);
    growthStatusRange
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('left');
  }

  // Categorize rows by type
  const appRows = [], weekRows = [];
  formatData.forEach(item => {
    if (item.type === 'APP') appRows.push(item.row);
    if (item.type === 'WEEK') weekRows.push(item.row);
  });
  
  const campaignRows = [];
  for (let r = 2; r <= numRows; r++) {
    if (!appRows.includes(r) && !weekRows.includes(r)) campaignRows.push(r);
  }

  // Apply row formatting
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

  campaignRows.forEach(r =>
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.CAMPAIGN_ROW.background)
         .setFontSize(9)
  );

  // Apply numeric formats
  if (numRows > 1) {
    sheet.getRange(2, 5, numRows - 1, 1).setNumberFormat('$0.00'); // Spend
    sheet.getRange(2, 8, numRows - 1, 1).setNumberFormat('$0.000'); // CPI
    sheet.getRange(2, 9, numRows - 1, 1).setNumberFormat('0.00'); // ROAS
    
    if (CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN') {
      // RR columns already formatted as percentages in data
      sheet.getRange(2, 13, numRows - 1, 1).setNumberFormat('$0.00'); // E-Profit
    } else {
      sheet.getRange(2, 10, numRows - 1, 1).setNumberFormat('0.0'); // IPM
      sheet.getRange(2, 11, numRows - 1, 1).setNumberFormat('$0.000'); // E-ARPU
      sheet.getRange(2, 13, numRows - 1, 1).setNumberFormat('$0.00'); // E-Profit
    }
  }

  // Apply conditional formatting
  applyConditionalFormatting(sheet, numRows);
  
  // Hide Level column
  sheet.hideColumns(1);
}

/**
 * Get project-specific column widths
 */
function getProjectColumnWidths() {
  if (CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN') {
    return [
      { c: 1, w: 80 },   // Level
      { c: 2, w: 300 },  // Week Range / Source App
      { c: 3, w: 50 },   // ID
      { c: 4, w: 50 },   // GEO
      { c: 5, w: 75 },   // Spend
      { c: 6, w: 80 },   // Spend WoW %
      { c: 7, w: 60 },   // Installs
      { c: 8, w: 60 },   // CPI
      { c: 9, w: 60 },   // ROAS D-1
      { c: 10, w: 50 },  // RR D-1
      { c: 11, w: 50 },  // RR D-7
      { c: 12, w: 75 },  // eROAS 365d
      { c: 13, w: 75 },  // eProfit 730d
      { c: 14, w: 85 },  // eProfit 730d WoW %
      { c: 15, w: 160 }, // Growth Status
      { c: 16, w: 250 }  // Comments
    ];
  } else {
    return TABLE_CONFIG.COLUMN_WIDTHS;
  }
}

/**
 * Comprehensive conditional formatting with all growth status variations
 */
function applyConditionalFormatting(sheet, numRows) {
  const config = getCurrentConfig();
  const rules = [];
  
  if (numRows > 1) {
    // Spend WoW % formatting (колонка F = 6)
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

    // E-ROAS 365d formatting (колонка 12 для всех проектов)
    const eroasColumn = 12;
    const eroasRange = sheet.getRange(2, eroasColumn, numRows - 1, 1);
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${String.fromCharCode(64 + eroasColumn)}2)), VALUE(SUBSTITUTE(${String.fromCharCode(64 + eroasColumn)}2,"%","")) >= ${config.TARGET_EROAS})`)
        .setBackground(COLORS.POSITIVE.background)
        .setFontColor(COLORS.POSITIVE.fontColor)
        .setRanges([eroasRange]).build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${String.fromCharCode(64 + eroasColumn)}2)), VALUE(SUBSTITUTE(${String.fromCharCode(64 + eroasColumn)}2,"%","")) >= 120, VALUE(SUBSTITUTE(${String.fromCharCode(64 + eroasColumn)}2,"%","")) < ${config.TARGET_EROAS})`)
        .setBackground(COLORS.WARNING.background)
        .setFontColor(COLORS.WARNING.fontColor)
        .setRanges([eroasRange]).build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${String.fromCharCode(64 + eroasColumn)}2)), VALUE(SUBSTITUTE(${String.fromCharCode(64 + eroasColumn)}2,"%","")) < 120)`)
        .setBackground(COLORS.NEGATIVE.background)
        .setFontColor(COLORS.NEGATIVE.fontColor)
        .setRanges([eroasRange]).build()
    );

    // eProfit WoW % formatting (колонка 14 для всех проектов)
    const profitColumn = 14;
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

    // Growth Status formatting (колонка 15 для всех проектов)
    const growthColumn = 15;
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

/**
 * Calculate totals for a week's campaigns - Updated for Google Ads/Applovin metrics
 */
function calculateWeekTotals(campaigns) {
  const totalSpend = campaigns.reduce((s, c) => s + c.spend, 0);
  const totalInstalls = campaigns.reduce((s, c) => s + c.installs, 0);
  const avgCpi = totalInstalls ? totalSpend / totalInstalls : 0;
  const avgRoas = campaigns.length ? campaigns.reduce((s, c) => s + c.roas, 0) / campaigns.length : 0;
  
  // Calculate spend-weighted average eROAS
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
  
  const totalProfit = campaigns.reduce((s, c) => s + c.eProfitForecast, 0);

  // Project-specific metrics
  if (CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN') {
    const avgRrD1 = campaigns.length ? campaigns.reduce((s, c) => s + c.rrD1, 0) / campaigns.length : 0;
    const avgRrD7 = campaigns.length ? campaigns.reduce((s, c) => s + c.rrD7, 0) / campaigns.length : 0;
    
    return {
      totalSpend, totalInstalls, avgCpi, avgRoas, avgERoas, totalProfit,
      avgRrD1, avgRrD7, avgIpm: 0, avgArpu: 0
    };
  } else {
    const avgIpm = campaigns.length ? campaigns.reduce((s, c) => s + c.ipm, 0) / campaigns.length : 0;
    const avgArpu = campaigns.length ? campaigns.reduce((s, c) => s + c.eArpuForecast, 0) / campaigns.length : 0;
    
    return {
      totalSpend, totalInstalls, avgCpi, avgRoas, avgERoas, totalProfit,
      avgIpm, avgArpu, avgRrD1: 0, avgRrD7: 0
    };
  }
}

/**
 * Add campaign rows to table data
 */
function addCampaignRows(tableData, campaigns, week, weekKey, wow, formatData) {
  // Sort campaigns by spend (highest first)
  campaigns.sort((a, b) => b.spend - a.spend).forEach(campaign => {
    const link = `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`;
    
    // Получаем WoW данные для кампании по campaignId
    const key = `${campaign.campaignId}_${weekKey}`;
    const campaignWoW = wow.campaignWoW[key] || {};
    
    const spendPct = campaignWoW.spendChangePercent !== undefined ? `${campaignWoW.spendChangePercent.toFixed(0)}%` : '';
    const profitPct = campaignWoW.eProfitChangePercent !== undefined ? `${campaignWoW.eProfitChangePercent.toFixed(0)}%` : '';
    const growthStatus = campaignWoW.growthStatus || '';
    
    formatData.push({ row: tableData.length + 1, type: 'CAMPAIGN' });
    
    const campaignRow = createCampaignRow(campaign, link, spendPct, profitPct, growthStatus);
    tableData.push(campaignRow);
  });
}

/**
 * Create campaign row based on project type
 */
function createCampaignRow(campaign, link, spendPct, profitPct, growthStatus) {
  if (CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN') {
    return [
      'CAMPAIGN',
      campaign.sourceApp,
      link,
      campaign.geo,
      campaign.spend.toFixed(2),
      spendPct,
      campaign.installs,
      campaign.cpi ? campaign.cpi.toFixed(3) : '0.000',
      campaign.roas.toFixed(2),
      `${campaign.rrD1.toFixed(1)}%`,
      `${campaign.rrD7.toFixed(1)}%`,
      `${campaign.eRoasForecast.toFixed(0)}%`,
      campaign.eProfitForecast.toFixed(2),
      profitPct,
      growthStatus,
      ''
    ];
  } else {
    return [
      'CAMPAIGN',
      campaign.sourceApp,
      link,
      campaign.geo,
      campaign.spend.toFixed(2),
      spendPct,
      campaign.installs,
      campaign.cpi ? campaign.cpi.toFixed(3) : '0.000',
      campaign.roas.toFixed(2),
      campaign.ipm.toFixed(1),
      campaign.eArpuForecast.toFixed(3),
      `${campaign.eRoasForecast.toFixed(0)}%`,
      campaign.eProfitForecast.toFixed(2),
      profitPct,
      growthStatus,
      ''
    ];
  }
}

/**
 * Create row grouping: App → Week → Campaign
 */
function createRowGrouping(sheet, tableData, appData) {
  const numCols = getProjectHeaders().length;

  // If appData not provided, parse structure from tableData
  if (!appData) {
    parseAndCreateGrouping(sheet, tableData, numCols);
    return;
  }

  // Original logic when appData is provided
  createGroupingFromAppData(sheet, appData, numCols);
}

/**
 * Parse table data and create grouping
 */
function parseAndCreateGrouping(sheet, tableData, numCols) {
  let currentApp = null;
  let appStartRow = null;
  let weekStartRow = null;
  const appGroups = [];
  const weekGroups = [];
  
  for (let i = 1; i < tableData.length; i++) {
    const row = tableData[i];
    const level = row[0];
    
    if (level === 'APP') {
      // Close previous week if exists
      if (weekStartRow !== null && i > weekStartRow + 1) {
        weekGroups.push({start: weekStartRow + 1, count: i - weekStartRow - 1});
      }
      // Close previous app if exists
      if (appStartRow !== null && i > appStartRow + 1) {
        appGroups.push({start: appStartRow + 1, count: i - appStartRow - 1});
      }
      
      currentApp = row[1];
      appStartRow = i;
      weekStartRow = null;
    } else if (level === 'WEEK') {
      // Close previous week if exists
      if (weekStartRow !== null && i > weekStartRow + 1) {
        weekGroups.push({start: weekStartRow + 1, count: i - weekStartRow - 1});
      }
      weekStartRow = i;
    }
  }
  
  // Close last week and app
  if (weekStartRow !== null && tableData.length > weekStartRow + 1) {
    weekGroups.push({start: weekStartRow + 1, count: tableData.length - weekStartRow - 1});
  }
  if (appStartRow !== null && tableData.length > appStartRow + 1) {
    appGroups.push({start: appStartRow + 1, count: tableData.length - appStartRow - 1});
  }
  
  // Apply grouping - weeks first (deeper level)
  weekGroups.forEach(group => {
    try {
      sheet.getRange(group.start, 1, group.count, numCols).shiftRowGroupDepth(1);
      sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
    } catch (e) {
      // Ignore grouping errors
    }
  });
  
  // Then apps
  appGroups.forEach(group => {
    try {
      sheet.getRange(group.start, 1, group.count, numCols).shiftRowGroupDepth(1);
      sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
    } catch (e) {
      // Ignore grouping errors
    }
  });
}

/**
 * Create grouping from app data structure
 */
function createGroupingFromAppData(sheet, appData, numCols) {
  let rowPointer = 2; // Start from second row (first is headers)

  // Sort apps by name
  const sortedApps = Object.keys(appData).sort((a, b) => 
    appData[a].appName.localeCompare(appData[b].appName)
  );

  sortedApps.forEach(appKey => {
    const app = appData[appKey];
    const appRow = rowPointer; // APP row
    rowPointer++; // Move to first WEEK

    // Process weeks within app
    const sortedWeeks = Object.keys(app.weeks).sort();
    sortedWeeks.forEach(weekKey => {
      const week = app.weeks[weekKey];
      const weekRow = rowPointer; // WEEK row
      rowPointer++; // Take this row

      const campaignCount = week.campaigns.length;
      if (campaignCount > 0) {
        try {
          // Group campaigns under this week
          sheet.getRange(weekRow + 1, 1, campaignCount, numCols)
               .shiftRowGroupDepth(1);
          sheet.getRange(weekRow + 1, 1, campaignCount, 1)
               .collapseGroups();
        } catch (e) {
          // Ignore grouping errors
        }
        rowPointer += campaignCount; // Skip all CAMPAIGN rows
      }
    });

    // Group all WEEK+CAMPAIGN rows within app
    const totalRowsInApp = rowPointer - appRow - 1; // -1 to exclude APP row itself
    if (totalRowsInApp > 0) {
      try {
        sheet.getRange(appRow + 1, 1, totalRowsInApp, numCols)
             .shiftRowGroupDepth(1);
        sheet.getRange(appRow + 1, 1, totalRowsInApp, 1)
             .collapseGroups();
      } catch (e) {
        // Ignore grouping errors
      }
    }
  });
}

/**
 * Create enhanced pivot table for specific project
 */
function createProjectPivotTable(projectName, appData) {
  // Set current project context
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    createEnhancedPivotTable(appData);
  } finally {
    // Restore original project context
    setCurrentProject(originalProject);
  }
}