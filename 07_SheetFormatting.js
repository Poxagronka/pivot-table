/**
 * Sheet Formatting and Table Creation - Multi Project Support
 * ИСПРАВЛЕНО: правильная группировка для TRICKY с Apps Database
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
      
      // WEEK row
      formatData.push({ row: tableData.length + 1, type: 'WEEK' });
      
      // For TRICKY project with source app grouping
      if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
        // Calculate week totals from all source apps
        const allCampaigns = [];
        Object.values(week.sourceApps).forEach(sourceApp => {
          allCampaigns.push(...sourceApp.campaigns);
        });
        
        const weekTotals = calculateWeekTotals(allCampaigns);
        
        // Get WoW data for week
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = wow.appWeekWoW[appWeekKey] || {};
        
        const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
        const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
        const status = weekWoW.growthStatus || '';
        
        const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status);
        tableData.push(weekRow);
        
        // Add source app rows and campaigns
        addSourceAppRows(tableData, week.sourceApps, weekKey, wow, formatData);
        
      } else {
        // For other projects without source app grouping
        const weekTotals = calculateWeekTotals(week.campaigns);
        
        // Get WoW data for week
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = wow.appWeekWoW[appWeekKey] || {};
        
        const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
        const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
        const status = weekWoW.growthStatus || '';
        
        const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status);
        tableData.push(weekRow);
        
        // Add campaign rows directly
        addCampaignRows(tableData, week.campaigns, week, weekKey, wow, formatData);
      }
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
 * Add source app rows and their campaigns (for TRICKY project)
 */
function addSourceAppRows(tableData, sourceApps, weekKey, wow, formatData) {
  // Sort source apps by total spend (highest first)
  const sourceAppKeys = Object.keys(sourceApps).sort((a, b) => {
    const totalSpendA = sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
    const totalSpendB = sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
    return totalSpendB - totalSpendA;
  });
  
  sourceAppKeys.forEach(sourceAppKey => {
    const sourceApp = sourceApps[sourceAppKey];
    
    // Calculate source app totals
    const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
    
    // Get WoW data for source app using correct key format
    const sourceAppWoWKey = `${sourceApp.sourceAppId}_${weekKey}`;
    const sourceAppWoW = wow.sourceAppWoW[sourceAppWoWKey] || {};
    
    const spendWoW = sourceAppWoW.spendChangePercent !== undefined ? `${sourceAppWoW.spendChangePercent.toFixed(0)}%` : '';
    const profitWoW = sourceAppWoW.eProfitChangePercent !== undefined ? `${sourceAppWoW.eProfitChangePercent.toFixed(0)}%` : '';
    const status = sourceAppWoW.growthStatus || '';
    
    // SOURCE_APP row
    formatData.push({ row: tableData.length + 1, type: 'SOURCE_APP' });
    
    const sourceAppDisplayName = sourceApp.sourceAppName;
    const sourceAppRow = createSourceAppRow(sourceAppDisplayName, sourceAppTotals, spendWoW, profitWoW, status);
    tableData.push(sourceAppRow);
    
    // Add campaign rows under this source app
    addCampaignRows(tableData, sourceApp.campaigns, { weekStart: weekKey.split('-').join('/'), weekEnd: '' }, weekKey, wow, formatData);
  });
}

/**
 * Create source app row
 */
function createSourceAppRow(sourceAppDisplayName, totals, spendWoW, profitWoW, status) {
  if (CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN') {
    return [
      'SOURCE_APP',
      sourceAppDisplayName,
      '', '',
      totals.totalSpend.toFixed(2),
      spendWoW,
      totals.totalInstalls,
      totals.avgCpi.toFixed(3),
      totals.avgRoas.toFixed(2),
      `${totals.avgRrD1.toFixed(1)}%`,
      `${totals.avgRrD7.toFixed(1)}%`,
      `${totals.avgERoas.toFixed(0)}%`,
      totals.totalProfit.toFixed(2),
      profitWoW,
      status,
      ''
    ];
  } else {
    return [
      'SOURCE_APP',
      sourceAppDisplayName,
      '', '',
      totals.totalSpend.toFixed(2),
      spendWoW,
      totals.totalInstalls,
      totals.avgCpi.toFixed(3),
      totals.avgRoas.toFixed(2),
      totals.avgIpm.toFixed(1),
      totals.avgArpu.toFixed(3),
      `${totals.avgERoas.toFixed(0)}%`,
      totals.totalProfit.toFixed(2),
      profitWoW,
      status,
      ''
    ];
  }
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
  const appRows = [], weekRows = [], sourceAppRows = [], campaignRows = [];
  formatData.forEach(item => {
    if (item.type === 'APP') appRows.push(item.row);
    if (item.type === 'WEEK') weekRows.push(item.row);
    if (item.type === 'SOURCE_APP') sourceAppRows.push(item.row);
    if (item.type === 'CAMPAIGN') campaignRows.push(item.row);
  });

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
    // ИСПРАВЛЕНО: Ссылки только для TRICKY и REGULAR
    let campaignIdValue;
    if (CURRENT_PROJECT === 'TRICKY' || CURRENT_PROJECT === 'REGULAR') {
      campaignIdValue = `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`;
    } else {
      campaignIdValue = campaign.campaignId;
    }
    
    // Получаем WoW данные для кампании по campaignId
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

/**
 * Create campaign row based on project type
 */
function createCampaignRow(campaign, campaignIdValue, spendPct, profitPct, growthStatus) {
  if (CURRENT_PROJECT === 'GOOGLE_ADS' || CURRENT_PROJECT === 'APPLOVIN') {
    return [
      'CAMPAIGN',
      campaign.sourceApp,
      campaignIdValue,
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
      campaignIdValue,
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
 * Create row grouping: App → Week → Source App (TRICKY only) → Campaign
 * ИСПРАВЛЕНО: правильная группировка без ошибок
 */
function createRowGrouping(sheet, tableData, appData) {
  const numCols = getProjectHeaders().length;

  try {
    let rowPointer = 2; // Start from second row (first is headers)

    // Sort apps by name
    const sortedApps = Object.keys(appData).sort((a, b) => 
      appData[a].appName.localeCompare(appData[b].appName)
    );

    sortedApps.forEach(appKey => {
      const app = appData[appKey];
      const appStartRow = rowPointer; // APP row
      rowPointer++; // Move past APP row

      // Process weeks within app
      const sortedWeeks = Object.keys(app.weeks).sort();
      
      sortedWeeks.forEach(weekKey => {
        const week = app.weeks[weekKey];
        const weekStartRow = rowPointer; // WEEK row
        rowPointer++; // Move past WEEK row

        let weekContentRows = 0;

        // For TRICKY project with source app grouping
        if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
          const sourceAppKeys = Object.keys(week.sourceApps);
          
          sourceAppKeys.forEach(sourceAppKey => {
            const sourceApp = week.sourceApps[sourceAppKey];
            const sourceAppStartRow = rowPointer; // SOURCE_APP row
            rowPointer++; // Move past SOURCE_APP row
            
            const campaignCount = sourceApp.campaigns.length;
            rowPointer += campaignCount; // Move past all campaigns
            weekContentRows += 1 + campaignCount; // SOURCE_APP + campaigns
            
            // Group campaigns under source app
            if (campaignCount > 0) {
              try {
                sheet.getRange(sourceAppStartRow + 1, 1, campaignCount, numCols)
                     .shiftRowGroupDepth(1);
                sheet.getRange(sourceAppStartRow + 1, 1, campaignCount, 1)
                     .collapseGroups();
              } catch (e) {
                console.log('Error grouping campaigns under source app:', e);
              }
            }
          });
          
          // Group all source apps + campaigns under week
          if (weekContentRows > 0) {
            try {
              sheet.getRange(weekStartRow + 1, 1, weekContentRows, numCols)
                   .shiftRowGroupDepth(1);
              sheet.getRange(weekStartRow + 1, 1, weekContentRows, 1)
                   .collapseGroups();
            } catch (e) {
              console.log('Error grouping week content:', e);
            }
          }
          
        } else {
          // For other projects without source app grouping
          const campaignCount = week.campaigns ? week.campaigns.length : 0;
          rowPointer += campaignCount; // Move past all campaigns
          weekContentRows = campaignCount;
          
          // Group campaigns under week
          if (campaignCount > 0) {
            try {
              sheet.getRange(weekStartRow + 1, 1, campaignCount, numCols)
                   .shiftRowGroupDepth(1);
              sheet.getRange(weekStartRow + 1, 1, campaignCount, 1)
                   .collapseGroups();
            } catch (e) {
              console.log('Error grouping campaigns under week:', e);
            }
          }
        }
      });

      // Group all weeks + their content under app
      const appContentRows = rowPointer - appStartRow - 1; // -1 to exclude APP row itself
      if (appContentRows > 0) {
        try {
          sheet.getRange(appStartRow + 1, 1, appContentRows, numCols)
               .shiftRowGroupDepth(1);
          sheet.getRange(appStartRow + 1, 1, appContentRows, 1)
               .collapseGroups();
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