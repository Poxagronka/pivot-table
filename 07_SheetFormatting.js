function createEnhancedPivotTable(appData) { createUnifiedPivotTable(appData); }
function createOverallPivotTable(appData) { createUnifiedPivotTable(appData); }
function createIncentTrafficPivotTable(networkData) { createUnifiedPivotTable(networkData); }

function createUnifiedPivotTable(data) {
  console.log('📊 Starting pivot table creation...');
  const startTime = Date.now();
  
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

  console.log(`⏱️ Initial eROAS cache... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
  const initialEROASCache = new InitialEROASCache();
  initialEROASCache.recordInitialValuesFromData(data);

  console.log(`⏱️ WoW calculations starting... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
  const wow = CURRENT_PROJECT === 'INCENT_TRAFFIC' ? 
    calculateIncentTrafficWoWMetrics(data) : 
    calculateWoWMetrics(data);
  
  console.log(`⏱️ Building table data... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
  const headers = getUnifiedHeaders();
  const tableData = [headers];
  const formatData = [];
  buildUnifiedTable(data, tableData, formatData, wow, initialEROASCache);

  console.log(`⏱️ Writing to sheet... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
  const range = sheet.getRange(1, 1, tableData.length, headers.length);
  range.setValues(tableData);
  
  console.log(`⏱️ Applying formatting... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
  applyOptimizedFormatting(sheet, tableData.length, headers.length, formatData, data);
  
  console.log(`⏱️ Creating row grouping... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
  createUnifiedRowGrouping(sheet, tableData, data);
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  console.log(`✅ Pivot table completed in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
}

function applyOptimizedFormatting(sheet, numRows, numCols, formatData, appData) {
  const startTime = Date.now();
  console.log('🎨 Starting optimized formatting...');
  
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
  console.log(`⏱️ Setting column widths... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
  
  const batchOperations = [];
  columnWidths.forEach(col => {
    batchOperations.push(() => sheet.setColumnWidth(col.c, col.w));
  });
  batchOperations.forEach(op => op());

  if (numRows > 1) {
    console.log(`⏱️ Basic formatting... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
    
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

  console.log(`⏱️ Row type formatting... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
  
  const rowTypeMap = { app: [], week: [], sourceApp: [], campaign: [], hyperlink: [], network: [] };
  formatData.forEach(item => {
    if (item.type === 'APP') rowTypeMap.app.push(item.row);
    if (item.type === 'WEEK') rowTypeMap.week.push(item.row);
    if (item.type === 'SOURCE_APP') rowTypeMap.sourceApp.push(item.row);
    if (item.type === 'CAMPAIGN') rowTypeMap.campaign.push(item.row);
    if (item.type === 'NETWORK') rowTypeMap.network.push(item.row);
    if (item.type === 'HYPERLINK') rowTypeMap.hyperlink.push(item.row);
  });

  console.log(`⏱️ Batch row formatting - APP: ${rowTypeMap.app.length}, WEEK: ${rowTypeMap.week.length}, SOURCE_APP: ${rowTypeMap.sourceApp.length}, CAMPAIGN: ${rowTypeMap.campaign.length}... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
  
  if (rowTypeMap.app.length > 0) {
    const appRanges = createOptimizedRanges(sheet, rowTypeMap.app, numCols);
    if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
      appRanges.forEach(range => {
        range.setBackground(COLORS.CAMPAIGN_ROW.background)
             .setFontWeight('normal')
             .setFontSize(9);
      });
    } else {
      appRanges.forEach(range => {
        range.setBackground(COLORS.APP_ROW.background)
             .setFontColor(COLORS.APP_ROW.fontColor)
             .setFontWeight('bold')
             .setFontSize(10);
      });
    }
  }

  if (rowTypeMap.week.length > 0) {
    const weekRanges = createOptimizedRanges(sheet, rowTypeMap.week, numCols);
    weekRanges.forEach(range => {
      range.setBackground(COLORS.WEEK_ROW.background).setFontSize(10);
    });
  }

  if (rowTypeMap.sourceApp.length > 0) {
    const sourceAppRanges = createOptimizedRanges(sheet, rowTypeMap.sourceApp, numCols);
    sourceAppRanges.forEach(range => {
      range.setBackground(COLORS.SOURCE_APP_ROW.background).setFontSize(10);
    });
  }

  if (rowTypeMap.campaign.length > 0) {
    const campaignRanges = createOptimizedRanges(sheet, rowTypeMap.campaign, numCols);
    campaignRanges.forEach(range => {
      range.setBackground(COLORS.CAMPAIGN_ROW.background).setFontSize(9);
    });
  }

  if (rowTypeMap.network.length > 0) {
    const networkRanges = createOptimizedRanges(sheet, rowTypeMap.network, numCols);
    if (CURRENT_PROJECT === 'OVERALL') {
      networkRanges.forEach(range => {
        range.setBackground(COLORS.CAMPAIGN_ROW.background)
             .setFontWeight('normal')
             .setFontSize(9);
      });
    } else {
      networkRanges.forEach(range => {
        range.setBackground(COLORS.APP_ROW.background)
             .setFontColor(COLORS.APP_ROW.fontColor)
             .setFontWeight('bold')
             .setFontSize(10);
      });
    }
  }

  if (rowTypeMap.hyperlink.length > 0 && CURRENT_PROJECT === 'TRICKY') {
    console.log(`⏱️ Hyperlink formatting... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
    const hyperlinkRanges = rowTypeMap.hyperlink.map(r => sheet.getRange(r, 2));
    if (hyperlinkRanges.length > 0) {
      sheet.getRangeList(hyperlinkRanges).setFontColor('#000000').setFontLine('none');
    }
  }

  if (numRows > 1) {
    console.log(`⏱️ Number formatting... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
    
    const numberFormatOperations = [
      { range: sheet.getRange(2, 5, numRows - 1, 1), format: '$0.0' },
      { range: sheet.getRange(2, 8, numRows - 1, 1), format: '$0.0' },
      { range: sheet.getRange(2, 10, numRows - 1, 1), format: '0.0' },
      { range: sheet.getRange(2, 13, numRows - 1, 1), format: '$0.0' },
      { range: sheet.getRange(2, 16, numRows - 1, 1), format: '$0.0' }
    ];
    
    numberFormatOperations.forEach(op => op.range.setNumberFormat(op.format));
  }

  console.log(`⏱️ Conditional formatting... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
  applyOptimizedConditionalFormatting(sheet, numRows, appData);
  
  console.log(`⏱️ eROAS rich text... (${((Date.now() - startTime) / 1000).toFixed(1)}s elapsed)`);
  applyEROASRichTextFormatting(sheet, numRows);
  
  sheet.hideColumns(1);
  sheet.hideColumns(13, 1);
  sheet.hideColumns(14, 1);
  
  console.log(`🎨 Optimized formatting completed in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
}

function createOptimizedRanges(sheet, rowNumbers, numCols) {
  if (rowNumbers.length === 0) return [];
  
  const ranges = [];
  const sortedRows = [...rowNumbers].sort((a, b) => a - b);
  
  let start = sortedRows[0];
  let end = start;
  
  for (let i = 1; i < sortedRows.length; i++) {
    if (sortedRows[i] === end + 1) {
      end = sortedRows[i];
    } else {
      ranges.push(sheet.getRange(start, 1, end - start + 1, numCols));
      start = sortedRows[i];
      end = start;
    }
  }
  
  ranges.push(sheet.getRange(start, 1, end - start + 1, numCols));
  return ranges;
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

function applyOptimizedConditionalFormatting(sheet, numRows, appData) {
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
    const data = sheet.getDataRange().getValues();
    
    const eroasRanges = [];
    const eroasRules = [];
    
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
      
      eroasRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${cellAddress})), ${extractValueFormula} >= ${targetEROAS})`)
          .setBackground(COLORS.POSITIVE.background)
          .setFontColor(COLORS.POSITIVE.fontColor)
          .setRanges([cellRange]).build()
      );
      
      eroasRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${cellAddress})), ${extractValueFormula} >= 120, ${extractValueFormula} < ${targetEROAS})`)
          .setBackground(COLORS.WARNING.background)
          .setFontColor(COLORS.WARNING.fontColor)
          .setRanges([cellRange]).build()
      );
      
      eroasRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${cellAddress})), ${extractValueFormula} < 120)`)
          .setBackground(COLORS.NEGATIVE.background)
          .setFontColor(COLORS.NEGATIVE.fontColor)
          .setRanges([cellRange]).build()
      );
    }

    rules.push(...eroasRules);

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

function createProjectPivotTable(projectName, appData) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    createUnifiedPivotTable(appData);
  } finally {
    setCurrentProject(originalProject);
  }
}