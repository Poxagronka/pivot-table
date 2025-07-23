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

function createProjectPivotTable(projectName, appData) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    createUnifiedPivotTable(appData);
  } finally {
    setCurrentProject(originalProject);
  }
}