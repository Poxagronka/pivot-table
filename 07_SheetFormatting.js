function createEnhancedPivotTable(appData) { createUnifiedPivotTable(appData); }
function createOverallPivotTable(appData) { createUnifiedPivotTable(appData); }
function createIncentTrafficPivotTable(networkData) { createUnifiedPivotTable(networkData); }

function createUnifiedPivotTable(data) {
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

  console.log(`${CURRENT_PROJECT}: Starting optimized table creation...`);

  const initialEROASCache = new InitialEROASCache();
  initialEROASCache.recordInitialValuesFromData(data);

  const wow = CURRENT_PROJECT === 'INCENT_TRAFFIC' ? 
    calculateIncentTrafficWoWMetrics(data) : 
    calculateWoWMetrics(data);
  
  const headers = getUnifiedHeaders();
  const tableData = [headers];
  const formatData = [];

  buildUnifiedTable(data, tableData, formatData, wow, initialEROASCache);

  console.log(`${CURRENT_PROJECT}: Writing ${tableData.length} rows in batches...`);
  writeBatchedData(sheet, tableData, headers.length);
  
  console.log(`${CURRENT_PROJECT}: Applying formatting...`);
  applyOptimizedFormatting(sheet, tableData.length, headers.length, formatData, data);
  
  console.log(`${CURRENT_PROJECT}: Creating row grouping...`);
  createUnifiedRowGrouping(sheet, tableData, data);
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  
  const endTime = Date.now();
  console.log(`${CURRENT_PROJECT}: Table creation completed in ${(endTime - startTime) / 1000}s`);
}

function writeBatchedData(sheet, tableData, numCols) {
  const BATCH_SIZE = 1000;
  const totalRows = tableData.length;
  
  if (totalRows <= BATCH_SIZE) {
    const range = sheet.getRange(1, 1, totalRows, numCols);
    range.setValues(tableData);
    return;
  }
  
  let currentRow = 1;
  let batchCount = 0;
  
  while (currentRow <= totalRows) {
    const remainingRows = totalRows - currentRow + 1;
    const batchSize = Math.min(BATCH_SIZE, remainingRows);
    const batchData = tableData.slice(currentRow - 1, currentRow - 1 + batchSize);
    
    console.log(`Writing batch ${++batchCount}: rows ${currentRow}-${currentRow + batchSize - 1}`);
    
    try {
      const range = sheet.getRange(currentRow, 1, batchSize, numCols);
      range.setValues(batchData);
    } catch (e) {
      console.error(`Error writing batch ${batchCount}:`, e);
      if (e.toString().includes('timed out')) {
        console.log('Write timeout, waiting and retrying...');
        Utilities.sleep(5000);
        const range = sheet.getRange(currentRow, 1, batchSize, numCols);
        range.setValues(batchData);
      } else {
        throw e;
      }
    }
    
    currentRow += batchSize;
    
    if (currentRow <= totalRows) {
      console.log(`Pausing between write batches...`);
      Utilities.sleep(1000);
      SpreadsheetApp.flush();
    }
  }
  
  console.log(`Completed writing ${totalRows} rows in ${batchCount} batches`);
}

function applyOptimizedFormatting(sheet, numRows, numCols, formatData, appData) {
  const formatStartTime = Date.now();
  
  try {
    console.log('Applying header formatting...');
    applyHeaderFormatting(sheet, numCols);
    
    console.log('Setting column widths...');
    setColumnWidths(sheet);

    if (numRows > 1) {
      console.log('Applying data range formatting...');
      applyDataRangeFormatting(sheet, numRows, numCols);
      
      console.log('Applying row type formatting...');
      applyRowTypeFormattingOptimized(sheet, numRows, numCols, formatData);
      
      console.log('Applying optimized conditional formatting...');
      applyOptimizedConditionalFormatting(sheet, numRows);
      
      console.log('Applying eROAS rich text formatting...');
      applyEROASRichTextFormattingOptimized(sheet, numRows);
    }
    
    sheet.hideColumns(1);
    sheet.hideColumns(13, 1);
    sheet.hideColumns(14, 1);
    
    const formatEndTime = Date.now();
    console.log(`${CURRENT_PROJECT}: Formatting completed in ${(formatEndTime - formatStartTime) / 1000}s`);
  } catch (e) {
    console.error(`Error in formatting: ${e}`);
    throw e;
  }
}

function applyHeaderFormatting(sheet, numCols) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground(COLORS.HEADER.background)
    .setFontColor(COLORS.HEADER.fontColor)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setWrap(true);
}

function setColumnWidths(sheet) {
  const columnWidths = TABLE_CONFIG.COLUMN_WIDTHS;
  columnWidths.forEach(col => {
    try {
      sheet.setColumnWidth(col.c, col.w);
    } catch (e) {
      console.error(`Error setting column width ${col.c}:`, e);
    }
  });
}

function applyDataRangeFormatting(sheet, numRows, numCols) {
  try {
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
    
    console.log('Applying number formats...');
    sheet.getRange(2, 5, numRows - 1, 1).setNumberFormat('$0');
    sheet.getRange(2, 8, numRows - 1, 1).setNumberFormat('$0.0');
    sheet.getRange(2, 10, numRows - 1, 1).setNumberFormat('0.0');
    sheet.getRange(2, 13, numRows - 1, 1).setNumberFormat('$0.0');
    sheet.getRange(2, 16, numRows - 1, 1).setNumberFormat('$0');
  } catch (e) {
    console.error('Error in data range formatting:', e);
  }
}

function applyRowTypeFormattingOptimized(sheet, numRows, numCols, formatData) {
  const rowTypeMap = { app: [], week: [], sourceApp: [], campaign: [], hyperlink: [], network: [] };
  formatData.forEach(item => {
    if (item.type === 'APP') rowTypeMap.app.push(item.row);
    if (item.type === 'WEEK') rowTypeMap.week.push(item.row);
    if (item.type === 'SOURCE_APP') rowTypeMap.sourceApp.push(item.row);
    if (item.type === 'CAMPAIGN') rowTypeMap.campaign.push(item.row);
    if (item.type === 'NETWORK') rowTypeMap.network.push(item.row);
    if (item.type === 'HYPERLINK') rowTypeMap.hyperlink.push(item.row);
  });

  if (rowTypeMap.app.length > 0) {
    console.log(`Formatting ${rowTypeMap.app.length} app rows...`);
    batchFormatRowsOptimized(sheet, rowTypeMap.app, numCols, (ranges) => {
      ranges.forEach(range => {
        if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
          range.setBackground(COLORS.CAMPAIGN_ROW.background)
               .setFontWeight('normal')
               .setFontSize(9);
        } else {
          range.setBackground(COLORS.APP_ROW.background)
               .setFontColor(COLORS.APP_ROW.fontColor)
               .setFontWeight('bold')
               .setFontSize(10);
        }
      });
    });
  }

  if (rowTypeMap.week.length > 0) {
    console.log(`Formatting ${rowTypeMap.week.length} week rows...`);
    batchFormatRowsOptimized(sheet, rowTypeMap.week, numCols, (ranges) => {
      ranges.forEach(range => range.setBackground(COLORS.WEEK_ROW.background).setFontSize(10));
    });
  }

  if (rowTypeMap.sourceApp.length > 0) {
    console.log(`Formatting ${rowTypeMap.sourceApp.length} source app rows...`);
    batchFormatRowsOptimized(sheet, rowTypeMap.sourceApp, numCols, (ranges) => {
      ranges.forEach(range => range.setBackground(COLORS.SOURCE_APP_ROW.background).setFontSize(10));
    });
  }

  if (rowTypeMap.campaign.length > 0) {
    console.log(`Formatting ${rowTypeMap.campaign.length} campaign rows...`);
    batchFormatRowsOptimized(sheet, rowTypeMap.campaign, numCols, (ranges) => {
      ranges.forEach(range => range.setBackground(COLORS.CAMPAIGN_ROW.background).setFontSize(9));
    });
  }

  if (rowTypeMap.network.length > 0) {
    console.log(`Formatting ${rowTypeMap.network.length} network rows...`);
    batchFormatRowsOptimized(sheet, rowTypeMap.network, numCols, (ranges) => {
      ranges.forEach(range => {
        if (CURRENT_PROJECT === 'OVERALL') {
          range.setBackground(COLORS.CAMPAIGN_ROW.background)
               .setFontWeight('normal')
               .setFontSize(9);
        } else {
          range.setBackground(COLORS.APP_ROW.background)
               .setFontColor(COLORS.APP_ROW.fontColor)
               .setFontWeight('bold')
               .setFontSize(10);
        }
      });
    });
  }

  if (rowTypeMap.hyperlink.length > 0 && CURRENT_PROJECT === 'TRICKY') {
    console.log(`Formatting ${rowTypeMap.hyperlink.length} hyperlink rows...`);
    try {
      rowTypeMap.hyperlink.forEach(r => {
        const linkCell = sheet.getRange(r, 2);
        linkCell.setFontColor('#000000').setFontLine('none');
      });
    } catch (e) {
      console.error('Error formatting hyperlinks:', e);
    }
  }
}

function batchFormatRowsOptimized(sheet, rows, numCols, formatFunction) {
  const BATCH_SIZE = 500;
  
  for (let i = 0; i < rows.length; i += BATCH_SIZE) {
    const batchRows = rows.slice(i, i + BATCH_SIZE);
    const ranges = [];
    
    try {
      batchRows.forEach(rowNum => {
        ranges.push(sheet.getRange(rowNum, 1, 1, numCols));
      });
      
      formatFunction(ranges);
      
      if (i + BATCH_SIZE < rows.length) {
        Utilities.sleep(100);
        SpreadsheetApp.flush();
      }
    } catch (e) {
      console.error(`Error in batch format (rows ${i}-${i + batchRows.length}):`, e);
      if (e.toString().includes('timed out')) {
        Utilities.sleep(3000);
        formatFunction(ranges);
      }
    }
  }
}

function applyOptimizedConditionalFormatting(sheet, numRows) {
  const rules = [];
  
  if (numRows <= 1) return;
  
  try {
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

    const eroasColumn = 15;
    const eroasRange = sheet.getRange(2, eroasColumn, numRows - 1, 1);
    
    const defaultTarget = getTargetEROAS(CURRENT_PROJECT, '');
    
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(NOT(ISBLANK(O2:O)), VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(O2:O,"→",REPT(" ",100)),100)),"%","")) >= ${defaultTarget})`)
        .setBackground(COLORS.POSITIVE.background)
        .setFontColor(COLORS.POSITIVE.fontColor)
        .setRanges([eroasRange]).build()
    );
    
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(NOT(ISBLANK(O2:O)), VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(O2:O,"→",REPT(" ",100)),100)),"%","")) >= 120, VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(O2:O,"→",REPT(" ",100)),100)),"%","")) < ${defaultTarget})`)
        .setBackground(COLORS.WARNING.background)
        .setFontColor(COLORS.WARNING.fontColor)
        .setRanges([eroasRange]).build()
    );
    
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(NOT(ISBLANK(O2:O)), VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(O2:O,"→",REPT(" ",100)),100)),"%","")) < 120)`)
        .setBackground(COLORS.NEGATIVE.background)
        .setFontColor(COLORS.NEGATIVE.fontColor)
        .setRanges([eroasRange]).build()
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
    
    console.log(`Applying ${rules.length} conditional format rules in optimized batches...`);
    const RULES_BATCH_SIZE = 200;
    
    for (let i = 0; i < rules.length; i += RULES_BATCH_SIZE) {
      const batchRules = rules.slice(i, i + RULES_BATCH_SIZE);
      
      try {
        sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat(batchRules));
        
        if (i + RULES_BATCH_SIZE < rules.length) {
          Utilities.sleep(200);
          SpreadsheetApp.flush();
        }
      } catch (e) {
        console.error(`Error applying conditional format rules batch ${i / RULES_BATCH_SIZE + 1}:`, e);
        if (e.toString().includes('timed out')) {
          console.log('Conditional format timeout, waiting and retrying...');
          Utilities.sleep(5000);
          sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat(batchRules));
        } else {
          console.log('Skipping problematic conditional format batch');
        }
      }
    }
    
    console.log(`Applied ${rules.length} conditional format rules successfully`);
  } catch (e) {
    console.error('Error in conditional formatting:', e);
  }
}

function applyEROASRichTextFormattingOptimized(sheet, numRows) {
  if (numRows <= 1) return;
  
  const eroasColumn = 15;
  const BATCH_SIZE = 500;
  
  for (let startRow = 2; startRow <= numRows; startRow += BATCH_SIZE) {
    const endRow = Math.min(startRow + BATCH_SIZE - 1, numRows);
    const batchSize = endRow - startRow + 1;
    
    try {
      const range = sheet.getRange(startRow, eroasColumn, batchSize, 1);
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
      
      if (endRow < numRows) {
        Utilities.sleep(200);
      }
    } catch (e) {
      console.error(`Error applying rich text formatting to rows ${startRow}-${endRow}:`, e);
      if (e.toString().includes('timed out')) {
        console.log('Rich text timeout, skipping this batch');
        Utilities.sleep(2000);
      }
    }
  }
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