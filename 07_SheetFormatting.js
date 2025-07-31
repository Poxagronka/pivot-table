function createEnhancedPivotTable(appData) { 
  return createUnifiedPivotTable(appData); 
}

function createOverallPivotTable(appData) { 
  return createUnifiedPivotTable(appData); 
}

function createIncentTrafficPivotTable(networkData) { 
  return createUnifiedPivotTable(networkData); 
}

function createUnifiedPivotTable(data) {
  const overallStartTime = Date.now();
  
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);
  else sheet.clear();

  if (!data || Object.keys(data).length === 0) {
    const headers = getUnifiedHeaders();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return headers.length;
  }

  const formatStartTime = Date.now();
  const initialMetricsCache = new InitialMetricsCache();
  initialMetricsCache.recordInitialValuesFromData(data);

  const wow = CURRENT_PROJECT === 'INCENT_TRAFFIC' ? 
    calculateIncentTrafficWoWMetrics(data) : 
    calculateWoWMetrics(data);
  
  const headers = getUnifiedHeaders();
  const tableData = [headers];
  const formatData = [];
  buildUnifiedTable(data, tableData, formatData, wow, initialMetricsCache);

  const range = sheet.getRange(1, 1, tableData.length, headers.length);
  range.setValues(tableData);
  const formatTime = Date.now() - formatStartTime;
  
  const groupingStartTime = Date.now();
  applyOptimizedFormatting(sheet, tableData.length, headers.length, formatData, data);
  const formatFinishTime = Date.now();
  
  createUnifiedRowGrouping(sheet, tableData, data);
  const groupingTime = Date.now() - groupingStartTime;
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  
  logDebugTiming({ 
    format: formatFinishTime - groupingStartTime, 
    grouping: groupingTime 
  });
  
  return tableData.length;
}

function hexToRgb(hex) {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    red: parseInt(result[1], 16) / 255,
    green: parseInt(result[2], 16) / 255,
    blue: parseInt(result[3], 16) / 255
  } : { red: 1, green: 1, blue: 1 };
}

function applyOptimizedFormatting(sheet, numRows, numCols, formatData, appData) {
  const startTime = Date.now();
  
  try {
    const spreadsheetId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();
    
    const columnWidths = [
      { c: 1, w: 50 }, { c: 2, w: 120 }, { c: 3, w: 50 }, { c: 4, w: 80 },
      { c: 5, w: 65 }, { c: 6, w: 70 }, { c: 7, w: 70 }, { c: 8, w: 60 },
      { c: 9, w: 85 }, { c: 10, w: 60 }, { c: 11, w: 60 }, { c: 12, w: 60 },
      { c: 13, w: 70 }, { c: 14, w: 50 }, { c: 15, w: 85 }, { c: 16, w: 80 },
      { c: 17, w: 130 }, { c: 18, w: 200 }
    ];
    
    columnWidths.forEach(col => {
      sheet.setColumnWidth(col.c, col.w);
    });

    const headerRange = sheet.getRange(1, 1, 1, numCols);
    headerRange
      .setBackground('#4285f4')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontSize(10)
      .setWrap(true);

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

      const eprofitRange = sheet.getRange(2, 16, numRows - 1, 1);
      eprofitRange.setHorizontalAlignment('right');
    }
    
    const rowTypeMap = { app: [], week: [], sourceApp: [], campaign: [], network: [], hyperlink: [] };
    
    formatData.forEach(item => {
      if (item.type === 'APP') rowTypeMap.app.push(item.row);
      else if (item.type === 'WEEK') rowTypeMap.week.push(item.row);
      else if (item.type === 'SOURCE_APP') rowTypeMap.sourceApp.push(item.row);
      else if (item.type === 'CAMPAIGN') rowTypeMap.campaign.push(item.row);
      else if (item.type === 'NETWORK') rowTypeMap.network.push(item.row);
      else if (item.type === 'HYPERLINK') rowTypeMap.hyperlink.push(item.row);
    });

    if (rowTypeMap.app.length > 0) {
      const appRanges = createOptimizedRanges(sheet, rowTypeMap.app, numCols);
      appRanges.forEach(range => {
        range.setBackground('#e8f0fe')
             .setFontColor('#1a73e8')
             .setFontWeight('bold')
             .setFontSize(11);
      });
    }

    if (rowTypeMap.week.length > 0) {
      const weekRanges = createOptimizedRanges(sheet, rowTypeMap.week, numCols);
      weekRanges.forEach(range => {
        range.setBackground('#fce8e6')
             .setFontColor('#d93025')
             .setFontWeight('bold')
             .setFontSize(10);
      });
    }

    if (rowTypeMap.sourceApp.length > 0) {
      const sourceAppRanges = createOptimizedRanges(sheet, rowTypeMap.sourceApp, numCols);
      sourceAppRanges.forEach(range => {
        range.setBackground('#e6f4ea')
             .setFontColor('#137333')
             .setFontWeight('bold')
             .setFontSize(10);
      });
    }

    if (rowTypeMap.campaign.length > 0) {
      const campaignRanges = createOptimizedRanges(sheet, rowTypeMap.campaign, numCols);
      campaignRanges.forEach(range => {
        range.setBackground('#ffffff')
             .setFontColor('#5f6368')
             .setFontWeight('normal')
             .setFontSize(9);
      });
    }

    if (rowTypeMap.network.length > 0) {
      const networkRanges = createOptimizedRanges(sheet, rowTypeMap.network, numCols);
      if (CURRENT_PROJECT === 'OVERALL') {
        networkRanges.forEach(range => {
          range.setBackground('#ffffff')
               .setFontWeight('normal')
               .setFontSize(9);
        });
      } else {
        networkRanges.forEach(range => {
          range.setBackground('#d1e7fe')
               .setFontColor('#000000')
               .setFontWeight('bold')
               .setFontSize(10);
        });
      }
    }
    
    if (rowTypeMap.hyperlink.length > 0 && CURRENT_PROJECT === 'TRICKY') {
      try {
        const validHyperlinkRows = rowTypeMap.hyperlink.filter(row => row >= 2 && row <= numRows);
        if (validHyperlinkRows.length > 0) {
          validHyperlinkRows.forEach(row => {
            try {
              const hyperlinkRange = sheet.getRange(row, 2, 1, 1);
              hyperlinkRange.setFontColor('#000000').setFontLine('none');
            } catch (e) {
              console.error(`Error formatting hyperlink row ${row}:`, e);
            }
          });
        }
      } catch (e) {
        console.error('Error in hyperlink formatting section:', e);
      }
    }

    if (numRows > 1) {
      const numberFormatOperations = [
        { range: sheet.getRange(2, 8, numRows - 1, 1), format: '$0.0' },
        { range: sheet.getRange(2, 10, numRows - 1, 1), format: '0.0' },
        { range: sheet.getRange(2, 13, numRows - 1, 1), format: '$0.0' },
        { range: sheet.getRange(2, 16, numRows - 1, 1), format: '$0.0' }
      ];
      
      numberFormatOperations.forEach(op => op.range.setNumberFormat(op.format));
    }

    applyOptimizedConditionalFormatting(sheet, numRows, appData);
    applyOptimizedEROASFormatting(sheet, numRows);
    
    sheet.hideColumns(1);
    sheet.hideColumns(13, 1);
    sheet.hideColumns(14, 1);
    sheet.hideColumns(3);
    
  } catch (e) {
    console.error('Error in applyOptimizedFormatting:', e);
    throw e;
  }
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

function applyOptimizedEROASFormatting(sheet, numRows) {
  if (numRows <= 1) return;
  
  try {
    const spreadsheetId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();
    const eroasColumn = 14;
    const eprofitColumn = 15;
    
    const range = sheet.getRange(2, 1, numRows - 1, eprofitColumn + 2);
    const allData = range.getValues();
    
    const requests = [];
    
    allData.forEach((row, index) => {
      const level = row[0];
      const eroasValue = row[eroasColumn];
      const eprofitValue = row[eprofitColumn];
      
      const rowIndex = index + 1;
      
      let baseFontSize = 10;
      
      switch (level) {
        case 'APP':
          if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
            baseFontSize = 9;
          } else {
            baseFontSize = 10;
          }
          break;
          
        case 'WEEK':
          baseFontSize = 9;
          break;
          
        case 'SOURCE_APP':
          baseFontSize = 9;
          break;
          
        case 'CAMPAIGN':
          baseFontSize = 8;
          break;
          
        case 'NETWORK':
          baseFontSize = 9;
          break;
          
        default:
          baseFontSize = 9;
      }
      
      const smallerFontSize = Math.max(7, baseFontSize - 1);
      
      if (eroasValue && typeof eroasValue === 'string' && eroasValue.includes('→')) {
        const arrowIndex = eroasValue.indexOf('→');
        
        if (arrowIndex > 0) {
          requests.push({
            updateCells: {
              range: {
                sheetId: sheetId,
                startRowIndex: rowIndex,
                endRowIndex: rowIndex + 1,
                startColumnIndex: eroasColumn,
                endColumnIndex: eroasColumn + 1
              },
              rows: [{
                values: [{
                  userEnteredValue: { stringValue: eroasValue },
                  textFormatRuns: [
                    {
                      startIndex: 0,
                      format: {
                        foregroundColor: { red: 0.5, green: 0.5, blue: 0.5 },
                        fontSize: smallerFontSize
                      }
                    },
                    {
                      startIndex: arrowIndex,
                      format: {
                        fontSize: baseFontSize
                      }
                    }
                  ]
                }]
              }],
              fields: 'userEnteredValue,textFormatRuns'
            }
          });
        }
      }
      
      if (eprofitValue && typeof eprofitValue === 'string' && eprofitValue.includes('→')) {
        const arrowIndex = eprofitValue.indexOf('→');
        
        if (arrowIndex > 0) {
          requests.push({
            updateCells: {
              range: {
                sheetId: sheetId,
                startRowIndex: rowIndex,
                endRowIndex: rowIndex + 1,
                startColumnIndex: eprofitColumn,
                endColumnIndex: eprofitColumn + 1
              },
              rows: [{
                values: [{
                  userEnteredValue: { stringValue: eprofitValue },
                  textFormatRuns: [
                    {
                      startIndex: 0,
                      format: {
                        foregroundColor: { red: 0.5, green: 0.5, blue: 0.5 },
                        fontSize: smallerFontSize
                      }
                    },
                    {
                      startIndex: arrowIndex,
                      format: {
                        fontSize: baseFontSize
                      }
                    }
                  ]
                }]
              }],
              fields: 'userEnteredValue,textFormatRuns'
            }
          });
        }
      }
    });
    
    if (requests.length > 0) {
      const batchSize = 500;
      for (let i = 0; i < requests.length; i += batchSize) {
        const batch = requests.slice(i, i + batchSize);
        Sheets.Spreadsheets.batchUpdate({
          requests: batch
        }, spreadsheetId);
        
        if (i + batchSize < requests.length) {
          Utilities.sleep(50);
        }
      }
    }
    
  } catch (e) {
    console.error('Error applying optimized eROAS/eProfit formatting:', e);
  }
}

function applyOptimizedConditionalFormatting(sheet, numRows, appData) {
  try {
    if (numRows <= 1) return;
    
    const spreadsheetId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();
    
    const conditionalFormatRequests = [];
    let ruleIndex = 0;
    
    const targetEROAS = getCurrentTargetEROAS();
    const growthColumn = 17;
    
    const ranges = [{
      sheetId: sheetId,
      startRowIndex: 1,
      endRowIndex: numRows,
      startColumnIndex: 14,
      endColumnIndex: 15
    }];
    
    conditionalFormatRequests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: ranges,
          booleanRule: {
            condition: {
              type: 'CUSTOM_FORMULA',
              values: [{
                userEnteredValue: `=AND(NOT(ISBLANK(INDIRECT("O" & ROW()))), IF(ISERROR(SEARCH("→", INDIRECT("O" & ROW()))), VALUE(SUBSTITUTE(INDIRECT("O" & ROW()), "%", "")), VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(INDIRECT("O" & ROW()), "→", REPT(" ", 100)), 100)), "%", ""))) >= ${targetEROAS})`
              }]
            },
            format: {
              backgroundColor: hexToRgb('#d1f2eb'),
              textFormat: { foregroundColor: hexToRgb('#0c5460') }
            }
          }
        },
        index: ruleIndex++
      }
    });
    
    conditionalFormatRequests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: ranges,
          booleanRule: {
            condition: {
              type: 'CUSTOM_FORMULA',
              values: [{
                userEnteredValue: `=AND(NOT(ISBLANK(INDIRECT("O" & ROW()))), IF(ISERROR(SEARCH("→", INDIRECT("O" & ROW()))), VALUE(SUBSTITUTE(INDIRECT("O" & ROW()), "%", "")), VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(INDIRECT("O" & ROW()), "→", REPT(" ", 100)), 100)), "%", ""))) >= 120, IF(ISERROR(SEARCH("→", INDIRECT("O" & ROW()))), VALUE(SUBSTITUTE(INDIRECT("O" & ROW()), "%", "")), VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(INDIRECT("O" & ROW()), "→", REPT(" ", 100)), 100)), "%", ""))) < ${targetEROAS})`
              }]
            },
            format: {
              backgroundColor: hexToRgb('#fff3cd'),
              textFormat: { foregroundColor: hexToRgb('#856404') }
            }
          }
        },
        index: ruleIndex++
      }
    });
    
    conditionalFormatRequests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: ranges,
          booleanRule: {
            condition: {
              type: 'CUSTOM_FORMULA',
              values: [{
                userEnteredValue: `=AND(NOT(ISBLANK(INDIRECT("O" & ROW()))), IF(ISERROR(SEARCH("→", INDIRECT("O" & ROW()))), VALUE(SUBSTITUTE(INDIRECT("O" & ROW()), "%", "")), VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(INDIRECT("O" & ROW()), "→", REPT(" ", 100)), 100)), "%", ""))) < 120)`
              }]
            },
            format: {
              backgroundColor: hexToRgb('#f8d7da'),
              textFormat: { foregroundColor: hexToRgb('#721c24') }
            }
          }
        },
        index: ruleIndex++
      }
    });
    
    const statusColors = {
      "🟢 Strong Growth": { background: "#d4edda", fontColor: "#155724" },
      "🟠 Declining": { background: "#f8d7da", fontColor: "#721c24" },
      "🔵 Scale - Stable": { background: "#cce5ff", fontColor: "#0056b3" },
      "🔵 Scale - Growing": { background: "#b3d9ff", fontColor: "#004080" },
      "🔵 Scale - Declining": { background: "#e6f3ff", fontColor: "#0073e6" },
      "⚫ Baseline - Stable": { background: "#f0f0f0", fontColor: "#666666" },
      "⚫ Baseline - Growing": { background: "#e8e8e8", fontColor: "#555555" },
      "⚫ Baseline - Declining": { background: "#f5f5f5", fontColor: "#777777" },
      "🟤 Baseline - Proportional": { background: "#f0f0f0", fontColor: "#666666" },
      "🟡 Efficiency Improvement": { background: "#e8f5e8", fontColor: "#2d5a2d" },
      "🟡 Minimal Growth": { background: "#fff8e1", fontColor: "#f57f17" },
      "🟡 Moderate Decline": { background: "#fff3cd", fontColor: "#856404" },
      "⚪ Stable": { background: "#f5f5f5", fontColor: "#616161" },
      "First Week": { background: "#e0e0e0", fontColor: "#757575" }
    };
    
    Object.entries(statusColors).forEach(([status, colors]) => {
      conditionalFormatRequests.push({
        addConditionalFormatRule: {
          rule: {
            ranges: [{
              sheetId: sheetId,
              startRowIndex: 1,
              endRowIndex: numRows,
              startColumnIndex: growthColumn - 1,
              endColumnIndex: growthColumn
            }],
            booleanRule: {
              condition: {
                type: 'TEXT_CONTAINS',
                values: [{ userEnteredValue: status }]
              },
              format: {
                backgroundColor: hexToRgb(colors.background),
                textFormat: { foregroundColor: hexToRgb(colors.fontColor) }
              }
            }
          },
          index: ruleIndex++
        }
      });
    });
    
    const batchUpdateRequest = {
      requests: conditionalFormatRequests
    };
    
    Sheets.Spreadsheets.batchUpdate(batchUpdateRequest, spreadsheetId);
    
  } catch (e) {
    console.error('Error applying conditional formatting:', e);
  }
}

function createProjectPivotTable(projectName, appData) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    return createUnifiedPivotTable(appData);
  } finally {
    setCurrentProject(originalProject);
  }
}