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
    const headers = getUnifiedHeaders();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

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
  
  applyOptimizedFormatting(sheet, tableData.length, headers.length, formatData, data);
  
  createUnifiedRowGrouping(sheet, tableData, data);
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  console.log(`Pivot table created in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
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
    
    const columnWidths = TABLE_CONFIG.COLUMN_WIDTHS;
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

    const rowTypeMap = { app: [], week: [], sourceApp: [], campaign: [], hyperlink: [], network: [], country: [] };
    
    // СПЕЦИАЛЬНАЯ ОБРАБОТКА ДЛЯ APPLOVIN_TEST
    if (CURRENT_PROJECT === 'APPLOVIN_TEST') {
      formatData.forEach(item => {
        if (item.type === 'APP') rowTypeMap.app.push(item.row);
        // Меняем местами форматирование для CAMPAIGN и WEEK
        if (item.type === 'CAMPAIGN') rowTypeMap.week.push(item.row);  // Кампании форматируем как недели
        if (item.type === 'WEEK') rowTypeMap.campaign.push(item.row);  // Недели форматируем как кампании
        if (item.type === 'COUNTRY') rowTypeMap.country.push(item.row);
      });
    } else {
      // Стандартная обработка для остальных проектов
      formatData.forEach(item => {
        if (item.type === 'APP') rowTypeMap.app.push(item.row);
        if (item.type === 'WEEK') rowTypeMap.week.push(item.row);
        if (item.type === 'SOURCE_APP') rowTypeMap.sourceApp.push(item.row);
        if (item.type === 'CAMPAIGN') rowTypeMap.campaign.push(item.row);
        if (item.type === 'NETWORK') rowTypeMap.network.push(item.row);
        if (item.type === 'COUNTRY') rowTypeMap.country.push(item.row);
        if (item.type === 'HYPERLINK') rowTypeMap.hyperlink.push(item.row);
      });
    }
    
    // Скрытие колонки GEO для INCENT_TRAFFIC
    if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
      try {
        sheet.hideColumns(4); // Скрываем колонку GEO (4-я колонка)
      } catch (e) {
        console.error('Error hiding GEO column:', e);
      }
    }

    // Далее идет стандартный код применения форматирования без изменений
    if (rowTypeMap.app.length > 0) {
      const appRanges = createOptimizedRanges(sheet, rowTypeMap.app, numCols);
      if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
        appRanges.forEach(range => {
          range.setBackground('#ffffff')
               .setFontWeight('normal')
               .setFontSize(9);
        });
      } else {
        appRanges.forEach(range => {
          range.setBackground('#d1e7fe')
               .setFontColor('#000000')
               .setFontWeight('bold')
               .setFontSize(10);
        });
      }
    }

    if (rowTypeMap.week.length > 0) {
      const weekRanges = createOptimizedRanges(sheet, rowTypeMap.week, numCols);
      weekRanges.forEach(range => {
        if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
          // Для INCENT_TRAFFIC в week попадают кампании - белый фон, размер 10
          range.setBackground('#ffffff').setFontSize(10);
        } else {
          // Стандартное форматирование недель - синий фон, размер 10
          range.setBackground('#e8f0fe').setFontSize(10);
        }
      });
    }

    if (rowTypeMap.sourceApp.length > 0) {
      const sourceAppRanges = createOptimizedRanges(sheet, rowTypeMap.sourceApp, numCols);
      sourceAppRanges.forEach(range => {
        range.setBackground('#f0f8ff').setFontSize(10);
      });
    }

    if (rowTypeMap.campaign.length > 0) {
      const campaignRanges = createOptimizedRanges(sheet, rowTypeMap.campaign, numCols);
      campaignRanges.forEach(range => {
        // Для APPLOVIN_TEST в campaign попадают недели (должен быть размер 10)
        // Для INCENT_TRAFFIC в campaign попадают недели (должен быть размер 9, белый фон)
        if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
          range.setBackground('#ffffff').setFontSize(9);
        } else {
          const fontSize = CURRENT_PROJECT === 'APPLOVIN_TEST' ? 10 : 9;
          range.setBackground('#ffffff').setFontSize(fontSize);
        }
      });
    }

    if (rowTypeMap.country && rowTypeMap.country.length > 0) {
      const countryRanges = createOptimizedRanges(sheet, rowTypeMap.country, numCols);
      if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
        countryRanges.forEach(range => {
          range.setBackground('#f0f8ff')
               .setFontSize(10)
               .setFontWeight('normal');
        });
      } else {
        countryRanges.forEach(range => {
          range.setBackground('#ffffff').setFontSize(9);
        });
      }
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

    // Остальной код остается без изменений...
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
        { range: sheet.getRange(2, 8, numRows - 1, 1), format: '$0.0' },  // CPI
        { range: sheet.getRange(2, 10, numRows - 1, 1), format: '0.0' },  // IPM
        { range: sheet.getRange(2, 13, numRows - 1, 1), format: '$0.0' }, // eARPU
        { range: sheet.getRange(2, 16, numRows - 1, 1), format: '$0.0' }  // eProfit
      ];
      
      numberFormatOperations.forEach(op => op.range.setNumberFormat(op.format));
    }

    applyOptimizedConditionalFormatting(sheet, numRows, appData);
    
    applyOptimizedEROASFormatting(sheet, numRows);
    
    // Существующий код скрытия колонок
    sheet.hideColumns(1);
    sheet.hideColumns(13, 1);
    sheet.hideColumns(14, 1);
    sheet.hideColumns(3);
    
    // Добавляем скрытие GEO для APPLOVIN_TEST
    if (CURRENT_PROJECT === 'APPLOVIN_TEST') {
      sheet.hideColumns(4); // Скрываем колонку GEO (4-я колонка)
    }
    
    console.log(`Formatting completed in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
    
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
  
  const startTime = Date.now();
  
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
          baseFontSize = 10;
          break;
          
        case 'SOURCE_APP':
          baseFontSize = 10;
          break;
          
        case 'CAMPAIGN':
          baseFontSize = 9;
          break;
          
        case 'NETWORK':
          if (CURRENT_PROJECT === 'OVERALL') {
            baseFontSize = 9;
          } else if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
            baseFontSize = 10;
          } else {
            baseFontSize = 10;
          }
          break;
          
        default:
          baseFontSize = 10;
      }
      
      const smallerFontSize = baseFontSize - 1;
      
      // Format eROAS column
      if (eroasValue && typeof eroasValue === 'string' && eroasValue.includes('→')) {
        const arrowIndex = eroasValue.indexOf('→');
        if (arrowIndex !== -1) {
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
      
      // Format eProfit column
      if (eprofitValue && typeof eprofitValue === 'string' && eprofitValue.includes('→')) {
        const arrowIndex = eprofitValue.indexOf('→');
        if (arrowIndex !== -1) {
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
    
    console.log(`eROAS/eProfit formatting completed in ${((Date.now() - startTime) / 1000).toFixed(1)}s (${requests.length} cells)`);
    
  } catch (e) {
    console.error('Error applying optimized eROAS/eProfit formatting:', e);
  }
}

function applyOptimizedConditionalFormatting(sheet, numRows, appData) {
  try {
    const startTime = Date.now();
    
    if (numRows <= 1) return;
    
    const spreadsheetId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();
    
    const conditionalFormatRequests = [];
    
    // 1. Правила для Spend WoW колонки (исправлен порядок и формулы)
    const spendColumn = 6;
    conditionalFormatRequests.push(
      // ПЕРВОЕ правило: отрицательные значения → красный
      {
        addConditionalFormatRule: {
          rule: {
            ranges: [{
              sheetId: sheetId,
              startRowIndex: 1,
              endRowIndex: numRows,
              startColumnIndex: spendColumn - 1,
              endColumnIndex: spendColumn
            }],
            booleanRule: {
              condition: {
                type: 'CUSTOM_FORMULA',
                values: [{ userEnteredValue: '=AND(NOT(ISBLANK($F2)), LEFT($F2,1)="-")' }]
              },
              format: {
                backgroundColor: hexToRgb('#f8d7da'),
                textFormat: { foregroundColor: hexToRgb('#721c24') }
              }
            }
          },
          index: 0
        }
      },
      // ВТОРОЕ правило: положительные значения → зеленый
      {
        addConditionalFormatRule: {
          rule: {
            ranges: [{
              sheetId: sheetId,
              startRowIndex: 1,
              endRowIndex: numRows,
              startColumnIndex: spendColumn - 1,
              endColumnIndex: spendColumn
            }],
            booleanRule: {
              condition: {
                type: 'CUSTOM_FORMULA',
                values: [{ userEnteredValue: '=AND(NOT(ISBLANK($F2)), $F2<>"", LEFT($F2,1)<>"-")' }]
              },
              format: {
                backgroundColor: hexToRgb('#d1f2eb'),
                textFormat: { foregroundColor: hexToRgb('#0c5460') }
              }
            }
          },
          index: 1
        }
      }
    );
    
    // 2. Правила для eROAS колонки - оптимизированный подход
    const eroasColumn = 15;
    const data = sheet.getDataRange().getValues();
    
    const targetGroups = new Map();
    
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
      
      if (!targetGroups.has(targetEROAS)) {
        targetGroups.set(targetEROAS, []);
      }
      targetGroups.get(targetEROAS).push(i + 1);
    }
    
    let ruleIndex = conditionalFormatRequests.length;
    
    targetGroups.forEach((rows, targetEROAS) => {
      const ranges = rows.map(row => ({
        sheetId: sheetId,
        startRowIndex: row - 1,
        endRowIndex: row,
        startColumnIndex: eroasColumn - 1,
        endColumnIndex: eroasColumn
      }));
      
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
    });
    
    // 3. Правила для Profit WoW колонки
    const profitColumn = 17;
    conditionalFormatRequests.push(
      {
        addConditionalFormatRule: {
          rule: {
            ranges: [{
              sheetId: sheetId,
              startRowIndex: 1,
              endRowIndex: numRows,
              startColumnIndex: profitColumn - 1,
              endColumnIndex: profitColumn
            }],
            booleanRule: {
              condition: {
                type: 'CUSTOM_FORMULA',
                values: [{ userEnteredValue: '=AND(ISNUMBER($Q2), $Q2>0)' }]
              },
              format: {
                backgroundColor: hexToRgb('#d1f2eb'),
                textFormat: { foregroundColor: hexToRgb('#0c5460') }
              }
            }
          },
          index: ruleIndex++
        }
      },
      {
        addConditionalFormatRule: {
          rule: {
            ranges: [{
              sheetId: sheetId,
              startRowIndex: 1,
              endRowIndex: numRows,
              startColumnIndex: profitColumn - 1,
              endColumnIndex: profitColumn
            }],
            booleanRule: {
              condition: {
                type: 'CUSTOM_FORMULA',
                values: [{ userEnteredValue: '=AND(ISNUMBER($Q2), $Q2<0)' }]
              },
              format: {
                backgroundColor: hexToRgb('#f8d7da'),
                textFormat: { foregroundColor: hexToRgb('#721c24') }
              }
            }
          },
          index: ruleIndex++
        }
      }
    );
    
    // 4. Правила для Growth Status колонки
    const growthColumn = 18;
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
    
    const endTime = Date.now();
    console.log(`Conditional formatting completed in ${(endTime - startTime) / 1000}s (${conditionalFormatRequests.length} rules)`);
    
  } catch (e) {
    console.error('Error applying conditional formatting:', e);
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