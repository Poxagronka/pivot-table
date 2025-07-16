function createEnhancedPivotTable(appData) {
  if (CURRENT_PROJECT === 'TRICKY') {
    createTrickyOptimizedPivotTable(appData);
    return;
  }
  createStandardEnhancedPivotTable(appData);
}

function createTrickyOptimizedPivotTable(appData) {
  console.log('Creating TRICKY optimized pivot table...');
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);

  const wow = calculateWoWMetrics(appData);
  const headers = getUnifiedHeaders();
  const tableData = [headers];
  const formatData = [];
  const hyperlinkData = [];

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
      
      const allCampaigns = [];
      Object.values(week.sourceApps || {}).forEach(sourceApp => {
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
      
      addTrickyOptimizedSourceAppRows(tableData, week.sourceApps, weekKey, wow, formatData, hyperlinkData);
    });
  });

  console.log(`Writing ${tableData.length} rows to sheet...`);
  sheet.getRange(1, 1, tableData.length, headers.length).setValues(tableData);
  
  console.log('Applying advanced formatting...');
  applyAdvancedFormattingBatch(sheet, tableData.length, headers.length, formatData, appData, hyperlinkData);
  
  console.log('Creating optimized grouping...');
  createTrickyOptimizedRowGrouping(sheet, tableData, appData);
  
  sheet.setFrozenRows(1);
  console.log('TRICKY optimized pivot table completed');
}

function addTrickyOptimizedSourceAppRows(tableData, sourceApps, weekKey, wow, formatData, hyperlinkData) {
  if (!sourceApps) return;
  
  const cache = initTrickyOptimizedCache();
  
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
    const appInfo = cache?.appsDbCache[sourceApp.sourceAppId];
    if (appInfo && appInfo.linkApp) {
      sourceAppDisplayName = `=HYPERLINK("${appInfo.linkApp}", "${sourceApp.sourceAppName}")`;
      hyperlinkData.push({ row: tableData.length + 1, col: 2 });
    }
    
    const sourceAppRow = createSourceAppRow(sourceAppDisplayName, sourceAppTotals, spendWoW, profitWoW, status);
    tableData.push(sourceAppRow);
    
    addTrickyOptimizedCampaignRows(tableData, sourceApp.campaigns, weekKey, wow, formatData);
  });
}

function addTrickyOptimizedCampaignRows(tableData, campaigns, weekKey, wow, formatData) {
  campaigns.sort((a, b) => b.spend - a.spend).forEach(campaign => {
    const campaignIdValue = `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`;
    
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

function applyAdvancedFormattingBatch(sheet, numRows, numCols, formatData, appData, hyperlinkData) {
  try {
    const spreadsheetId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();
    const requests = [];
    
    requests.push({
      updateSheetProperties: {
        properties: {
          sheetId: sheetId,
          gridProperties: {
            frozenRowCount: 1,
            hideGridlines: false
          }
        },
        fields: 'gridProperties.frozenRowCount,gridProperties.hideGridlines'
      }
    });
    
    requests.push({
      updateDimensionProperties: {
        range: {
          sheetId: sheetId,
          dimension: 'COLUMNS',
          startIndex: 0,
          endIndex: 1
        },
        properties: {
          hiddenByUser: true
        },
        fields: 'hiddenByUser'
      }
    });
    
    const columnWidths = TABLE_CONFIG.COLUMN_WIDTHS;
    columnWidths.forEach(col => {
      requests.push({
        updateDimensionProperties: {
          range: {
            sheetId: sheetId,
            dimension: 'COLUMNS',
            startIndex: col.c - 1,
            endIndex: col.c
          },
          properties: {
            pixelSize: col.w
          },
          fields: 'pixelSize'
        }
      });
    });
    
    requests.push({
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: 0,
          endRowIndex: 1,
          startColumnIndex: 0,
          endColumnIndex: numCols
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.26, green: 0.52, blue: 0.96 },
            textFormat: {
              foregroundColor: { red: 1, green: 1, blue: 1 },
              bold: true,
              fontSize: 10
            },
            horizontalAlignment: 'CENTER',
            verticalAlignment: 'MIDDLE',
            wrapStrategy: 'WRAP'
          }
        },
        fields: 'userEnteredFormat'
      }
    });
    
    const rowsByType = {
      app: [],
      week: [],
      sourceApp: [],
      campaign: []
    };
    
    formatData.forEach(item => {
      if (item.type === 'APP') rowsByType.app.push(item.row - 1);
      if (item.type === 'WEEK') rowsByType.week.push(item.row - 1);
      if (item.type === 'SOURCE_APP') rowsByType.sourceApp.push(item.row - 1);
      if (item.type === 'CAMPAIGN') rowsByType.campaign.push(item.row - 1);
    });
    
    const batchFormatRows = (rows, backgroundColor, textColor, bold, fontSize) => {
      const batchSize = 50;
      for (let i = 0; i < rows.length; i += batchSize) {
        const batch = rows.slice(i, i + batchSize);
        batch.forEach(r => {
          requests.push({
            repeatCell: {
              range: {
                sheetId: sheetId,
                startRowIndex: r,
                endRowIndex: r + 1,
                startColumnIndex: 0,
                endColumnIndex: numCols
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: backgroundColor,
                  textFormat: {
                    foregroundColor: textColor,
                    bold: bold,
                    fontSize: fontSize
                  },
                  verticalAlignment: 'MIDDLE'
                }
              },
              fields: 'userEnteredFormat'
            }
          });
        });
      }
    };
    
    batchFormatRows(rowsByType.app, 
      { red: 0.82, green: 0.91, blue: 0.996 }, 
      { red: 0, green: 0, blue: 0 }, 
      true, 10
    );
    
    batchFormatRows(rowsByType.week, 
      { red: 0.91, green: 0.94, blue: 0.996 }, 
      { red: 0, green: 0, blue: 0 }, 
      false, 10
    );
    
    batchFormatRows(rowsByType.sourceApp, 
      { red: 0.94, green: 0.97, blue: 1 }, 
      { red: 0, green: 0, blue: 0 }, 
      false, 9
    );
    
    batchFormatRows(rowsByType.campaign, 
      { red: 1, green: 1, blue: 1 }, 
      { red: 0, green: 0, blue: 0 }, 
      false, 9
    );
    
    if (hyperlinkData.length > 0) {
      hyperlinkData.forEach(link => {
        requests.push({
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
                textFormat: {
                  foregroundColor: { red: 0, green: 0, blue: 0 },
                  underline: false
                }
              }
            },
            fields: 'userEnteredFormat.textFormat'
          }
        });
      });
    }
    
    if (numRows > 1) {
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: 1,
            endRowIndex: numRows,
            startColumnIndex: 4,
            endColumnIndex: 5
          },
          cell: {
            userEnteredFormat: {
              numberFormat: {
                type: 'CURRENCY',
                pattern: '$0.00'
              }
            }
          },
          fields: 'userEnteredFormat.numberFormat'
        }
      });
      
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: 1,
            endRowIndex: numRows,
            startColumnIndex: 7,
            endColumnIndex: 8
          },
          cell: {
            userEnteredFormat: {
              numberFormat: {
                type: 'CURRENCY',
                pattern: '$0.000'
              }
            }
          },
          fields: 'userEnteredFormat.numberFormat'
        }
      });
      
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: 1,
            endRowIndex: numRows,
            startColumnIndex: 8,
            endColumnIndex: 9
          },
          cell: {
            userEnteredFormat: {
              numberFormat: {
                type: 'NUMBER',
                pattern: '0.00'
              }
            }
          },
          fields: 'userEnteredFormat.numberFormat'
        }
      });
      
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: 1,
            endRowIndex: numRows,
            startColumnIndex: 12,
            endColumnIndex: 13
          },
          cell: {
            userEnteredFormat: {
              numberFormat: {
                type: 'CURRENCY',
                pattern: '$0.000'
              }
            }
          },
          fields: 'userEnteredFormat.numberFormat'
        }
      });
      
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: 1,
            endRowIndex: numRows,
            startColumnIndex: 15,
            endColumnIndex: 16
          },
          cell: {
            userEnteredFormat: {
              numberFormat: {
                type: 'CURRENCY',
                pattern: '$0.00'
              }
            }
          },
          fields: 'userEnteredFormat.numberFormat'
        }
      });
      
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: 1,
            endRowIndex: numRows,
            startColumnIndex: numCols - 1,
            endColumnIndex: numCols
          },
          cell: {
            userEnteredFormat: {
              wrapStrategy: 'WRAP',
              horizontalAlignment: 'LEFT'
            }
          },
          fields: 'userEnteredFormat.wrapStrategy,userEnteredFormat.horizontalAlignment'
        }
      });
      
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: 1,
            endRowIndex: numRows,
            startColumnIndex: numCols - 2,
            endColumnIndex: numCols - 1
          },
          cell: {
            userEnteredFormat: {
              wrapStrategy: 'WRAP',
              horizontalAlignment: 'LEFT'
            }
          },
          fields: 'userEnteredFormat.wrapStrategy,userEnteredFormat.horizontalAlignment'
        }
      });
    }
    
    console.log(`Executing ${requests.length} format requests...`);
    
    const batchSize = 100;
    for (let i = 0; i < requests.length; i += batchSize) {
      const batch = requests.slice(i, i + batchSize);
      Sheets.Spreadsheets.batchUpdate({
        requests: batch
      }, spreadsheetId);
      
      if (i + batchSize < requests.length) {
        Utilities.sleep(300);
      }
    }
    
    console.log('Advanced formatting applied successfully');
    
    applyAdvancedConditionalFormatting(sheet, numRows, appData);
    
  } catch (e) {
    console.error('Error in advanced formatting:', e);
    throw e;
  }
}

function applyAdvancedConditionalFormatting(sheet, numRows, appData) {
  try {
    const spreadsheetId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();
    const requests = [];
    
    sheet.clearConditionalFormatRules();
    
    if (numRows > 1) {
      requests.push({
        addConditionalFormatRule: {
          rule: {
            ranges: [{
              sheetId: sheetId,
              startRowIndex: 1,
              endRowIndex: numRows,
              startColumnIndex: 5,
              endColumnIndex: 6
            }],
            booleanRule: {
              condition: {
                type: 'NUMBER_GREATER',
                values: [{ numberValue: 0 }]
              },
              format: {
                backgroundColor: { red: 0.82, green: 0.95, blue: 0.92 },
                textFormat: { foregroundColor: { red: 0.05, green: 0.33, blue: 0.38 } }
              }
            }
          },
          index: 0
        }
      });
      
      requests.push({
        addConditionalFormatRule: {
          rule: {
            ranges: [{
              sheetId: sheetId,
              startRowIndex: 1,
              endRowIndex: numRows,
              startColumnIndex: 5,
              endColumnIndex: 6
            }],
            booleanRule: {
              condition: {
                type: 'NUMBER_LESS',
                values: [{ numberValue: 0 }]
              },
              format: {
                backgroundColor: { red: 0.97, green: 0.84, blue: 0.85 },
                textFormat: { foregroundColor: { red: 0.45, green: 0.11, blue: 0.14 } }
              }
            }
          },
          index: 1
        }
      });
      
      requests.push({
        addConditionalFormatRule: {
          rule: {
            ranges: [{
              sheetId: sheetId,
              startRowIndex: 1,
              endRowIndex: numRows,
              startColumnIndex: 16,
              endColumnIndex: 17
            }],
            booleanRule: {
              condition: {
                type: 'NUMBER_GREATER',
                values: [{ numberValue: 0 }]
              },
              format: {
                backgroundColor: { red: 0.82, green: 0.95, blue: 0.92 },
                textFormat: { foregroundColor: { red: 0.05, green: 0.33, blue: 0.38 } }
              }
            }
          },
          index: 2
        }
      });
      
      requests.push({
        addConditionalFormatRule: {
          rule: {
            ranges: [{
              sheetId: sheetId,
              startRowIndex: 1,
              endRowIndex: numRows,
              startColumnIndex: 16,
              endColumnIndex: 17
            }],
            booleanRule: {
              condition: {
                type: 'NUMBER_LESS',
                values: [{ numberValue: 0 }]
              },
              format: {
                backgroundColor: { red: 0.97, green: 0.84, blue: 0.85 },
                textFormat: { foregroundColor: { red: 0.45, green: 0.11, blue: 0.14 } }
              }
            }
          },
          index: 3
        }
      });
    }
    
    console.log(`Applying ${requests.length} conditional formatting rules...`);
    
    if (requests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({
        requests: requests
      }, spreadsheetId);
    }
    
    console.log('Conditional formatting applied successfully');
    
  } catch (e) {
    console.error('Error in conditional formatting:', e);
    throw e;
  }
}

function createTrickyOptimizedRowGrouping(sheet, tableData, appData) {
  try {
    const spreadsheetId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();
    const requests = [];
    
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

        if (week.sourceApps) {
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
              requests.push({
                addDimensionGroup: {
                  range: {
                    sheetId: sheetId,
                    dimension: 'ROWS',
                    startIndex: sourceAppStartRow,
                    endIndex: sourceAppStartRow + campaignCount
                  }
                }
              });
              
              requests.push({
                updateDimensionGroup: {
                  dimensionGroup: {
                    range: {
                      sheetId: sheetId,
                      dimension: 'ROWS',
                      startIndex: sourceAppStartRow,
                      endIndex: sourceAppStartRow + campaignCount
                    },
                    depth: 1,
                    collapsed: true
                  },
                  fields: 'collapsed'
                }
              });
            }
          });
          
          if (weekContentRows > 0) {
            requests.push({
              addDimensionGroup: {
                range: {
                  sheetId: sheetId,
                  dimension: 'ROWS',
                  startIndex: weekStartRow,
                  endIndex: weekStartRow + weekContentRows
                }
              }
            });
            
            requests.push({
              updateDimensionGroup: {
                dimensionGroup: {
                  range: {
                    sheetId: sheetId,
                    dimension: 'ROWS',
                    startIndex: weekStartRow,
                    endIndex: weekStartRow + weekContentRows
                  },
                  depth: 1,
                  collapsed: true
                },
                fields: 'collapsed'
              }
            });
          }
        }
      });

      const appContentRows = rowPointer - appStartRow - 1;
      if (appContentRows > 0) {
        requests.push({
          addDimensionGroup: {
            range: {
              sheetId: sheetId,
              dimension: 'ROWS',
              startIndex: appStartRow,
              endIndex: appStartRow + appContentRows
            }
          }
        });
        
        requests.push({
          updateDimensionGroup: {
            dimensionGroup: {
              range: {
                sheetId: sheetId,
                dimension: 'ROWS',
                startIndex: appStartRow,
                endIndex: appStartRow + appContentRows
              },
              depth: 1,
              collapsed: true
            },
            fields: 'collapsed'
          }
        });
      }
    });
    
    console.log(`Creating ${requests.length} groups...`);
    
    if (requests.length > 0) {
      const batchSize = 50;
      for (let i = 0; i < requests.length; i += batchSize) {
        const batch = requests.slice(i, i + batchSize);
        Sheets.Spreadsheets.batchUpdate({
          requests: batch
        }, spreadsheetId);
        
        if (i + batchSize < requests.length) {
          Utilities.sleep(200);
        }
      }
    }
    
    console.log('TRICKY optimized row grouping completed');
    
  } catch (e) {
    console.error('Error in createTrickyOptimizedRowGrouping:', e);
    throw e;
  }
}

function createStandardEnhancedPivotTable(appData) {
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
    const emptyRow = new Array(headers.length).fill('');
    emptyRow[0] = 'APP';
    emptyRow[1] = app.appName;
    tableData.push(emptyRow);

    const weekKeys = Object.keys(app.weeks).sort();
    weekKeys.forEach(weekKey => {
      const week = app.weeks[weekKey];
      
      formatData.push({ row: tableData.length + 1, type: 'WEEK' });
      
      if (week.sourceApps) {
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
        
        addStandardSourceAppRows(tableData, week.sourceApps, weekKey, wow, formatData);
        
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
}

function addStandardSourceAppRows(tableData, sourceApps, weekKey, wow, formatData) {
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

function createOverallPivotTable(appData) {
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

  console.log(`Writing ${tableData.length} rows to sheet...`);
  sheet.getRange(1, 1, tableData.length, headers.length).setValues(tableData);
  
  console.log('Applying batch formatting...');
  applyOverallFormattingBatch(sheet, tableData.length, headers.length, formatData, appData);
  
  console.log('Applying overall grouping...');
  createOverallRowGrouping(sheet, tableData, appData);
  
  sheet.setFrozenRows(1);
}

function applyOverallFormattingBatch(sheet, numRows, numCols, formatData, appData) {
  try {
    const spreadsheetId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();
    const requests = [];
    
    requests.push({
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: 0,
          endRowIndex: 1,
          startColumnIndex: 0,
          endColumnIndex: numCols
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.26, green: 0.52, blue: 0.96 },
            textFormat: {
              foregroundColor: { red: 1, green: 1, blue: 1 },
              bold: true,
              fontSize: 10
            },
            horizontalAlignment: 'CENTER',
            verticalAlignment: 'MIDDLE',
            wrapStrategy: 'WRAP'
          }
        },
        fields: 'userEnteredFormat'
      }
    });
    
    const columnWidths = TABLE_CONFIG.COLUMN_WIDTHS;
    columnWidths.forEach(col => {
      requests.push({
        updateDimensionProperties: {
          range: {
            sheetId: sheetId,
            dimension: 'COLUMNS',
            startIndex: col.c - 1,
            endIndex: col.c
          },
          properties: {
            pixelSize: col.w
          },
          fields: 'pixelSize'
        }
      });
    });
    
    const rowsByType = { app: [], week: [] };
    
    formatData.forEach(item => {
      if (item.type === 'APP') rowsByType.app.push(item.row - 1);
      if (item.type === 'WEEK') rowsByType.week.push(item.row - 1);
    });
    
    rowsByType.app.forEach(r => {
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: r,
            endRowIndex: r + 1,
            startColumnIndex: 0,
            endColumnIndex: numCols
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.82, green: 0.91, blue: 0.996 },
              textFormat: {
                foregroundColor: { red: 0, green: 0, blue: 0 },
                bold: true,
                fontSize: 10
              },
              verticalAlignment: 'MIDDLE'
            }
          },
          fields: 'userEnteredFormat'
        }
      });
    });
    
    rowsByType.week.forEach(r => {
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: r,
            endRowIndex: r + 1,
            startColumnIndex: 0,
            endColumnIndex: numCols
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: { red: 0.91, green: 0.94, blue: 0.996 },
              textFormat: {
                foregroundColor: { red: 0, green: 0, blue: 0 },
                bold: false,
                fontSize: 10
              },
              verticalAlignment: 'MIDDLE'
            }
          },
          fields: 'userEnteredFormat'
        }
      });
    });
    
    console.log(`Executing ${requests.length} format requests...`);
    
    const batchSize = 100;
    for (let i = 0; i < requests.length; i += batchSize) {
      const batch = requests.slice(i, i + batchSize);
      Sheets.Spreadsheets.batchUpdate({
        requests: batch
      }, spreadsheetId);
      
      if (i + batchSize < requests.length) {
        Utilities.sleep(300);
      }
    }
    
    console.log('Formatting applied successfully');
    
  } catch (e) {
    console.error('Error in overall formatting:', e);
    throw e;
  }
}

function createOverallRowGrouping(sheet, tableData, appData) {
  try {
    const spreadsheetId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();
    const requests = [];
    
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
        requests.push({
          addDimensionGroup: {
            range: {
              sheetId: sheetId,
              dimension: 'ROWS',
              startIndex: appStartRow,
              endIndex: appStartRow + weekCount
            }
          }
        });
        
        requests.push({
          updateDimensionGroup: {
            dimensionGroup: {
              range: {
                sheetId: sheetId,
                dimension: 'ROWS',
                startIndex: appStartRow,
                endIndex: appStartRow + weekCount
              },
              depth: 1,
              collapsed: true
            },
            fields: 'collapsed'
          }
        });
      }
    });
    
    console.log(`Creating ${requests.length} groups...`);
    
    if (requests.length > 0) {
      const batchSize = 50;
      for (let i = 0; i < requests.length; i += batchSize) {
        const batch = requests.slice(i, i + batchSize);
        Sheets.Spreadsheets.batchUpdate({
          requests: batch
        }, spreadsheetId);
        
        if (i + batchSize < requests.length) {
          Utilities.sleep(200);
        }
      }
    }
    
    console.log('Collapsing groups...');
    console.log(`Collapsing ${requests.length / 2} groups...`);
    
    console.log('Overall row grouping completed');
    
  } catch (e) {
    console.error('Error in createOverallRowGrouping:', e);
    throw e;
  }
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
    weekTotals.avgRoas.toFixed(2), weekTotals.avgIpm.toFixed(1), `${weekTotals.avgRrD1.toFixed(1)}%`, `${weekTotals.avgRrD7.toFixed(1)}%`,
    weekTotals.avgArpu.toFixed(3), `${weekTotals.avgERoas.toFixed(0)}%`, `${weekTotals.avgEROASD730.toFixed(0)}%`,
    weekTotals.totalProfit.toFixed(2), profitWoW, status, ''
  ];
}

function applyEnhancedFormatting(sheet, numRows, numCols, formatData, appData) {
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
    sheet.getRange(2, 5, numRows - 1, 1).setNumberFormat('$0.00');
    sheet.getRange(2, 8, numRows - 1, 1).setNumberFormat('$0.000');
    sheet.getRange(2, 9, numRows - 1, 1).setNumberFormat('0.00');
    sheet.getRange(2, 10, numRows - 1, 1).setNumberFormat('0.0');
    sheet.getRange(2, 13, numRows - 1, 1).setNumberFormat('$0.000');
    sheet.getRange(2, 16, numRows - 1, 1).setNumberFormat('$0.00');
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
    campaign.roas.toFixed(2), campaign.ipm.toFixed(1), `${campaign.rrD1.toFixed(1)}%`, `${campaign.rrD7.toFixed(1)}%`,
    campaign.eArpuForecast.toFixed(3), `${campaign.eRoasForecast.toFixed(0)}%`, `${campaign.eRoasForecastD730.toFixed(0)}%`,
    campaign.eProfitForecast.toFixed(2), profitPct, growthStatus, ''
  ];
}

function createRowGrouping(sheet, tableData, appData) {
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

        if (week.sourceApps) {
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
                sheet.getRange(sourceAppStartRow + 1, 1, campaignCount, 1).shiftRowGroupDepth(1);
                sheet.getRange(sourceAppStartRow + 1, 1, campaignCount, 1).collapseGroups();
              } catch (e) {
                console.log('Error grouping campaigns under source app:', e);
              }
            }
          });
          
          if (weekContentRows > 0) {
            try {
              sheet.getRange(weekStartRow + 1, 1, weekContentRows, 1).shiftRowGroupDepth(1);
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
              sheet.getRange(weekStartRow + 1, 1, campaignCount, 1).shiftRowGroupDepth(1);
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
          sheet.getRange(appStartRow + 1, 1, appContentRows, 1).shiftRowGroupDepth(1);
          sheet.getRange(appStartRow + 1, 1, appContentRows, 1).collapseGroups();
        } catch (e) {
          console.log('Error grouping app content:', e);
        }
      }
    });
    
    console.log('Row grouping completed');
    
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