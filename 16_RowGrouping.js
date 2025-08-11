function createUnifiedRowGrouping(sheet, tableData, data) {
  try {
    const startTime = new Date().getTime();
    
    const sheetId = sheet.getSheetId();
    const spreadsheetId = sheet.getParent().getId();
    
    const allCreateRequests = [];
    const allCollapseRequests = [];
    
    if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
      const sortedNetworks = Object.keys(data).sort((a, b) => 
        data[a].networkName.localeCompare(data[b].networkName)
      );
      
      for (const networkKey of sortedNetworks) {
        const createRequests = buildCreateGroupsForEntity(data, networkKey, 'network', sheetId);
        const collapseRequests = buildCollapseGroupsForEntity(data, networkKey, 'network', sheetId);
        
        allCreateRequests.push(...createRequests);
        allCollapseRequests.push(...collapseRequests);
      }
    } else {
      const sortedApps = Object.keys(data).sort((a, b) => 
        data[a].appName.localeCompare(data[b].appName)
      );
      
      for (const appKey of sortedApps) {
        const createRequests = buildCreateGroupsForEntity(data, appKey, 'app', sheetId);
        const collapseRequests = buildCollapseGroupsForEntity(data, appKey, 'app', sheetId);
        
        allCreateRequests.push(...createRequests);
        allCollapseRequests.push(...collapseRequests);
      }
    }
    
    console.log(`Row grouping: ${allCreateRequests.length} create + ${allCollapseRequests.length} collapse requests`);
    
    const BATCH_SIZE = 100;
    
    for (let i = 0; i < allCreateRequests.length; i += BATCH_SIZE) {
      const batch = allCreateRequests.slice(i, i + BATCH_SIZE);
      Sheets.Spreadsheets.batchUpdate({
        requests: batch
      }, spreadsheetId);
    }
    
    for (let i = 0; i < allCollapseRequests.length; i += BATCH_SIZE) {
      const batch = allCollapseRequests.slice(i, i + BATCH_SIZE);
      Sheets.Spreadsheets.batchUpdate({
        requests: batch
      }, spreadsheetId);
    }
    
    console.log(`Row grouping completed in ${(new Date().getTime() - startTime)/1000}s`);
    
  } catch (e) {
    console.error('Error in unified row grouping:', e);
  }
}

function processEntityGroups(spreadsheetId, sheetId, data, entityKey, entityType) {
  try {
    const createRequests = buildCreateGroupsForEntity(data, entityKey, entityType, sheetId);
    
    if (createRequests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({
        requests: createRequests
      }, spreadsheetId);
    }
    
    const collapseRequests = buildCollapseGroupsForEntity(data, entityKey, entityType, sheetId);
    
    if (collapseRequests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({
        requests: collapseRequests
      }, spreadsheetId);
    }
    
  } catch (e) {
    console.error(`Error processing ${entityType} ${entityKey}:`, e);
  }
}

function buildCreateGroupsForEntity(data, entityKey, entityType, sheetId) {
  const groupRequests = [];
  let rowPointer = calculateRowPointer(data, entityKey, entityType);
  
  if (entityType === 'network') {
    const network = data[entityKey];
    const networkStartRow = rowPointer;
    rowPointer++;
    let networkTotalRows = 0;

    const sortedWeeks = Object.keys(network.weeks).sort();
    
    sortedWeeks.forEach(weekKey => {
      const week = network.weeks[weekKey];
      const weekStartRow = rowPointer;
      rowPointer++;
      
      const appCount = Object.keys(week.apps).length;
      rowPointer += appCount;
      networkTotalRows += 1 + appCount;
      
      if (appCount > 0) {
        groupRequests.push({
          addDimensionGroup: {
            range: {
              sheetId: sheetId,
              dimension: "ROWS",
              startIndex: weekStartRow,
              endIndex: weekStartRow + appCount
            }
          }
        });
      }
    });

    if (networkTotalRows > 0) {
      groupRequests.push({
        addDimensionGroup: {
          range: {
            sheetId: sheetId,
            dimension: "ROWS",
            startIndex: networkStartRow,
            endIndex: networkStartRow + networkTotalRows
          }
        }
      });
    }
  } else {
    const app = data[entityKey];
    const appStartRow = rowPointer;
    rowPointer++;
    let appTotalRows = 0;

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
          
          if (campaignCount > 0) {
            groupRequests.push({
              addDimensionGroup: {
                range: {
                  sheetId: sheetId,
                  dimension: "ROWS",
                  startIndex: sourceAppStartRow,
                  endIndex: sourceAppStartRow + campaignCount
                }
              }
            });
          }
          
          weekContentRows += 1 + campaignCount;
        });
      } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
        const networkCount = Object.keys(week.networks).length;
        rowPointer += networkCount;
        weekContentRows = networkCount;
      } else if (CURRENT_PROJECT !== 'OVERALL' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
        const campaignCount = week.campaigns ? week.campaigns.length : 0;
        rowPointer += campaignCount;
        weekContentRows = campaignCount;
      }
      
      if (weekContentRows > 0) {
        groupRequests.push({
          addDimensionGroup: {
            range: {
              sheetId: sheetId,
              dimension: "ROWS",
              startIndex: weekStartRow,
              endIndex: weekStartRow + weekContentRows
            }
          }
        });
        
        appTotalRows += 1 + weekContentRows;
      } else {
        appTotalRows += 1;
      }
    });

    if (appTotalRows > 0) {
      groupRequests.push({
        addDimensionGroup: {
          range: {
            sheetId: sheetId,
            dimension: "ROWS",
            startIndex: appStartRow,
            endIndex: appStartRow + appTotalRows
          }
        }
      });
    }
  }
  
  return groupRequests;
}

function buildCollapseGroupsForEntity(data, entityKey, entityType, sheetId) {
  const collapseRequests = [];
  let rowPointer = calculateRowPointer(data, entityKey, entityType);
  
  if (entityType === 'network') {
    const network = data[entityKey];
    const networkStartRow = rowPointer;
    rowPointer++;
    let networkTotalRows = 0;

    const sortedWeeks = Object.keys(network.weeks).sort();
    
    sortedWeeks.forEach(weekKey => {
      const week = network.weeks[weekKey];
      const weekStartRow = rowPointer;
      rowPointer++;
      
      const appCount = Object.keys(week.apps).length;
      rowPointer += appCount;
      networkTotalRows += 1 + appCount;
      
      if (appCount > 0) {
        collapseRequests.push({
          updateDimensionGroup: {
            dimensionGroup: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: weekStartRow,
                endIndex: weekStartRow + appCount
              },
              depth: 2,
              collapsed: true
            },
            fields: "collapsed"
          }
        });
      }
    });

    if (networkTotalRows > 0) {
      collapseRequests.push({
        updateDimensionGroup: {
          dimensionGroup: {
            range: {
              sheetId: sheetId,
              dimension: "ROWS",
              startIndex: networkStartRow,
              endIndex: networkStartRow + networkTotalRows
            },
            depth: 1,
            collapsed: true
          },
          fields: "collapsed"
        }
      });
    }
  } else {
    const app = data[entityKey];
    const appStartRow = rowPointer;
    rowPointer++;
    let appTotalRows = 0;

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
          
          if (campaignCount > 0) {
            collapseRequests.push({
              updateDimensionGroup: {
                dimensionGroup: {
                  range: {
                    sheetId: sheetId,
                    dimension: "ROWS",
                    startIndex: sourceAppStartRow,
                    endIndex: sourceAppStartRow + campaignCount
                  },
                  depth: 3,
                  collapsed: true
                },
                fields: "collapsed"
              }
            });
          }
          
          weekContentRows += 1 + campaignCount;
        });
      } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
        const networkCount = Object.keys(week.networks).length;
        rowPointer += networkCount;
        weekContentRows = networkCount;
      } else if (CURRENT_PROJECT !== 'OVERALL' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
        const campaignCount = week.campaigns ? week.campaigns.length : 0;
        rowPointer += campaignCount;
        weekContentRows = campaignCount;
      }
      
      if (weekContentRows > 0) {
        collapseRequests.push({
          updateDimensionGroup: {
            dimensionGroup: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: weekStartRow,
                endIndex: weekStartRow + weekContentRows
              },
              depth: 2,
              collapsed: true
            },
            fields: "collapsed"
          }
        });
        
        appTotalRows += 1 + weekContentRows;
      } else {
        appTotalRows += 1;
      }
    });

    if (appTotalRows > 0) {
      collapseRequests.push({
        updateDimensionGroup: {
          dimensionGroup: {
            range: {
              sheetId: sheetId,
              dimension: "ROWS",
              startIndex: appStartRow,
              endIndex: appStartRow + appTotalRows
            },
            depth: 1,
            collapsed: true
          },
          fields: "collapsed"
        }
      });
    }
  }
  
  return collapseRequests;
}

function calculateRowPointer(data, entityKey, entityType) {
  let rowPointer = 2;
  
  if (entityType === 'network') {
    const sortedNetworks = Object.keys(data).sort((a, b) => 
      data[a].networkName.localeCompare(data[b].networkName)
    );
    
    for (const networkKey of sortedNetworks) {
      if (networkKey === entityKey) break;
      
      const network = data[networkKey];
      rowPointer++;
      
      Object.keys(network.weeks).forEach(weekKey => {
        const week = network.weeks[weekKey];
        rowPointer++; 
        rowPointer += Object.keys(week.apps).length;
      });
    }
  } else {
    const sortedApps = Object.keys(data).sort((a, b) => 
      data[a].appName.localeCompare(data[b].appName)
    );
    
    for (const appKey of sortedApps) {
      if (appKey === entityKey) break;
      
      const app = data[appKey];
      rowPointer++;
      
      Object.keys(app.weeks).forEach(weekKey => {
        const week = app.weeks[weekKey];
        rowPointer++;
        
        if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
          Object.keys(week.sourceApps).forEach(sourceAppKey => {
            const sourceApp = week.sourceApps[sourceAppKey];
            rowPointer++;
            rowPointer += sourceApp.campaigns.length;
          });
        } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
          rowPointer += Object.keys(week.networks).length;
        } else if (CURRENT_PROJECT !== 'OVERALL' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
          rowPointer += week.campaigns ? week.campaigns.length : 0;
        }
      });
    }
  }
  
  return rowPointer;
}