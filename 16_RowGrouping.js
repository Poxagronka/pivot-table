function createUnifiedRowGrouping(sheet, tableData, data) {
  try {
    console.log('Starting optimized unified row grouping...');
    const startTime = new Date().getTime();
    
    const sheetId = sheet.getSheetId();
    const spreadsheetId = sheet.getParent().getId();
    
    if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
      const sortedNetworks = Object.keys(data).sort((a, b) => 
        data[a].networkName.localeCompare(data[b].networkName)
      );
      
      console.log(`Processing ${sortedNetworks.length} networks...`);
      
      for (let networkIndex = 0; networkIndex < sortedNetworks.length; networkIndex++) {
        const networkKey = sortedNetworks[networkIndex];
        console.log(`Processing network ${networkIndex + 1}/${sortedNetworks.length}: ${data[networkKey].networkName}`);
        
        processEntityGroupsOptimized(spreadsheetId, sheetId, data, networkKey, 'network');
        
        if (networkIndex < sortedNetworks.length - 1 && networkIndex % 5 === 4) {
          console.log('Pausing after batch of 5 networks...');
          Utilities.sleep(2000);
        }
      }
    } else {
      const sortedApps = Object.keys(data).sort((a, b) => 
        data[a].appName.localeCompare(data[b].appName)
      );
      
      console.log(`Processing ${sortedApps.length} apps...`);
      
      for (let appIndex = 0; appIndex < sortedApps.length; appIndex++) {
        const appKey = sortedApps[appIndex];
        console.log(`Processing app ${appIndex + 1}/${sortedApps.length}: ${data[appKey].appName}`);
        
        processEntityGroupsOptimized(spreadsheetId, sheetId, data, appKey, 'app');
        
        if (appIndex < sortedApps.length - 1 && appIndex % 3 === 2) {
          console.log('Pausing after batch of 3 apps...');
          Utilities.sleep(2000);
        }
      }
    }
    
    const endTime = new Date().getTime();
    console.log(`Optimized unified row grouping completed in ${(endTime - startTime)/1000}s`);
    
  } catch (e) {
    console.error('Error in unified row grouping:', e);
  }
}

function processEntityGroupsOptimized(spreadsheetId, sheetId, data, entityKey, entityType) {
  try {
    console.log(`Creating groups for ${entityType} ${entityKey}...`);
    const createRequests = buildCreateGroupsForEntity(data, entityKey, entityType, sheetId);
    
    if (createRequests.length > 0) {
      executeBatchedGroupRequests(spreadsheetId, createRequests, 'CREATE');
      console.log(`Created ${createRequests.length} groups`);
      
      Utilities.sleep(1000);
      SpreadsheetApp.flush();
    }
    
    console.log(`Collapsing groups for ${entityType} ${entityKey}...`);
    const collapseRequests = buildCollapseGroupsForEntity(data, entityKey, entityType, sheetId);
    
    if (collapseRequests.length > 0) {
      executeBatchedGroupRequests(spreadsheetId, collapseRequests, 'COLLAPSE');
      console.log(`Collapsed ${collapseRequests.length} groups`);
    }
    
  } catch (e) {
    console.error(`Error processing ${entityType} ${entityKey}:`, e);
  }
}

function executeBatchedGroupRequests(spreadsheetId, requests, operation) {
  const BATCH_SIZE = 50;
  
  if (requests.length <= BATCH_SIZE) {
    Sheets.Spreadsheets.batchUpdate({
      requests: requests
    }, spreadsheetId);
    return;
  }
  
  let batchCount = 0;
  for (let i = 0; i < requests.length; i += BATCH_SIZE) {
    const batchRequests = requests.slice(i, i + BATCH_SIZE);
    batchCount++;
    
    console.log(`${operation} batch ${batchCount}: ${batchRequests.length} operations`);
    
    try {
      Sheets.Spreadsheets.batchUpdate({
        requests: batchRequests
      }, spreadsheetId);
      
      if (i + BATCH_SIZE < requests.length) {
        console.log('Pausing between group batches...');
        Utilities.sleep(1500);
      }
    } catch (e) {
      console.error(`Error in ${operation} batch ${batchCount}:`, e);
      if (e.toString().includes('quota')) {
        console.log('Quota exceeded, waiting longer...');
        Utilities.sleep(5000);
      }
    }
  }
  
  console.log(`Completed ${operation} operation: ${requests.length} requests in ${batchCount} batches`);
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
          weekContentRows += 1 + campaignCount;
          
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
        });
        
      } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
        const networkCount = Object.keys(week.networks).length;
        rowPointer += networkCount;
        weekContentRows = networkCount;
        
        if (networkCount > 0) {
          groupRequests.push({
            addDimensionGroup: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: weekStartRow,
                endIndex: weekStartRow + networkCount
              }
            }
          });
        }
        
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
          weekContentRows += 1 + campaignCount;
          
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
        });
        
      } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
        const networkCount = Object.keys(week.networks).length;
        rowPointer += networkCount;
        weekContentRows = networkCount;
        
        if (networkCount > 0) {
          collapseRequests.push({
            updateDimensionGroup: {
              dimensionGroup: {
                range: {
                  sheetId: sheetId,
                  dimension: "ROWS",
                  startIndex: weekStartRow,
                  endIndex: weekStartRow + networkCount
                },
                depth: 2,
                collapsed: true
              },
              fields: "collapsed"
            }
          });
        }
        
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