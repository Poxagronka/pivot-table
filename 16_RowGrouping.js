function createUnifiedRowGrouping(sheet, tableData, data) {
  try {
    console.log('Starting two-phase unified row grouping...');
    const startTime = new Date().getTime();
    
    const sheetId = sheet.getSheetId();
    const spreadsheetId = sheet.getParent().getId();
    
    if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
      const sortedNetworks = Object.keys(data).sort((a, b) => 
        data[a].networkName.localeCompare(data[b].networkName)
      );
      
      let networkIndex = 0;
      for (const networkKey of sortedNetworks) {
        console.log(`Processing network ${networkIndex + 1}/${sortedNetworks.length}: ${data[networkKey].networkName}`);
        
        processEntityGroups(spreadsheetId, sheetId, data, networkKey, 'network');
        
        networkIndex++;
        if (networkIndex < sortedNetworks.length) {
          console.log('Waiting 1 seconds before next network...');
        }
      }
    } else {
      const sortedApps = Object.keys(data).sort((a, b) => 
        data[a].appName.localeCompare(data[b].appName)
      );
      
      let appIndex = 0;
      for (const appKey of sortedApps) {
        console.log(`Processing app ${appIndex + 1}/${sortedApps.length}: ${data[appKey].appName}`);
        
        processEntityGroups(spreadsheetId, sheetId, data, appKey, 'app');
        
        appIndex++;
        if (appIndex < sortedApps.length) {
          console.log('Waiting 1 seconds before next app...');
        }
      }
    }
    
    const endTime = new Date().getTime();
    console.log(`Two-phase unified row grouping completed in ${(endTime - startTime)/1000}s`);
    
  } catch (e) {
    console.error('Error in unified row grouping:', e);
  }
}

function processEntityGroups(spreadsheetId, sheetId, data, entityKey, entityType) {
  try {
    console.log(`Phase 1: Creating groups for ${entityType} ${entityKey}`);
    const createRequests = buildCreateGroupsForEntity(data, entityKey, entityType, sheetId);
    
    if (createRequests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({
        requests: createRequests
      }, spreadsheetId);
      console.log(`Created ${createRequests.length} groups`);
      
      Utilities.sleep(1000);
    }
    
    console.log(`Phase 2: Collapsing groups for ${entityType} ${entityKey}`);
    const collapseRequests = buildCollapseGroupsForEntity(data, entityKey, entityType, sheetId);
    
    if (collapseRequests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({
        requests: collapseRequests
      }, spreadsheetId);
      console.log(`Collapsed ${collapseRequests.length} groups`);
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
        
        if (campaignCount > 0) {
          groupRequests.push({
            addDimensionGroup: {
              range: {
                sheetId: sheetId,
                dimension: "ROWS",
                startIndex: weekStartRow,
                endIndex: weekStartRow + campaignCount
              }
            }
          });
        }
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
        
        if (campaignCount > 0) {
          collapseRequests.push({
            updateDimensionGroup: {
              dimensionGroup: {
                range: {
                  sheetId: sheetId,
                  dimension: "ROWS",
                  startIndex: weekStartRow,
                  endIndex: weekStartRow + campaignCount
                },
                depth: 2,
                collapsed: true
              },
              fields: "collapsed"
            }
          });
        }
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