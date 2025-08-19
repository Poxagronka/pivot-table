/**
 * Row Grouping - Optimized: single API call
 */

function createUnifiedRowGrouping(sheet, tableData, data) {
  try {
    const startTime = Date.now();
    const sheetId = sheet.getSheetId();
    const spreadsheetId = sheet.getParent().getId();
    
    // Собираем ВСЕ requests сразу
    const allRequests = [];
    
    const entities = CURRENT_PROJECT === 'INCENT_TRAFFIC' ?
      Object.keys(data).sort((a, b) => data[a].networkName.localeCompare(data[b].networkName)) :
      Object.keys(data).sort((a, b) => data[a].appName.localeCompare(data[b].appName));
    
    const entityType = CURRENT_PROJECT === 'INCENT_TRAFFIC' ? 'network' : 'app';
    
    // Собираем create и collapse для каждой entity
    entities.forEach(entityKey => {
      const requests = buildGroupRequests(data, entityKey, entityType, sheetId);
      allRequests.push(...requests);
    });
    
    console.log(`Row grouping: ${allRequests.length} total requests`);
    
    // ОДИН batchUpdate для всего
    const BATCH_SIZE = 500;
    for (let i = 0; i < allRequests.length; i += BATCH_SIZE) {
      Sheets.Spreadsheets.batchUpdate({
        requests: allRequests.slice(i, i + BATCH_SIZE)
      }, spreadsheetId);
      
      if (i + BATCH_SIZE < allRequests.length) Utilities.sleep(20);
    }
    
    console.log(`Row grouping completed in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
  } catch (e) {
    console.error('Error in row grouping:', e);
  }
}

function buildGroupRequests(data, entityKey, entityType, sheetId) {
  const requests = [];
  let rowPointer = calculateRowPointer(data, entityKey, entityType);
  const entityStartRow = rowPointer;
  rowPointer++;
  
  const entity = data[entityKey];
  
  if (CURRENT_PROJECT === 'APPLOVIN_TEST' && entity.campaignGroups) {
    // УНИФИЦИРОВАННАЯ логика для APPLOVIN_TEST - точно как в TRICKY
    const campaignKeys = Object.keys(entity.campaignGroups).sort((a, b) => {
      const spendA = Object.values(entity.campaignGroups[a].weeks).reduce((sum, w) => 
        sum + Object.values(w.countries || {}).reduce((s, country) => 
          s + country.campaigns.reduce((cs, c) => cs + c.spend, 0), 0), 0);
      const spendB = Object.values(entity.campaignGroups[b].weeks).reduce((sum, w) => 
        sum + Object.values(w.countries || {}).reduce((s, country) => 
          s + country.campaigns.reduce((cs, c) => cs + c.spend, 0), 0), 0);
      return spendB - spendA;
    });
    
    const collapseData = [];
    let entityTotalRows = 0;
    
    campaignKeys.forEach(campaignKey => {
      const campaignGroup = entity.campaignGroups[campaignKey];
      const campaignStartRow = rowPointer;
      rowPointer++; // Campaign header row
      let campaignContentRows = 0;
      
      const weekKeys = Object.keys(campaignGroup.weeks).sort();
      
      weekKeys.forEach(weekKey => {
        const week = campaignGroup.weeks[weekKey];
        const weekStartRow = rowPointer;
        rowPointer++; // Week header row
        
        // УНИФИЦИРОВАНО: считаем countries как campaigns в TRICKY
        const countries = week.countries || {};
        const countryCount = Object.keys(countries).length;
        rowPointer += countryCount;
        
        // Группа для стран внутри недели (depth 3) - аналог sourceApp в TRICKY
        if (countryCount > 0) {
          requests.push({
            addDimensionGroup: {
              range: { sheetId, dimension: "ROWS", startIndex: weekStartRow, endIndex: weekStartRow + countryCount }
            }
          });
          collapseData.push({ start: weekStartRow, end: weekStartRow + countryCount, depth: 3 });
        }
        
        campaignContentRows += 1 + countryCount; // 1 week header + countries
      });
      
      // Группа для всей кампании (depth 2) - аналог week в TRICKY
      if (campaignContentRows > 0) {
        requests.push({
          addDimensionGroup: {
            range: { sheetId, dimension: "ROWS", startIndex: campaignStartRow, endIndex: campaignStartRow + campaignContentRows }
          }
        });
        collapseData.push({ start: campaignStartRow, end: campaignStartRow + campaignContentRows, depth: 2 });
      }
      
      entityTotalRows += 1 + campaignContentRows; // 1 campaign header + content
    });
    
    // Группа для всего приложения (depth 1)
    if (entityTotalRows > 0) {
      requests.push({
        addDimensionGroup: {
          range: { sheetId, dimension: "ROWS", startIndex: entityStartRow, endIndex: entityStartRow + entityTotalRows }
        }
      });
      collapseData.push({ start: entityStartRow, end: entityStartRow + entityTotalRows, depth: 1 });
    }
    
    // Добавляем collapse в обратном порядке
    collapseData.sort((a, b) => b.depth - a.depth);
    collapseData.forEach(item => {
      requests.push({
        updateDimensionGroup: {
          dimensionGroup: {
            range: { sheetId, dimension: "ROWS", startIndex: item.start, endIndex: item.end },
            depth: item.depth,
            collapsed: true
          },
          fields: "collapsed"
        }
      });
    });
    
    return requests;
  }
  
  // СПЕЦИАЛЬНАЯ ОБРАБОТКА ДЛЯ НОВОЙ СТРУКТУРЫ INCENT_TRAFFIC - КАК В TRICKY
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC' && entity.countries) {
    const collapseData = [];
    let networkTotalRows = 0;
    
    const countryKeys = Object.keys(entity.countries).sort((a, b) =>
      entity.countries[a].countryName.localeCompare(entity.countries[b].countryName)
    );
    
    countryKeys.forEach(countryCode => {
      const country = entity.countries[countryCode];
      const countryStartRow = rowPointer;
      rowPointer++; // Country header row
      let countryContentRows = 0;
      
      const campaignKeys = Object.keys(country.campaigns).sort((a, b) => {
        const spendA = Object.values(country.campaigns[a].weeks).reduce((sum, w) => 
          sum + w.data.reduce((s, d) => s + d.spend, 0), 0);
        const spendB = Object.values(country.campaigns[b].weeks).reduce((sum, w) => 
          sum + w.data.reduce((s, d) => s + d.spend, 0), 0);
        return spendB - spendA;
      });
      
      campaignKeys.forEach(campaignId => {
        const campaign = country.campaigns[campaignId];
        const campaignStartRow = rowPointer;
        rowPointer++; // Campaign header row
        
        const weekCount = Object.keys(campaign.weeks).length;
        rowPointer += weekCount;
        
        // Группа для недель внутри кампании (depth 3) - аналог sourceApp в TRICKY
        if (weekCount > 0) {
          requests.push({
            addDimensionGroup: {
              range: { sheetId, dimension: "ROWS", 
                      startIndex: campaignStartRow, 
                      endIndex: campaignStartRow + weekCount }
            }
          });
          collapseData.push({ start: campaignStartRow, 
                             end: campaignStartRow + weekCount, depth: 3 });
        }
        
        countryContentRows += 1 + weekCount; // 1 campaign header + weeks
      });
      
      // Группа для всех кампаний в стране (depth 2) - аналог week в TRICKY
      if (countryContentRows > 0) {
        requests.push({
          addDimensionGroup: {
            range: { sheetId, dimension: "ROWS", 
                    startIndex: countryStartRow, 
                    endIndex: countryStartRow + countryContentRows }
          }
        });
        collapseData.push({ start: countryStartRow, 
                           end: countryStartRow + countryContentRows, depth: 2 });
      }
      
      networkTotalRows += 1 + countryContentRows; // 1 country header + content
    });
    
    // Группа для всех стран в сети (depth 1) - аналог app в TRICKY
    if (networkTotalRows > 0) {
      requests.push({
        addDimensionGroup: {
          range: { sheetId, dimension: "ROWS", 
                  startIndex: entityStartRow, 
                  endIndex: entityStartRow + networkTotalRows }
        }
      });
      collapseData.push({ start: entityStartRow, 
                         end: entityStartRow + networkTotalRows, depth: 1 });
    }
    
    // Добавляем collapse в обратном порядке (от самого глубокого к самому высокому)
    collapseData.sort((a, b) => b.depth - a.depth);
    collapseData.forEach(item => {
      requests.push({
        updateDimensionGroup: {
          dimensionGroup: {
            range: { sheetId, dimension: "ROWS", 
                    startIndex: item.start, endIndex: item.end },
            depth: item.depth,
            collapsed: true
          },
          fields: "collapsed"
        }
      });
    });
    
    return requests;
  }
  
  // Оригинальный код для остальных проектов
  const weeks = entity.weeks;
  const sortedWeeks = Object.keys(weeks).sort();
  
  const collapseData = [];
  let entityTotalRows = 0;
  
  sortedWeeks.forEach(weekKey => {
    const week = weeks[weekKey];
    const weekStartRow = rowPointer;
    rowPointer++;
    let weekContentRows = 0;
    
    if (entityType === 'network') {
      weekContentRows = Object.keys(week.apps).length;
      rowPointer += weekContentRows;
      
      if (weekContentRows > 0) {
        requests.push({
          addDimensionGroup: {
            range: { sheetId, dimension: "ROWS", startIndex: weekStartRow, endIndex: weekStartRow + weekContentRows }
          }
        });
        collapseData.push({ start: weekStartRow, end: weekStartRow + weekContentRows, depth: 2 });
      }
    } else {
      if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
        const sourceAppKeys = Object.keys(week.sourceApps).sort((a, b) => {
          const spendA = week.sourceApps[a].campaigns.reduce((s, c) => s + c.spend, 0);
          const spendB = week.sourceApps[b].campaigns.reduce((s, c) => s + c.spend, 0);
          return spendB - spendA;
        });
        
        sourceAppKeys.forEach(sourceAppKey => {
          const sourceApp = week.sourceApps[sourceAppKey];
          const sourceAppStartRow = rowPointer;
          rowPointer++;
          const campaignCount = sourceApp.campaigns.length;
          rowPointer += campaignCount;
          
          if (campaignCount > 0) {
            requests.push({
              addDimensionGroup: {
                range: { sheetId, dimension: "ROWS", startIndex: sourceAppStartRow, endIndex: sourceAppStartRow + campaignCount }
              }
            });
            collapseData.push({ start: sourceAppStartRow, end: sourceAppStartRow + campaignCount, depth: 3 });
          }
          weekContentRows += 1 + campaignCount;
        });
      } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
        weekContentRows = Object.keys(week.networks).length;
        rowPointer += weekContentRows;
      } else if (week.campaigns) {
        weekContentRows = week.campaigns.length;
        rowPointer += weekContentRows;
      }
      
      if (weekContentRows > 0) {
        requests.push({
          addDimensionGroup: {
            range: { sheetId, dimension: "ROWS", startIndex: weekStartRow, endIndex: weekStartRow + weekContentRows }
          }
        });
        collapseData.push({ start: weekStartRow, end: weekStartRow + weekContentRows, depth: 2 });
      }
    }
    entityTotalRows += 1 + weekContentRows;
  });
  
  if (entityTotalRows > 0) {
    requests.push({
      addDimensionGroup: {
        range: { sheetId, dimension: "ROWS", startIndex: entityStartRow, endIndex: entityStartRow + entityTotalRows }
      }
    });
    collapseData.push({ start: entityStartRow, end: entityStartRow + entityTotalRows, depth: 1 });
  }
  
  // Добавляем collapse в обратном порядке (от глубоких к поверхностным)
  collapseData.sort((a, b) => b.depth - a.depth);
  collapseData.forEach(item => {
    requests.push({
      updateDimensionGroup: {
        dimensionGroup: {
          range: { sheetId, dimension: "ROWS", startIndex: item.start, endIndex: item.end },
          depth: item.depth,
          collapsed: true
        },
        fields: "collapsed"
      }
    });
  });
  
  return requests;
}

function calculateRowPointer(data, targetEntityKey, entityType) {
  let rowPointer = 2;
  
  const entities = entityType === 'network' ?
    Object.keys(data).sort((a, b) => data[a].networkName.localeCompare(data[b].networkName)) :
    Object.keys(data).sort((a, b) => data[a].appName.localeCompare(data[b].appName));
  
  for (const entityKey of entities) {
    if (entityKey === targetEntityKey) break;
    
    const entity = data[entityKey];
    rowPointer++;
    
    if (CURRENT_PROJECT === 'APPLOVIN_TEST' && entity.campaignGroups) {
      Object.values(entity.campaignGroups).forEach(campaignGroup => {
        rowPointer++; // Campaign row
        Object.values(campaignGroup.weeks).forEach(week => {
          rowPointer++; // Week row
          rowPointer += Object.keys(week.countries || {}).length; // Country rows
        });
      });
    } else if (CURRENT_PROJECT === 'INCENT_TRAFFIC' && entity.countries) {
      // Новая структура для INCENT_TRAFFIC: network -> country -> campaign -> week
      Object.values(entity.countries).forEach(country => {
        rowPointer++; // Country row
        Object.values(country.campaigns).forEach(campaign => {
          rowPointer++; // Campaign row
          rowPointer += Object.keys(campaign.weeks).length; // Week rows
        });
      });
    } else {
      // Оригинальная логика для других проектов
      Object.keys(entity.weeks).forEach(weekKey => {
        const week = entity.weeks[weekKey];
        rowPointer++;
        
        if (entityType === 'network') {
          rowPointer += Object.keys(week.apps).length;
        } else if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
          Object.values(week.sourceApps).forEach(sourceApp => {
            rowPointer++;
            rowPointer += sourceApp.campaigns.length;
          });
        } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
          rowPointer += Object.keys(week.networks).length;
        } else if (week.campaigns) {
          rowPointer += week.campaigns.length;
        }
      });
    }
  }
  
  return rowPointer;
}

// Legacy functions
function processEntityGroups(spreadsheetId, sheetId, data, entityKey, entityType) {
  try {
    const requests = buildGroupRequests(data, entityKey, entityType, sheetId);
    if (requests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
    }
  } catch (e) {
    console.error(`Error processing ${entityType} ${entityKey}:`, e);
  }
}

function buildCreateGroupsForEntity(data, entityKey, entityType, sheetId) {
  return buildGroupRequests(data, entityKey, entityType, sheetId)
    .filter(r => r.addDimensionGroup);
}

function buildCollapseGroupsForEntity(data, entityKey, entityType, sheetId) {
  return buildGroupRequests(data, entityKey, entityType, sheetId)
    .filter(r => r.updateDimensionGroup);
}