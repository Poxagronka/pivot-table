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
    const BATCH_SIZE = 5000;
    for (let i = 0; i < allRequests.length; i += BATCH_SIZE) {
      Sheets.Spreadsheets.batchUpdate({
        requests: allRequests.slice(i, i + BATCH_SIZE)
      }, spreadsheetId);
      
      if (i + BATCH_SIZE < allRequests.length) Utilities.sleep(50);
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
  
  let entityTotalRows = 0;
  const entity = data[entityKey];
  const weeks = entity.weeks;
  const sortedWeeks = Object.keys(weeks).sort();
  
  // Сначала собираем все позиции и create requests
  const collapseData = [];
  
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