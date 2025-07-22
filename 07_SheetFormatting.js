/**
 * Sheet Formatting and Table Creation - ОБНОВЛЕНО: объединенный столбец ROAS D1→D3→D7→D30 + поддержка сеток для OVERALL + кеширование первоначальных eROAS 730d + rich text форматирование
 */

function createEnhancedPivotTable(appData) {
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);
  else sheet.clear();

  // Инициализация кеша первоначальных eROAS
  const initialEROASCache = new InitialEROASCache();
  initialEROASCache.recordInitialValuesFromData(appData);

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
      
      if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
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
        
        const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status, app.appName, initialEROASCache);
        tableData.push(weekRow);
        
        addSourceAppRows(tableData, week.sourceApps, weekKey, wow, formatData, app.appName, week, initialEROASCache);
        
      } else {
        const weekTotals = calculateWeekTotals(week.campaigns);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = wow.appWeekWoW[appWeekKey] || {};
        
        const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
        const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
        const status = weekWoW.growthStatus || '';
        
        const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status, app.appName, initialEROASCache);
        tableData.push(weekRow);
        
        addCampaignRows(tableData, week.campaigns, week, weekKey, wow, formatData, app.appName, initialEROASCache);
      }
    });
  });

  const range = sheet.getRange(1, 1, tableData.length, headers.length);
  range.setValues(tableData);
  
  applyEnhancedFormatting(sheet, tableData.length, headers.length, formatData, appData);
  createRowGrouping(sheet, tableData, appData);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2); // Заморозить первые 2 столбца (Level скрыт, Week Range / Source App видимый)
}

function createOverallPivotTable(appData) {
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);
  else sheet.clear();

  // Проверка на пустые данные
  if (!appData || Object.keys(appData).length === 0) {
    console.log('OVERALL: No data to display');
    const headers = getUnifiedHeaders();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  // Инициализация кеша первоначальных eROAS
  const initialEROASCache = new InitialEROASCache();
  initialEROASCache.recordInitialValuesFromData(appData);

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
      
      // Собираем все кампании из всех сеток для расчета итогов недели
      const allCampaigns = [];
      if (week.networks) {
        Object.values(week.networks).forEach(network => {
          allCampaigns.push(...network.campaigns);
        });
      } else if (week.campaigns) {
        // Обратная совместимость со старой структурой
        allCampaigns.push(...week.campaigns);
      }
      
      const weekTotals = calculateWeekTotals(allCampaigns);
      const appWeekKey = `${app.appName}_${weekKey}`;
      const weekWoW = wow.appWeekWoW[appWeekKey] || {};
      
      const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = weekWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'WEEK' });
      const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status, app.appName, initialEROASCache);
      tableData.push(weekRow);
      
      // Добавляем строки сеток внутри недели
      if (week.networks) {
        const networkKeys = Object.keys(week.networks).sort((a, b) => {
          const totalSpendA = week.networks[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
          const totalSpendB = week.networks[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
          return totalSpendB - totalSpendA;
        });
        
        networkKeys.forEach(networkKey => {
          const network = week.networks[networkKey];
          const networkTotals = calculateWeekTotals(network.campaigns);
          
          // Получаем WoW для сетки
          const networkWoWKey = `${networkKey}_${weekKey}`;
          const networkWoW = wow.campaignWoW[networkWoWKey] || {};
          
          const spendWoW = networkWoW.spendChangePercent !== undefined ? `${networkWoW.spendChangePercent.toFixed(0)}%` : '';
          const profitWoW = networkWoW.eProfitChangePercent !== undefined ? `${networkWoW.eProfitChangePercent.toFixed(0)}%` : '';
          const status = networkWoW.growthStatus || '';
          
          formatData.push({ row: tableData.length + 1, type: 'NETWORK' });
          
          // Создаем строку для сетки
          const weekRange = `${week.weekStart} - ${week.weekEnd}`;
          const networkRow = createNetworkRow(network.networkName, networkTotals, spendWoW, profitWoW, status, app.appName, weekRange, network.networkId, initialEROASCache);
          tableData.push(networkRow);
        });
      }
    });
  });

  const range = sheet.getRange(1, 1, tableData.length, headers.length);
  range.setValues(tableData);
  
  applyEnhancedFormatting(sheet, tableData.length, headers.length, formatData, appData);
  createOverallRowGrouping(sheet, tableData, appData);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

function createIncentTrafficPivotTable(networkData) {
  console.log('=== createIncentTrafficPivotTable START ===');
  console.log('networkData type:', typeof networkData);
  console.log('networkData keys:', Object.keys(networkData || {}));
  
  try {
    const config = getCurrentConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
    let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
    if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);
    else sheet.clear();

    // Инициализация кеша первоначальных eROAS
    const initialEROASCache = new InitialEROASCache();
    initialEROASCache.recordInitialValuesFromData(networkData);

    const wow = calculateIncentTrafficWoWMetrics(networkData);
    const headers = getUnifiedHeaders();
    const tableData = [headers];
    const formatData = [];

    // Сортировка сеток по имени
    const networkKeys = Object.keys(networkData).sort((a, b) => 
      networkData[a].networkName.localeCompare(networkData[b].networkName)
    );
    
    networkKeys.forEach(networkKey => {
      const network = networkData[networkKey];
      
      // Строка сетки
      formatData.push({ row: tableData.length + 1, type: 'NETWORK' });
      const emptyRow = new Array(headers.length).fill('');
      emptyRow[0] = 'NETWORK';  // Это уже правильно
      emptyRow[1] = network.networkName;
      tableData.push(emptyRow);
      
      const weekKeys = Object.keys(network.weeks).sort();
      weekKeys.forEach(weekKey => {
        const week = network.weeks[weekKey];
        
        // Собираем все кампании недели для totals
        const allCampaigns = [];
        Object.values(week.apps).forEach(app => {
          allCampaigns.push(...app.campaigns);
        });
        
        const weekTotals = calculateWeekTotals(allCampaigns);
        const weekWoWKey = `${networkKey}_${weekKey}`;
        const weekWoW = wow.weekWoW[weekWoWKey] || {};
        
        const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
        const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
        const status = weekWoW.growthStatus || '';
        
        formatData.push({ row: tableData.length + 1, type: 'WEEK' });
        const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status, network.networkName, initialEROASCache);
        tableData.push(weekRow);
        
        // Добавляем приложения
        const appKeys = Object.keys(week.apps).sort((a, b) => {
          const totalSpendA = week.apps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
          const totalSpendB = week.apps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
          return totalSpendB - totalSpendA; // Сортировка по убыванию spend
        });
        
        appKeys.forEach(appKey => {
          const app = week.apps[appKey];
          const appTotals = calculateWeekTotals(app.campaigns);
          
          const appWoWKey = `${networkKey}_${weekKey}_${appKey}`;
          const appWoW = wow.appWoW[appWoWKey] || {};
          
          const spendWoW = appWoW.spendChangePercent !== undefined ? `${appWoW.spendChangePercent.toFixed(0)}%` : '';
          const profitWoW = appWoW.eProfitChangePercent !== undefined ? `${appWoW.eProfitChangePercent.toFixed(0)}%` : '';
          const status = appWoW.growthStatus || '';
          
          formatData.push({ row: tableData.length + 1, type: 'APP' });

          const weekRange = `${week.weekStart} - ${week.weekEnd}`;
          
          // Форматируем eROAS 730d с первоначальным значением
          let eROAS730Display = `${appTotals.avgEROASD730.toFixed(0)}%`;
          if (initialEROASCache) {
            eROAS730Display = initialEROASCache.formatEROASWithInitial('APP', network.networkName, weekRange, appTotals.avgEROASD730, app.appId, app.appName);
          }

          const appRow = new Array(headers.length).fill('');
          appRow[0] = 'APP';  // Level
          appRow[1] = app.appName;  // Week Range / Source App
          appRow[2] = '';  // ID
          appRow[3] = '';  // GEO
          appRow[4] = appTotals.totalSpend.toFixed(2);
          appRow[5] = spendWoW;
          appRow[6] = appTotals.totalInstalls;
          appRow[7] = appTotals.avgCpi.toFixed(3);
          appRow[8] = `${appTotals.avgRoasD1.toFixed(0)}% → ${appTotals.avgRoasD3.toFixed(0)}% → ${appTotals.avgRoasD7.toFixed(0)}% → ${appTotals.avgRoasD30.toFixed(0)}%`;
          appRow[9] = appTotals.avgIpm.toFixed(1);
          appRow[10] = `${appTotals.avgRrD1.toFixed(0)}%`;
          appRow[11] = `${appTotals.avgRrD7.toFixed(0)}%`;
          appRow[12] = appTotals.avgArpu.toFixed(3);
          appRow[13] = `${appTotals.avgERoas.toFixed(0)}%`;
          appRow[14] = eROAS730Display;
          appRow[15] = appTotals.totalProfit.toFixed(2);
          appRow[16] = profitWoW;
          appRow[17] = status;
          appRow[18] = '';

          tableData.push(appRow);
        });
      });
    });

    const range = sheet.getRange(1, 1, tableData.length, headers.length);
    range.setValues(tableData);
    
    applyEnhancedFormatting(sheet, tableData.length, headers.length, formatData, networkData);
    createIncentTrafficRowGrouping(sheet, tableData, networkData);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);
    
  } catch (e) {
    console.error('Error in createIncentTrafficPivotTable:', e);
    console.error('Stack trace:', e.stack);
    throw e;
  }
  
  console.log('=== createIncentTrafficPivotTable END ===');
}

function createIncentTrafficRowGrouping(sheet, tableData, networkData) {
  const numCols = getUnifiedHeaders().length;

  try {
    let rowPointer = 2;
    const sortedNetworks = Object.keys(networkData).sort((a, b) => 
      networkData[a].networkName.localeCompare(networkData[b].networkName)
    );

    sortedNetworks.forEach(networkKey => {
      const network = networkData[networkKey];
      const networkStartRow = rowPointer;
      rowPointer++; // Network row

      const sortedWeeks = Object.keys(network.weeks).sort();
      
      sortedWeeks.forEach(weekKey => {
        const week = network.weeks[weekKey];
        const weekStartRow = rowPointer;
        rowPointer++; // Week row
        
        const appCount = Object.keys(week.apps).length;
        rowPointer += appCount;
        
        // Группируем приложения под неделей
        if (appCount > 0) {
          try {
            sheet.getRange(weekStartRow + 1, 1, appCount, numCols).shiftRowGroupDepth(1);
            sheet.getRange(weekStartRow + 1, 1, appCount, 1).collapseGroups();
          } catch (e) {
            console.log('Error grouping apps under week:', e);
          }
        }
      });

      // Группируем все недели под сеткой
      const networkContentRows = rowPointer - networkStartRow - 1;
      if (networkContentRows > 0) {
        try {
          sheet.getRange(networkStartRow + 1, 1, networkContentRows, numCols).shiftRowGroupDepth(1);
          sheet.getRange(networkStartRow + 1, 1, networkContentRows, 1).collapseGroups();
        } catch (e) {
          console.log('Error grouping weeks under network:', e);
        }
      }
    });
    
    console.log('Incent Traffic row grouping completed successfully');
    
  } catch (e) {
    console.error('Error in createIncentTrafficRowGrouping:', e);
  }
}

function createOverallRowGrouping(sheet, tableData, appData) {
  const numCols = getUnifiedHeaders().length;

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
        
        // Считаем количество сеток в неделе
        let networkCount = 0;
        if (week.networks) {
          networkCount = Object.keys(week.networks).length;
        }
        rowPointer += networkCount;
        
        // Группируем сетки под неделей
        if (networkCount > 0) {
          try {
            sheet.getRange(weekStartRow + 1, 1, networkCount, numCols).shiftRowGroupDepth(1);
            sheet.getRange(weekStartRow + 1, 1, networkCount, 1).collapseGroups();
          } catch (e) {
            console.log('Error grouping networks under week:', e);
          }
        }
      });

      // Группируем все недели под приложением
      const appContentRows = rowPointer - appStartRow - 1;
      if (appContentRows > 0) {
        try {
          sheet.getRange(appStartRow + 1, 1, appContentRows, numCols).shiftRowGroupDepth(1);
          sheet.getRange(appStartRow + 1, 1, appContentRows, 1).collapseGroups();
        } catch (e) {
          console.log('Error grouping weeks under app:', e);
        }
      }
    });
    
    console.log('Overall row grouping with networks completed successfully');
    
  } catch (e) {
    console.error('Error in createOverallRowGrouping:', e);
  }
}

function addSourceAppRows(tableData, sourceApps, weekKey, wow, formatData, appName, week, initialEROASCache) {
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
    
    const weekRange = `${week.weekStart} - ${week.weekEnd}`;
    const sourceAppRow = createSourceAppRow(sourceAppDisplayName, sourceAppTotals, spendWoW, profitWoW, status, appName, weekRange, sourceApp.sourceAppId, sourceApp.sourceAppName, initialEROASCache);
    tableData.push(sourceAppRow);
    
    addCampaignRows(tableData, sourceApp.campaigns, { weekStart: weekKey.split('-').join('/'), weekEnd: '' }, weekKey, wow, formatData, appName, initialEROASCache);
  });
}

function createSourceAppRow(sourceAppDisplayName, totals, spendWoW, profitWoW, status, appName = '', weekRange = '', sourceAppId = '', sourceAppName = '', initialEROASCache = null) {
  // ОБНОВЛЕНО: объединенный ROAS столбец с процентами
  const combinedRoas = `${totals.avgRoasD1.toFixed(0)}% → ${totals.avgRoasD3.toFixed(0)}% → ${totals.avgRoasD7.toFixed(0)}% → ${totals.avgRoasD30.toFixed(0)}%`;
  
  // Форматируем eROAS 730d с первоначальным значением
  let eROAS730Display = `${totals.avgEROASD730.toFixed(0)}%`;
  if (initialEROASCache && appName && weekRange) {
    eROAS730Display = initialEROASCache.formatEROASWithInitial('SOURCE_APP', appName, weekRange, totals.avgEROASD730, sourceAppId, sourceAppName);
  }
  
  return [
    'SOURCE_APP', sourceAppDisplayName, '', '',
    totals.totalSpend.toFixed(2), spendWoW, totals.totalInstalls, totals.avgCpi.toFixed(3),
    combinedRoas, totals.avgIpm.toFixed(1), `${totals.avgRrD1.toFixed(0)}%`, `${totals.avgRrD7.toFixed(0)}%`,
    totals.avgArpu.toFixed(3), `${totals.avgERoas.toFixed(0)}%`, eROAS730Display,
    totals.totalProfit.toFixed(2), profitWoW, status, ''
  ];
}

function createNetworkRow(networkName, totals, spendWoW, profitWoW, status, appName = '', weekRange = '', networkId = '', initialEROASCache = null) {
  // Аналогично campaign row, но для сетки
  const combinedRoas = `${totals.avgRoasD1.toFixed(0)}% → ${totals.avgRoasD3.toFixed(0)}% → ${totals.avgRoasD7.toFixed(0)}% → ${totals.avgRoasD30.toFixed(0)}%`;
  
  // Форматируем eROAS 730d с первоначальным значением
  let eROAS730Display = `${totals.avgEROASD730.toFixed(0)}%`;
  if (initialEROASCache && appName && weekRange) {
    eROAS730Display = initialEROASCache.formatEROASWithInitial('NETWORK', appName, weekRange, totals.avgEROASD730, networkId, networkName);
  }
  
  return [
    'NETWORK', networkName, '', '',
    totals.totalSpend.toFixed(2), spendWoW, totals.totalInstalls, totals.avgCpi.toFixed(3),
    combinedRoas, totals.avgIpm.toFixed(1), `${totals.avgRrD1.toFixed(0)}%`, `${totals.avgRrD7.toFixed(0)}%`,
    totals.avgArpu.toFixed(3), `${totals.avgERoas.toFixed(0)}%`, eROAS730Display,
    totals.totalProfit.toFixed(2), profitWoW, status, ''
  ];
}

// ОБНОВЛЕНО: новые заголовки с объединенным ROAS столбцом
function getUnifiedHeaders() {
  return [
    'Level', 'Week Range / Source App', 'ID', 'GEO',
    'Spend', 'Spend WoW %', 'Installs', 'CPI', 'ROAS D1→D3→D7→D30', 'IPM',
    'RR D-1', 'RR D-7', 'eARPU 365d', 'eROAS 365d', 'eROAS 730d', 'eProfit 730d', 'eProfit 730d WoW %', 'Growth Status', 'Comments'
  ];
}

function createWeekRow(week, weekTotals, spendWoW, profitWoW, status, appName = '', initialEROASCache = null) {
  // ОБНОВЛЕНО: объединенный ROAS столбец с процентами
  const combinedRoas = `${weekTotals.avgRoasD1.toFixed(0)}% → ${weekTotals.avgRoasD3.toFixed(0)}% → ${weekTotals.avgRoasD7.toFixed(0)}% → ${weekTotals.avgRoasD30.toFixed(0)}%`;
  
  // Форматируем eROAS 730d с первоначальным значением
  let eROAS730Display = `${weekTotals.avgEROASD730.toFixed(0)}%`;
  if (initialEROASCache && appName) {
    const weekRange = `${week.weekStart} - ${week.weekEnd}`;
    eROAS730Display = initialEROASCache.formatEROASWithInitial('WEEK', appName, weekRange, weekTotals.avgEROASD730);
  }
  
  return [
    'WEEK', `${week.weekStart} - ${week.weekEnd}`, '', '',
    weekTotals.totalSpend.toFixed(2), spendWoW, weekTotals.totalInstalls, weekTotals.avgCpi.toFixed(3),
    combinedRoas, weekTotals.avgIpm.toFixed(1), `${weekTotals.avgRrD1.toFixed(0)}%`, `${weekTotals.avgRrD7.toFixed(0)}%`,
    weekTotals.avgArpu.toFixed(3), `${weekTotals.avgERoas.toFixed(0)}%`, eROAS730Display,
    weekTotals.totalProfit.toFixed(2), profitWoW, status, ''
  ];
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
    
    // ОБНОВЛЕНО: включить wrap для ROAS столбца (столбец 9)
    const roasRange = sheet.getRange(2, 9, numRows - 1, 1);
    roasRange.setWrap(true).setHorizontalAlignment('center');
    
    const commentsRange = sheet.getRange(2, numCols, numRows - 1, 1);
    commentsRange.setWrap(true).setHorizontalAlignment('left');
    
    const growthStatusRange = sheet.getRange(2, numCols - 1, numRows - 1, 1);
    growthStatusRange.setWrap(true).setHorizontalAlignment('left');

    const eroasRange = sheet.getRange(2, 15, numRows - 1, 1);
    eroasRange.setHorizontalAlignment('right');
  }

  const appRows = [], weekRows = [], sourceAppRows = [], campaignRows = [], hyperlinkRows = [], networkRows = [];
  formatData.forEach(item => {
    if (item.type === 'APP') appRows.push(item.row);
    if (item.type === 'WEEK') weekRows.push(item.row);
    if (item.type === 'SOURCE_APP') sourceAppRows.push(item.row);
    if (item.type === 'CAMPAIGN') campaignRows.push(item.row);
    if (item.type === 'NETWORK') networkRows.push(item.row);
    if (item.type === 'HYPERLINK') hyperlinkRows.push(item.row);
  });

  appRows.forEach(r => {
  // Для INCENT_TRAFFIC уровень APP форматируется как кампании (без выделения)
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.CAMPAIGN_ROW.background)
         .setFontWeight('normal')  // Явно указываем обычный шрифт
         .setFontSize(9);
  } else {
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.APP_ROW.background)
         .setFontColor(COLORS.APP_ROW.fontColor)
         .setFontWeight('bold')
         .setFontSize(10);
  }
});

  weekRows.forEach(r =>
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.WEEK_ROW.background)
         .setFontSize(10)
  );

  sourceAppRows.forEach(r =>
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.SOURCE_APP_ROW.background)
         .setFontSize(10)
  );

  campaignRows.forEach(r =>
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.CAMPAIGN_ROW.background)
         .setFontSize(9)
  );

 // Замените код форматирования networkRows в функции applyEnhancedFormatting в файле 07_SheetFormatting.js:

// ОБНОВЛЕНО: форматирование сеток зависит от проекта
networkRows.forEach(r => {
  if (CURRENT_PROJECT === 'OVERALL') {
    // Для OVERALL сетки форматируются как кампании (белый фон, обычный текст, 9 шрифт)
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.CAMPAIGN_ROW.background)
         .setFontWeight('normal')
         .setFontSize(9);
  } else {
    // Для других проектов сетки форматируются как APP_ROW (синий фон, жирный текст, 10 шрифт)
    sheet.getRange(r, 1, 1, numCols)
         .setBackground(COLORS.APP_ROW.background)
         .setFontColor(COLORS.APP_ROW.fontColor)
         .setFontWeight('bold')
         .setFontSize(10);
  }
});

  if (hyperlinkRows.length > 0 && CURRENT_PROJECT === 'TRICKY') {
    hyperlinkRows.forEach(r => {
      const linkCell = sheet.getRange(r, 2);
      linkCell.setFontColor('#000000').setFontLine('none');
    });
  }

  if (numRows > 1) {
    sheet.getRange(2, 5, numRows - 1, 1).setNumberFormat('$0');        // Spend - до целого
    sheet.getRange(2, 8, numRows - 1, 1).setNumberFormat('$0.0');      // CPI - 1 знак после точки
    // ROAS столбец (9) остается текстовым для правильного отображения стрелочек
    sheet.getRange(2, 10, numRows - 1, 1).setNumberFormat('0.0');      // IPM - без изменений
    sheet.getRange(2, 13, numRows - 1, 1).setNumberFormat('$0.0');     // eARPU 365d - 1 знак после точки
    sheet.getRange(2, 16, numRows - 1, 1).setNumberFormat('$0');       // eProfit 730d - до целого
  }

  // Применяем условное форматирование
  applyConditionalFormatting(sheet, numRows, appData);

  // НОВОЕ: Применяем rich text форматирование для eROAS 730d столбца
  applyEROASRichTextFormatting(sheet, numRows);

  
  sheet.hideColumns(1);
  sheet.hideColumns(13, 1); // eARPU 365d
  sheet.hideColumns(14, 1); // eROAS 365d 
}

/**
 * НОВАЯ ФУНКЦИЯ: Применяет rich text форматирование к столбцу eROAS 730d
 * Делает часть до стрелки и саму стрелку серыми
 */
function applyEROASRichTextFormatting(sheet, numRows) {
  if (numRows <= 1) return;
  
  const eroasColumn = 15; // eROAS 730d столбец
  const range = sheet.getRange(2, eroasColumn, numRows - 1, 1);
  const values = range.getValues();
  
  const richTextValues = values.map(row => {
    const cellValue = row[0];
    if (!cellValue || typeof cellValue !== 'string' || !cellValue.includes('→')) {
      // Если нет стрелки, возвращаем как есть
      return SpreadsheetApp.newRichTextValue().setText(cellValue || '').build();
    }
    
    const arrowIndex = cellValue.indexOf('→');
    if (arrowIndex === -1) {
      return SpreadsheetApp.newRichTextValue().setText(cellValue).build();
    }
    
    // Разделяем на часть до стрелки (включая стрелку) и часть после
    const beforeArrow = cellValue.substring(0, arrowIndex + 1); // включаем стрелку
    const afterArrow = cellValue.substring(arrowIndex + 1);
    
    // Создаем rich text value
    const richTextBuilder = SpreadsheetApp.newRichTextValue()
      .setText(cellValue)
      .setTextStyle(0, beforeArrow.length, SpreadsheetApp.newTextStyle()
        .setForegroundColor('#808080')  // Серый цвет
        .build())
      .setTextStyle(beforeArrow.length, cellValue.length, SpreadsheetApp.newTextStyle()
        .setForegroundColor('#000000')  // Черный цвет
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
      
      // ИСПРАВЛЕНО: Формула для извлечения значения после стрелки
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

// ОБНОВЛЕНО: расчет всех ROAS метрик
function calculateWeekTotals(campaigns) {
  const totalSpend = campaigns.reduce((s, c) => s + c.spend, 0);
  const totalInstalls = campaigns.reduce((s, c) => s + c.installs, 0);
  const avgCpi = totalInstalls ? totalSpend / totalInstalls : 0;
  
  // ОБНОВЛЕНО: расчет всех ROAS метрик
  const avgRoasD1 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD1, 0) / campaigns.length : 0;
  const avgRoasD3 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD3, 0) / campaigns.length : 0;
  const avgRoasD7 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD7, 0) / campaigns.length : 0;
  const avgRoasD30 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD30, 0) / campaigns.length : 0;
  
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
    totalSpend, totalInstalls, avgCpi, avgRoasD1, avgRoasD3, avgRoasD7, avgRoasD30, avgIpm, avgRrD1, avgRrD7,
    avgArpu, avgERoas, avgEROASD730, totalProfit
  };
}

function addCampaignRows(tableData, campaigns, week, weekKey, wow, formatData, appName = '', initialEROASCache = null) {
  if (CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'INCENT_TRAFFIC') {
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
    
    const weekRange = `${week.weekStart} - ${week.weekEnd}`;
    const campaignRow = createCampaignRow(campaign, campaignIdValue, spendPct, profitPct, growthStatus, appName, weekRange, initialEROASCache);
    tableData.push(campaignRow);
  });
}

function createCampaignRow(campaign, campaignIdValue, spendPct, profitPct, growthStatus, appName = '', weekRange = '', initialEROASCache = null) {
  // ОБНОВЛЕНО: объединенный ROAS столбец с процентами
  const combinedRoas = `${campaign.roasD1.toFixed(0)}% → ${campaign.roasD3.toFixed(0)}% → ${campaign.roasD7.toFixed(0)}% → ${campaign.roasD30.toFixed(0)}%`;
  
  // Форматируем eROAS 730d с первоначальным значением
  let eROAS730Display = `${campaign.eRoasForecastD730.toFixed(0)}%`;
  if (initialEROASCache && appName && weekRange) {
    eROAS730Display = initialEROASCache.formatEROASWithInitial('CAMPAIGN', appName, weekRange, campaign.eRoasForecastD730, campaign.campaignId, campaign.sourceApp);
  }
  
  return [
    'CAMPAIGN', campaign.sourceApp, campaignIdValue, campaign.geo,
    campaign.spend.toFixed(2), spendPct, campaign.installs, campaign.cpi ? campaign.cpi.toFixed(3) : '0.000',
    combinedRoas, campaign.ipm.toFixed(1), `${campaign.rrD1.toFixed(0)}%`, `${campaign.rrD7.toFixed(0)}%`,
    campaign.eArpuForecast.toFixed(3), `${campaign.eRoasForecast.toFixed(0)}%`, eROAS730Display,
    campaign.eProfitForecast.toFixed(2), profitPct, growthStatus, ''
  ];
}

function createRowGrouping(sheet, tableData, appData) {
  const numCols = getUnifiedHeaders().length;

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
              try {
                sheet.getRange(sourceAppStartRow + 1, 1, campaignCount, numCols).shiftRowGroupDepth(1);
                sheet.getRange(sourceAppStartRow + 1, 1, campaignCount, 1).collapseGroups();
              } catch (e) {
                console.log('Error grouping campaigns under source app:', e);
              }
            }
          });
          
          if (weekContentRows > 0) {
            try {
              sheet.getRange(weekStartRow + 1, 1, weekContentRows, numCols).shiftRowGroupDepth(1);
              sheet.getRange(weekStartRow + 1, 1, weekContentRows, 1).collapseGroups();
            } catch (e) {
              console.log('Error grouping week content:', e);
            }
          }
          
        } else if (CURRENT_PROJECT !== 'OVERALL' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
          const campaignCount = week.campaigns ? week.campaigns.length : 0;
          rowPointer += campaignCount;
          weekContentRows = campaignCount;
          
          if (campaignCount > 0) {
            try {
              sheet.getRange(weekStartRow + 1, 1, campaignCount, numCols).shiftRowGroupDepth(1);
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
          sheet.getRange(appStartRow + 1, 1, appContentRows, numCols).shiftRowGroupDepth(1);
          sheet.getRange(appStartRow + 1, 1, appContentRows, 1).collapseGroups();
        } catch (e) {
          console.log('Error grouping app content:', e);
        }
      }
    });
    
    console.log('Row grouping completed successfully');
    
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
    } else if (projectName === 'INCENT_TRAFFIC') {
      createIncentTrafficPivotTable(appData);
    } else {
      createEnhancedPivotTable(appData);
    }
  } finally {
    setCurrentProject(originalProject);
  }
}