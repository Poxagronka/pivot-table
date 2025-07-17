function createEnhancedPivotTable(appData) {
  console.log('=== СОЗДАНИЕ ENHANCED PIVOT TABLE ===');
  console.log(`Приложений получено: ${Object.keys(appData).length}`);
  
  if (CURRENT_PROJECT === 'TRICKY') {
    createTrickyOptimizedPivotTable(appData);
  } else {
    createStandardEnhancedPivotTable(appData);
  }
}

function createTrickyOptimizedPivotTable(appData) {
  console.log('Создание TRICKY оптимизированной таблицы...');
  const config = getCurrentConfig();
  const wow = calculateWoWMetrics(appData);
  const headers = getUnifiedHeaders();
  
  const tableData = [headers];
  const formatData = [];
  const hyperlinkData = [];
  const groupingData = [];
  
  let currentRow = 1;
  const appKeys = Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName));
  console.log(`Обработка ${appKeys.length} приложений TRICKY...`);
  
  appKeys.forEach((appKey, appIndex) => {
    const app = appData[appKey];
    const appStartRow = currentRow + 1;
    console.log(`  [${appIndex + 1}/${appKeys.length}] ${app.appName}`);
    
    formatData.push({ row: currentRow + 1, type: 'APP' });
    const emptyRow = new Array(headers.length).fill('');
    emptyRow[0] = 'APP';
    emptyRow[1] = app.appName;
    tableData.push(emptyRow);
    currentRow++;

    const weekKeys = Object.keys(app.weeks).sort();
    let appContentRows = 0;
    console.log(`    Недель: ${weekKeys.length}`);

    weekKeys.forEach((weekKey, weekIndex) => {
      const week = app.weeks[weekKey];
      const weekStartRow = currentRow + 1;
      console.log(`    [${weekIndex + 1}/${weekKeys.length}] Неделя ${weekKey}`);
      
      formatData.push({ row: currentRow + 1, type: 'WEEK' });
      
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
      currentRow++;
      
      const weekContentRows = addTrickyOptimizedSourceAppRows(tableData, week.sourceApps, weekKey, wow, formatData, hyperlinkData, currentRow);
      console.log(`      Добавлено source app строк: ${weekContentRows}`);
      currentRow += weekContentRows;
      appContentRows += 1 + weekContentRows;
      
      if (weekContentRows > 0) {
        groupingData.push({
          type: 'week',
          startRow: weekStartRow,
          rowCount: weekContentRows,
          depth: 1
        });
      }
    });
    
    if (appContentRows > 0) {
      groupingData.push({
        type: 'app',
        startRow: appStartRow,
        rowCount: appContentRows,
        depth: 1
      });
      console.log(`  Группа для ${app.appName}: строки ${appStartRow}-${appStartRow + appContentRows - 1}`);
    }
  });

  console.log(`TRICKY подготовка завершена: ${tableData.length} строк, ${groupingData.length} групп, ${hyperlinkData.length} гиперссылок`);
  writeTableWithTrickyOptimization(config, tableData, formatData, hyperlinkData, groupingData, headers.length, appData);
}

function createStandardEnhancedPivotTable(appData) {
  console.log('Создание стандартной Enhanced таблицы...');
  const config = getCurrentConfig();
  const wow = calculateWoWMetrics(appData);
  const headers = getUnifiedHeaders();
  
  const tableData = [headers];
  const formatData = [];
  const groupingData = [];
  
  let currentRow = 1;
  const appKeys = Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName));
  console.log(`Обработка ${appKeys.length} приложений...`);
  
  appKeys.forEach((appKey, appIndex) => {
    const app = appData[appKey];
    const appStartRow = currentRow + 1;
    console.log(`  [${appIndex + 1}/${appKeys.length}] ${app.appName}`);
    
    formatData.push({ row: currentRow + 1, type: 'APP' });
    const emptyRow = new Array(headers.length).fill('');
    emptyRow[0] = 'APP';
    emptyRow[1] = app.appName;
    tableData.push(emptyRow);
    currentRow++;

    const weekKeys = Object.keys(app.weeks).sort();
    let appContentRows = 0;
    console.log(`    Недель: ${weekKeys.length}`);

    weekKeys.forEach((weekKey, weekIndex) => {
      const week = app.weeks[weekKey];
      const weekStartRow = currentRow + 1;
      console.log(`    [${weekIndex + 1}/${weekKeys.length}] Неделя ${weekKey}`);
      
      formatData.push({ row: currentRow + 1, type: 'WEEK' });
      
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
        currentRow++;
        
        const weekContentRows = addStandardSourceAppRows(tableData, week.sourceApps, weekKey, wow, formatData, currentRow);
        console.log(`      Добавлено source app строк: ${weekContentRows}`);
        currentRow += weekContentRows;
        appContentRows += 1 + weekContentRows;
        
        if (weekContentRows > 0) {
          groupingData.push({
            type: 'week',
            startRow: weekStartRow,
            rowCount: weekContentRows,
            depth: 1
          });
        }
        
      } else {
        const weekTotals = calculateWeekTotals(week.campaigns || []);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = wow.appWeekWoW[appWeekKey] || {};
        
        const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
        const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
        const status = weekWoW.growthStatus || '';
        
        const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status);
        tableData.push(weekRow);
        currentRow++;
        
        const campaignCount = addCampaignRows(tableData, week.campaigns || [], week, weekKey, wow, formatData, currentRow);
        console.log(`      Добавлено кампаний: ${campaignCount}`);
        currentRow += campaignCount;
        appContentRows += 1 + campaignCount;
        
        if (campaignCount > 0) {
          groupingData.push({
            type: 'week',
            startRow: weekStartRow,
            rowCount: campaignCount,
            depth: 1
          });
        }
      }
    });
    
    if (appContentRows > 0) {
      groupingData.push({
        type: 'app',
        startRow: appStartRow,
        rowCount: appContentRows,
        depth: 1
      });
      console.log(`  Группа для ${app.appName}: строки ${appStartRow}-${appStartRow + appContentRows - 1}`);
    }
  });

  console.log(`Стандартная подготовка завершена: ${tableData.length} строк, ${groupingData.length} групп`);
  writeTableWithCompleteFlow(config, tableData, formatData, groupingData, headers.length, appData);
}

function createOverallPivotTable(appData) {
  console.log('=== СОЗДАНИЕ OVERALL PIVOT TABLE ===');
  console.log(`Приложений получено: ${Object.keys(appData).length}`);
  
  const config = getCurrentConfig();
  const wow = calculateWoWMetrics(appData);
  const headers = getUnifiedHeaders();
  
  const tableData = [headers];
  const formatData = [];
  const groupingData = [];
  
  let currentRow = 1;
  const appKeys = Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName));
  console.log(`Обработка ${appKeys.length} приложений OVERALL...`);

  appKeys.forEach((appKey, appIndex) => {
    const app = appData[appKey];
    const appStartRow = currentRow + 1;
    console.log(`  [${appIndex + 1}/${appKeys.length}] ${app.appName}`);
    
    formatData.push({ row: currentRow + 1, type: 'APP' });
    const emptyRow = new Array(headers.length).fill('');
    emptyRow[0] = 'APP';
    emptyRow[1] = app.appName;
    tableData.push(emptyRow);
    currentRow++;

    const weekKeys = Object.keys(app.weeks).sort();
    const weekCount = weekKeys.length;
    console.log(`    Недель: ${weekCount}`);

    weekKeys.forEach((weekKey, weekIndex) => {
      const week = app.weeks[weekKey];
      console.log(`    [${weekIndex + 1}/${weekCount}] Неделя ${weekKey}`);
      
      const weekTotals = calculateWeekTotals(week.campaigns || []);
      const appWeekKey = `${app.appName}_${weekKey}`;
      const weekWoW = wow.appWeekWoW[appWeekKey] || {};
      
      const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = weekWoW.growthStatus || '';
      
      formatData.push({ row: currentRow + 1, type: 'WEEK' });
      const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status);
      tableData.push(weekRow);
      currentRow++;
    });
    
    if (weekCount > 0) {
      groupingData.push({
        type: 'app',
        startRow: appStartRow,
        rowCount: weekCount,
        depth: 1
      });
      console.log(`  Группа для ${app.appName}: строки ${appStartRow}-${appStartRow + weekCount - 1}`);
    }
  });

  console.log(`OVERALL подготовка завершена: ${tableData.length} строк, ${groupingData.length} групп`);
  writeTableWithCompleteFlow(config, tableData, formatData, groupingData, headers.length, appData);
}

function writeTableWithTrickyOptimization(config, tableData, formatData, hyperlinkData, groupingData, numCols, appData) {
  console.log('=== TRICKY ОПТИМИЗИРОВАННАЯ ЗАПИСЬ ===');
  const numRows = tableData.length;
  const sheetName = config.SHEET_NAME;
  console.log(`TRICKY таблица: ${numRows} строк x ${numCols} колонок, ${hyperlinkData.length} гиперссылок`);
  
  let sheetId;
  
  try {
    console.log('ЭТАП 1: Подготовка листа...');
    const existingSheet = getSheetByName(config.SHEET_ID, sheetName);
    if (existingSheet) {
      console.log('Сохранение комментариев...');
      try {
        new CommentCache().syncCommentsFromSheet();
        console.log('✅ Комментарии сохранены');
      } catch (e) {
        console.log('⚠️ Ошибка сохранения комментариев:', e.toString());
      }
    }
    
    sheetId = ensureSheetExists(config.SHEET_ID, sheetName, true);
    console.log(`✅ TRICKY лист готов с ID: ${sheetId}`);
    
    console.log('ЭТАП 2: Запись данных...');
    Sheets.Spreadsheets.Values.update({
      majorDimension: 'ROWS',
      values: tableData
    }, config.SHEET_ID, `${sheetName}!A1:${getColumnLetter(numCols)}${numRows}`, {
      valueInputOption: 'USER_ENTERED'
    });
    console.log(`✅ TRICKY данные записаны: ${numRows} строк`);
    
    console.log('ЭТАП 3: TRICKY оптимизированное форматирование...');
    applyTrickyOptimizedFormatting(config.SHEET_ID, sheetId, sheetName, numRows, numCols, formatData, hyperlinkData, appData);
    console.log('✅ TRICKY форматирование применено');
    
    console.log('ЭТАП 4: TRICKY группировка...');
    applyBatchGrouping(config.SHEET_ID, sheetId, groupingData);
    console.log('✅ TRICKY группировка создана');
    
    console.log('ЭТАП 5: Финальные настройки...');
    const finalRequests = [
      {
        updateSheetProperties: {
          properties: {
            sheetId: sheetId,
            gridProperties: {
              frozenRowCount: 1
            }
          },
          fields: 'gridProperties.frozenRowCount'
        }
      },
      {
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
      }
    ];
    
    Sheets.Spreadsheets.batchUpdate({
      requests: finalRequests
    }, config.SHEET_ID);
    console.log('✅ TRICKY финальные настройки применены');
    
    console.log('ЭТАП 6: Восстановление комментариев...');
    try {
      const cache = new CommentCache();
      cache.applyCommentsToSheet();
      console.log('✅ TRICKY комментарии восстановлены');
    } catch (e) {
      console.log('⚠️ Ошибка восстановления комментариев:', e.toString());
    }
    
    console.log('=== TRICKY ТАБЛИЦА СОЗДАНА ===');
    
  } catch (e) {
    console.error('❌ КРИТИЧЕСКАЯ ОШИБКА в TRICKY writeTable:', e.toString());
    throw e;
  }
}

function writeTableWithCompleteFlow(config, tableData, formatData, groupingData, numCols, appData) {
  console.log('=== СТАНДАРТНЫЙ ПОТОК ЗАПИСИ ТАБЛИЦЫ ===');
  const numRows = tableData.length;
  const sheetName = config.SHEET_NAME;
  console.log(`Таблица: ${numRows} строк x ${numCols} колонок, лист: ${sheetName}`);
  
  let sheetId;
  
  try {
    console.log('ЭТАП 1: Проверка существования листа...');
    const existingSheet = getSheetByName(config.SHEET_ID, sheetName);
    if (existingSheet) {
      console.log(`Существующий лист найден с ID: ${existingSheet.properties.sheetId}`);
      console.log('Сохранение комментариев перед пересозданием...');
      try {
        new CommentCache().syncCommentsFromSheet();
        console.log('✅ Комментарии сохранены');
      } catch (e) {
        console.log('⚠️ Ошибка сохранения комментариев:', e.toString());
      }
    }
    
    console.log('ЭТАП 2: Создание/пересоздание листа...');
    sheetId = ensureSheetExists(config.SHEET_ID, sheetName, true);
    console.log(`✅ Лист готов с ID: ${sheetId}`);
    
    console.log('ЭТАП 3: Запись данных...');
    Sheets.Spreadsheets.Values.update({
      majorDimension: 'ROWS',
      values: tableData
    }, config.SHEET_ID, `${sheetName}!A1:${getColumnLetter(numCols)}${numRows}`, {
      valueInputOption: 'USER_ENTERED'
    });
    console.log(`✅ Данные записаны: ${numRows} строк`);
    
    console.log('ЭТАП 4: Применение форматирования...');
    applyCompleteFormatting(config.SHEET_ID, sheetId, sheetName, numRows, numCols, formatData, appData);
    console.log('✅ Форматирование применено');
    
    console.log('ЭТАП 5: Создание группировки...');
    applyBatchGrouping(config.SHEET_ID, sheetId, groupingData);
    console.log('✅ Группировка создана');
    
    console.log('ЭТАП 6: Финальные настройки...');
    const finalRequests = [
      {
        updateSheetProperties: {
          properties: {
            sheetId: sheetId,
            gridProperties: {
              frozenRowCount: 1
            }
          },
          fields: 'gridProperties.frozenRowCount'
        }
      },
      {
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
      }
    ];
    
    Sheets.Spreadsheets.batchUpdate({
      requests: finalRequests
    }, config.SHEET_ID);
    console.log('✅ Финальные настройки применены');
    
    console.log('ЭТАП 7: Восстановление комментариев...');
    try {
      const cache = new CommentCache();
      cache.applyCommentsToSheet();
      console.log('✅ Комментарии восстановлены');
    } catch (e) {
      console.log('⚠️ Ошибка восстановления комментариев:', e.toString());
    }
    
    console.log('=== ТАБЛИЦА СОЗДАНА УСПЕШНО ===');
    
  } catch (e) {
    console.error('❌ КРИТИЧЕСКАЯ ОШИБКА в writeTableWithCompleteFlow:', e.toString());
    console.error('Stack trace:', e.stack || 'Нет stack trace');
    throw e;
  }
}

function addTrickyOptimizedSourceAppRows(tableData, sourceApps, weekKey, wow, formatData, hyperlinkData, currentRow) {
  if (!sourceApps) return 0;
  
  let addedRows = 0;
  const cache = initTrickyOptimizedCache();
  
  const sourceAppKeys = Object.keys(sourceApps).sort((a, b) => {
    const totalSpendA = sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
    const totalSpendB = sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
    return totalSpendB - totalSpendA;
  });
  
  sourceAppKeys.forEach((sourceAppKey, index) => {
    const sourceApp = sourceApps[sourceAppKey];
    console.log(`        [${index + 1}/${sourceAppKeys.length}] ${sourceApp.sourceAppName}: ${sourceApp.campaigns.length} кампаний`);
    
    const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
    
    const sourceAppWoWKey = `${sourceApp.sourceAppId}_${weekKey}`;
    const sourceAppWoW = wow.sourceAppWoW[sourceAppWoWKey] || {};
    
    const spendWoW = sourceAppWoW.spendChangePercent !== undefined ? `${sourceAppWoW.spendChangePercent.toFixed(0)}%` : '';
    const profitWoW = sourceAppWoW.eProfitChangePercent !== undefined ? `${sourceAppWoW.eProfitChangePercent.toFixed(0)}%` : '';
    const status = sourceAppWoW.growthStatus || '';
    
    formatData.push({ row: currentRow + addedRows + 1, type: 'SOURCE_APP' });
    
    let sourceAppDisplayName = sourceApp.sourceAppName;
    const appInfo = cache?.appsDbCache[sourceApp.sourceAppId];
    if (appInfo && appInfo.linkApp) {
      sourceAppDisplayName = `=HYPERLINK("${appInfo.linkApp}", "${sourceApp.sourceAppName}")`;
      hyperlinkData.push({ row: currentRow + addedRows + 1, col: 2 });
      console.log(`          Гиперссылка добавлена для ${sourceApp.sourceAppName}`);
    }
    
    const sourceAppRow = createSourceAppRow(sourceAppDisplayName, sourceAppTotals, spendWoW, profitWoW, status);
    tableData.push(sourceAppRow);
    addedRows++;
    
    const campaignRowsAdded = addTrickyOptimizedCampaignRows(tableData, sourceApp.campaigns, weekKey, wow, formatData, currentRow + addedRows);
    addedRows += campaignRowsAdded;
    console.log(`          Кампаний добавлено: ${campaignRowsAdded}`);
  });
  
  return addedRows;
}

function addTrickyOptimizedCampaignRows(tableData, campaigns, weekKey, wow, formatData, currentRow) {
  let addedRows = 0;
  
  campaigns.sort((a, b) => b.spend - a.spend).forEach(campaign => {
    const campaignIdValue = `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`;
    
    const key = `${campaign.campaignId}_${weekKey}`;
    const campaignWoW = wow.campaignWoW[key] || {};
    
    const spendPct = campaignWoW.spendChangePercent !== undefined ? `${campaignWoW.spendChangePercent.toFixed(0)}%` : '';
    const profitPct = campaignWoW.eProfitChangePercent !== undefined ? `${campaignWoW.eProfitChangePercent.toFixed(0)}%` : '';
    const growthStatus = campaignWoW.growthStatus || '';
    
    formatData.push({ row: currentRow + addedRows + 1, type: 'CAMPAIGN' });
    
    const campaignRow = createCampaignRow(campaign, campaignIdValue, spendPct, profitPct, growthStatus);
    tableData.push(campaignRow);
    addedRows++;
  });
  
  return addedRows;
}

function addStandardSourceAppRows(tableData, sourceApps, weekKey, wow, formatData, currentRow) {
  let addedRows = 0;
  
  const sourceAppKeys = Object.keys(sourceApps).sort((a, b) => {
    const totalSpendA = sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
    const totalSpendB = sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
    return totalSpendB - totalSpendA;
  });
  
  sourceAppKeys.forEach((sourceAppKey, index) => {
    const sourceApp = sourceApps[sourceAppKey];
    console.log(`        [${index + 1}/${sourceAppKeys.length}] ${sourceApp.sourceAppName}: ${sourceApp.campaigns.length} кампаний`);
    
    const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
    
    const sourceAppWoWKey = `${sourceApp.sourceAppId}_${weekKey}`;
    const sourceAppWoW = wow.sourceAppWoW[sourceAppWoWKey] || {};
    
    const spendWoW = sourceAppWoW.spendChangePercent !== undefined ? `${sourceAppWoW.spendChangePercent.toFixed(0)}%` : '';
    const profitWoW = sourceAppWoW.eProfitChangePercent !== undefined ? `${sourceAppWoW.eProfitChangePercent.toFixed(0)}%` : '';
    const status = sourceAppWoW.growthStatus || '';
    
    formatData.push({ row: currentRow + addedRows + 1, type: 'SOURCE_APP' });
    
    let sourceAppDisplayName = sourceApp.sourceAppName;
    if (CURRENT_PROJECT === 'TRICKY') {
      try {
        const appsDb = new AppsDatabase('TRICKY');
        const cache = appsDb.loadFromCache();
        const appInfo = cache[sourceApp.sourceAppId];
        if (appInfo && appInfo.linkApp) {
          sourceAppDisplayName = `=HYPERLINK("${appInfo.linkApp}", "${sourceApp.sourceAppName}")`;
          console.log(`          Гиперссылка добавлена для ${sourceApp.sourceAppName}`);
        }
      } catch (e) {
        console.log(`          ⚠️ Ошибка получения ссылки для ${sourceApp.sourceAppName}`);
      }
    }
    
    const sourceAppRow = createSourceAppRow(sourceAppDisplayName, sourceAppTotals, spendWoW, profitWoW, status);
    tableData.push(sourceAppRow);
    addedRows++;
    
    const campaignRowsAdded = addCampaignRows(tableData, sourceApp.campaigns, { weekStart: weekKey.split('-').join('/'), weekEnd: '' }, weekKey, wow, formatData, currentRow + addedRows);
    addedRows += campaignRowsAdded;
    console.log(`          Кампаний добавлено: ${campaignRowsAdded}`);
  });
  
  return addedRows;
}

function addCampaignRows(tableData, campaigns, week, weekKey, wow, formatData, currentRow) {
  if (CURRENT_PROJECT === 'OVERALL') {
    return 0;
  }
  
  if (!campaigns || campaigns.length === 0) {
    return 0;
  }
  
  let addedRows = 0;
  
  campaigns.sort((a, b) => b.spend - a.spend).forEach((campaign, index) => {
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
    
    formatData.push({ row: currentRow + addedRows + 1, type: 'CAMPAIGN' });
    
    const campaignRow = createCampaignRow(campaign, campaignIdValue, spendPct, profitPct, growthStatus);
    tableData.push(campaignRow);
    addedRows++;
  });
  
  return addedRows;
}

function applyTrickyOptimizedFormatting(spreadsheetId, sheetId, sheetName, numRows, numCols, formatData, hyperlinkData, appData) {
  console.log('Применение TRICKY оптимизированного форматирования...');
  const requests = [];
  
  console.log('  TRICKY заголовки...');
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
  
  console.log('  TRICKY ширина колонок...');
  const columnWidths = [
    { index: 0, width: 80 }, { index: 1, width: 300 }, { index: 2, width: 40 }, { index: 3, width: 40 },
    { index: 4, width: 75 }, { index: 5, width: 55 }, { index: 6, width: 55 }, { index: 7, width: 55 },
    { index: 8, width: 55 }, { index: 9, width: 55 }, { index: 10, width: 55 }, { index: 11, width: 55 },
    { index: 12, width: 55 }, { index: 13, width: 55 }, { index: 14, width: 55 }, { index: 15, width: 75 },
    { index: 16, width: 85 }, { index: 17, width: 160 }, { index: 18, width: 250 }
  ];
  
  columnWidths.forEach(col => {
    requests.push({
      updateDimensionProperties: {
        range: {
          sheetId: sheetId,
          dimension: 'COLUMNS',
          startIndex: col.index,
          endIndex: col.index + 1
        },
        properties: {
          pixelSize: col.width
        },
        fields: 'pixelSize'
      }
    });
  });
  
  console.log('  TRICKY группировка строк...');
  const rowsByType = {
    app: [],
    week: [],
    sourceApp: [],
    campaign: []
  };
  
  formatData.forEach(item => {
    const rowIndex = item.row - 1;
    if (item.type === 'APP') rowsByType.app.push(rowIndex);
    if (item.type === 'WEEK') rowsByType.week.push(rowIndex);
    if (item.type === 'SOURCE_APP') rowsByType.sourceApp.push(rowIndex);
    if (item.type === 'CAMPAIGN') rowsByType.campaign.push(rowIndex);
  });

  console.log(`    TRICKY APP строк: ${rowsByType.app.length}`);
  console.log(`    TRICKY WEEK строк: ${rowsByType.week.length}`);
  console.log(`    TRICKY SOURCE_APP строк: ${rowsByType.sourceApp.length}`);
  console.log(`    TRICKY CAMPAIGN строк: ${rowsByType.campaign.length}`);

  const formatRanges = [
    { rows: rowsByType.app, bg: { red: 0.82, green: 0.91, blue: 1 }, bold: true, size: 10 },
    { rows: rowsByType.week, bg: { red: 0.91, green: 0.94, blue: 1 }, bold: false, size: 10 },
    { rows: rowsByType.sourceApp, bg: { red: 0.94, green: 0.97, blue: 1 }, bold: false, size: 9 },
    { rows: rowsByType.campaign, bg: { red: 1, green: 1, blue: 1 }, bold: false, size: 9 }
  ];
  
  console.log('  TRICKY применение форматирования строк...');
  formatRanges.forEach((format, formatIndex) => {
    console.log(`    TRICKY формат ${formatIndex + 1}: ${format.rows.length} строк`);
    format.rows.forEach(rowIndex => {
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: rowIndex,
            endRowIndex: rowIndex + 1,
            startColumnIndex: 0,
            endColumnIndex: numCols
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: format.bg,
              textFormat: { 
                bold: format.bold,
                fontSize: format.size
              },
              verticalAlignment: 'MIDDLE'
            }
          },
          fields: 'userEnteredFormat'
        }
      });
    });
  });
  
  console.log('  TRICKY числовые форматы...');
  if (numRows > 1) {
    const numberFormats = [
      { range: [4, 5], pattern: '$0.00' },
      { range: [7, 8], pattern: '$0.000' },
      { range: [12, 13], pattern: '$0.000' },
      { range: [15, 16], pattern: '$0.00' }
    ];
    
    numberFormats.forEach(format => {
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: 1,
            endRowIndex: numRows,
            startColumnIndex: format.range[0],
            endColumnIndex: format.range[1]
          },
          cell: {
            userEnteredFormat: {
              numberFormat: {
                type: 'CURRENCY',
                pattern: format.pattern
              }
            }
          },
          fields: 'userEnteredFormat.numberFormat'
        }
      });
    });
  }
  
  console.log(`  TRICKY отправка ${requests.length} базовых запросов...`);
  const batchSize = 100;
  let processedRequests = 0;
  
  for (let i = 0; i < requests.length; i += batchSize) {
    const batch = requests.slice(i, i + batchSize);
    
    try {
      Sheets.Spreadsheets.batchUpdate({
        requests: batch
      }, spreadsheetId);
      
      processedRequests += batch.length;
      console.log(`    TRICKY обработано: ${processedRequests}/${requests.length} запросов`);
      
      if (batch.length === batchSize && i + batchSize < requests.length) {
        Utilities.sleep(100);
      }
    } catch (e) {
      console.log(`    ⚠️ TRICKY ошибка в пакете ${i}-${i + batch.length}: ${e.toString()}`);
    }
  }
  
  console.log('  TRICKY детальное условное форматирование...');
  applyAdvancedConditionalFormatting(spreadsheetId, sheetId, sheetName, numRows, appData);
  
  console.log('✅ TRICKY оптимизированное форматирование завершено');
}

function applyCompleteFormatting(spreadsheetId, sheetId, sheetName, numRows, numCols, formatData, appData) {
  console.log('Применение полного форматирования...');
  const requests = [];
  
  console.log('  Заголовки...');
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
  
  console.log('  Ширина колонок...');
  const columnWidths = [
    { index: 0, width: 80 }, { index: 1, width: 300 }, { index: 2, width: 40 }, { index: 3, width: 40 },
    { index: 4, width: 75 }, { index: 5, width: 55 }, { index: 6, width: 55 }, { index: 7, width: 55 },
    { index: 8, width: 55 }, { index: 9, width: 55 }, { index: 10, width: 55 }, { index: 11, width: 55 },
    { index: 12, width: 55 }, { index: 13, width: 55 }, { index: 14, width: 55 }, { index: 15, width: 75 },
    { index: 16, width: 85 }, { index: 17, width: 160 }, { index: 18, width: 250 }
  ];
  
  columnWidths.forEach(col => {
    requests.push({
      updateDimensionProperties: {
        range: {
          sheetId: sheetId,
          dimension: 'COLUMNS',
          startIndex: col.index,
          endIndex: col.index + 1
        },
        properties: {
          pixelSize: col.width
        },
        fields: 'pixelSize'
      }
    });
  });
  
  console.log('  Группировка строк по типам...');
  const rowsByType = {
    app: [],
    week: [],
    sourceApp: [],
    campaign: []
  };
  
  formatData.forEach(item => {
    const rowIndex = item.row - 1;
    if (item.type === 'APP') rowsByType.app.push(rowIndex);
    if (item.type === 'WEEK') rowsByType.week.push(rowIndex);
    if (item.type === 'SOURCE_APP') rowsByType.sourceApp.push(rowIndex);
    if (item.type === 'CAMPAIGN') rowsByType.campaign.push(rowIndex);
  });

  console.log(`    APP строк: ${rowsByType.app.length}`);
  console.log(`    WEEK строк: ${rowsByType.week.length}`);
  console.log(`    SOURCE_APP строк: ${rowsByType.sourceApp.length}`);
  console.log(`    CAMPAIGN строк: ${rowsByType.campaign.length}`);

  const formatRanges = [
    { rows: rowsByType.app, bg: { red: 0.82, green: 0.91, blue: 1 }, bold: true, size: 10 },
    { rows: rowsByType.week, bg: { red: 0.91, green: 0.94, blue: 1 }, bold: false, size: 10 },
    { rows: rowsByType.sourceApp, bg: { red: 0.94, green: 0.97, blue: 1 }, bold: false, size: 9 },
    { rows: rowsByType.campaign, bg: { red: 1, green: 1, blue: 1 }, bold: false, size: 9 }
  ];
  
  console.log('  Применение форматирования строк...');
  formatRanges.forEach((format, formatIndex) => {
    console.log(`    Формат ${formatIndex + 1}: ${format.rows.length} строк`);
    format.rows.forEach(rowIndex => {
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: rowIndex,
            endRowIndex: rowIndex + 1,
            startColumnIndex: 0,
            endColumnIndex: numCols
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: format.bg,
              textFormat: { 
                bold: format.bold,
                fontSize: format.size
              },
              verticalAlignment: 'MIDDLE'
            }
          },
          fields: 'userEnteredFormat'
        }
      });
    });
  });
  
  console.log('  Числовые форматы...');
  if (numRows > 1) {
    const numberFormats = [
      { range: [4, 5], pattern: '$0.00' },
      { range: [7, 8], pattern: '$0.000' },
      { range: [12, 13], pattern: '$0.000' },
      { range: [15, 16], pattern: '$0.00' }
    ];
    
    numberFormats.forEach(format => {
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: 1,
            endRowIndex: numRows,
            startColumnIndex: format.range[0],
            endColumnIndex: format.range[1]
          },
          cell: {
            userEnteredFormat: {
              numberFormat: {
                type: 'CURRENCY',
                pattern: format.pattern
              }
            }
          },
          fields: 'userEnteredFormat.numberFormat'
        }
      });
    });
  }
  
  console.log(`  Отправка ${requests.length} базовых запросов...`);
  const batchSize = 100;
  let processedRequests = 0;
  
  for (let i = 0; i < requests.length; i += batchSize) {
    const batch = requests.slice(i, i + batchSize);
    
    try {
      Sheets.Spreadsheets.batchUpdate({
        requests: batch
      }, spreadsheetId);
      
      processedRequests += batch.length;
      console.log(`    Обработано: ${processedRequests}/${requests.length} запросов`);
      
      if (batch.length === batchSize && i + batchSize < requests.length) {
        Utilities.sleep(100);
      }
    } catch (e) {
      console.log(`    ⚠️ Ошибка в пакете ${i}-${i + batch.length}: ${e.toString()}`);
    }
  }
  
  console.log('  Детальное условное форматирование...');
  applyAdvancedConditionalFormatting(spreadsheetId, sheetId, sheetName, numRows, appData);
  
  console.log('✅ Полное форматирование завершено');
}

function applyAdvancedConditionalFormatting(spreadsheetId, sheetId, sheetName, numRows, appData) {
  console.log('Применение детального условного форматирования...');
  
  try {
    const conditionalRequests = [];
    
    console.log('  WoW изменения...');
    conditionalRequests.push({
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
              type: 'CUSTOM_FORMULA',
              values: [{
                userEnteredValue: '=AND(NOT(ISBLANK(F2)), ISNUMBER(VALUE(SUBSTITUTE(F2,"%",""))), VALUE(SUBSTITUTE(F2,"%","")) > 0)'
              }]
            },
            format: {
              backgroundColor: { red: 0.82, green: 0.94, blue: 0.92 },
              textFormat: { foregroundColor: { red: 0.05, green: 0.33, blue: 0.38 } }
            }
          }
        },
        index: 0
      }
    });
    
    conditionalRequests.push({
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
              type: 'CUSTOM_FORMULA',
              values: [{
                userEnteredValue: '=AND(NOT(ISBLANK(F2)), ISNUMBER(VALUE(SUBSTITUTE(F2,"%",""))), VALUE(SUBSTITUTE(F2,"%","")) < 0)'
              }]
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
    
    conditionalRequests.push({
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
              type: 'CUSTOM_FORMULA',
              values: [{
                userEnteredValue: '=AND(NOT(ISBLANK(Q2)), ISNUMBER(VALUE(SUBSTITUTE(Q2,"%",""))), VALUE(SUBSTITUTE(Q2,"%","")) > 0)'
              }]
            },
            format: {
              backgroundColor: { red: 0.82, green: 0.94, blue: 0.92 },
              textFormat: { foregroundColor: { red: 0.05, green: 0.33, blue: 0.38 } }
            }
          }
        },
        index: 2
      }
    });
    
    conditionalRequests.push({
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
              type: 'CUSTOM_FORMULA',
              values: [{
                userEnteredValue: '=AND(NOT(ISBLANK(Q2)), ISNUMBER(VALUE(SUBSTITUTE(Q2,"%",""))), VALUE(SUBSTITUTE(Q2,"%","")) < 0)'
              }]
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
    
    console.log('  Индивидуальные eROAS правила...');
    try {
      const sheetValues = Sheets.Spreadsheets.Values.get(spreadsheetId, `${sheetName}!A:O`).values;
      
      if (sheetValues && sheetValues.length > 1) {
        let ruleIndex = 4;
        
        for (let i = 1; i < Math.min(sheetValues.length, numRows); i++) {
          const level = sheetValues[i][0];
          let appName = '';
          let targetEROAS = 150;
          
          if (level === 'APP') {
            appName = sheetValues[i][1];
            targetEROAS = getTargetEROAS(CURRENT_PROJECT, appName);
          } else {
            for (let j = i - 1; j >= 1; j--) {
              if (sheetValues[j][0] === 'APP') {
                appName = sheetValues[j][1];
                targetEROAS = getTargetEROAS(CURRENT_PROJECT, appName);
                break;
              }
            }
          }
          
          const cellFormula = `O${i + 1}`;
          
          conditionalRequests.push({
            addConditionalFormatRule: {
              rule: {
                ranges: [{
                  sheetId: sheetId,
                  startRowIndex: i,
                  endRowIndex: i + 1,
                  startColumnIndex: 14,
                  endColumnIndex: 15
                }],
                booleanRule: {
                  condition: {
                    type: 'CUSTOM_FORMULA',
                    values: [{
                      userEnteredValue: `=AND(NOT(ISBLANK(${cellFormula})), VALUE(SUBSTITUTE(${cellFormula},"%","")) >= ${targetEROAS})`
                    }]
                  },
                  format: {
                    backgroundColor: { red: 0.83, green: 0.93, blue: 0.85 },
                    textFormat: { foregroundColor: { red: 0.08, green: 0.34, blue: 0.14 } }
                  }
                }
              },
              index: ruleIndex++
            }
          });
          
          conditionalRequests.push({
            addConditionalFormatRule: {
              rule: {
                ranges: [{
                  sheetId: sheetId,
                  startRowIndex: i,
                  endRowIndex: i + 1,
                  startColumnIndex: 14,
                  endColumnIndex: 15
                }],
                booleanRule: {
                  condition: {
                    type: 'CUSTOM_FORMULA',
                    values: [{
                      userEnteredValue: `=AND(NOT(ISBLANK(${cellFormula})), VALUE(SUBSTITUTE(${cellFormula},"%","")) >= 120, VALUE(SUBSTITUTE(${cellFormula},"%","")) < ${targetEROAS})`
                    }]
                  },
                  format: {
                    backgroundColor: { red: 1, green: 0.95, blue: 0.8 },
                    textFormat: { foregroundColor: { red: 0.52, green: 0.39, blue: 0.02 } }
                  }
                }
              },
              index: ruleIndex++
            }
          });
          
          conditionalRequests.push({
            addConditionalFormatRule: {
              rule: {
                ranges: [{
                  sheetId: sheetId,
                  startRowIndex: i,
                  endRowIndex: i + 1,
                  startColumnIndex: 14,
                  endColumnIndex: 15
                }],
                booleanRule: {
                  condition: {
                    type: 'CUSTOM_FORMULA',
                    values: [{
                      userEnteredValue: `=AND(NOT(ISBLANK(${cellFormula})), VALUE(SUBSTITUTE(${cellFormula},"%","")) < 120)`
                    }]
                  },
                  format: {
                    backgroundColor: { red: 0.97, green: 0.84, blue: 0.85 },
                    textFormat: { foregroundColor: { red: 0.45, green: 0.11, blue: 0.14 } }
                  }
                }
              },
              index: ruleIndex++
            }
          });
          
          if (ruleIndex >= 100) break;
        }
      }
    } catch (e) {
      console.log('⚠️ Ошибка получения данных для eROAS правил:', e.toString());
    }
    
    console.log('  Growth Status цвета...');
    const statusColors = [
      { text: '🟢 Healthy Growth', bg: { red: 0.83, green: 0.93, blue: 0.85 }, fg: { red: 0.08, green: 0.34, blue: 0.14 } },
      { text: '🟢 Efficiency Improvement', bg: { red: 0.82, green: 0.95, blue: 0.92 }, fg: { red: 0.05, green: 0.33, blue: 0.38 } },
      { text: '🔴 Inefficient Growth', bg: { red: 0.97, green: 0.84, blue: 0.85 }, fg: { red: 0.45, green: 0.11, blue: 0.14 } },
      { text: '🟠 Declining Efficiency', bg: { red: 1, green: 0.6, blue: 0 }, fg: { red: 1, green: 1, blue: 1 } },
      { text: '🔵 Scaling Down', bg: { red: 0.8, green: 0.91, blue: 1 }, fg: { red: 0, green: 0.25, blue: 0.52 } },
      { text: '🟡 Moderate Growth', bg: { red: 1, green: 0.95, blue: 0.8 }, fg: { red: 0.52, green: 0.39, blue: 0.02 } },
      { text: '🟡 Moderate Decline', bg: { red: 1, green: 0.95, blue: 0.8 }, fg: { red: 0.52, green: 0.39, blue: 0.02 } },
      { text: '⚪ Stable', bg: { red: 0.96, green: 0.96, blue: 0.96 }, fg: { red: 0.38, green: 0.38, blue: 0.38 } }
    ];
    
    statusColors.forEach((status, statusIndex) => {
      conditionalRequests.push({
        addConditionalFormatRule: {
          rule: {
            ranges: [{
              sheetId: sheetId,
              startRowIndex: 1,
              endRowIndex: numRows,
              startColumnIndex: 17,
              endColumnIndex: 18
            }],
            booleanRule: {
              condition: {
                type: 'TEXT_CONTAINS',
                values: [{ userEnteredValue: status.text }]
              },
              format: {
                backgroundColor: status.bg,
                textFormat: { foregroundColor: status.fg }
              }
            }
          },
          index: 200 + statusIndex
        }
      });
    });
    
    console.log(`  Отправка ${conditionalRequests.length} условных правил...`);
    const condBatchSize = 50;
    
    for (let i = 0; i < conditionalRequests.length; i += condBatchSize) {
      const batch = conditionalRequests.slice(i, i + condBatchSize);
      
      try {
        Sheets.Spreadsheets.batchUpdate({
          requests: batch
        }, spreadsheetId);
        
        console.log(`    Условные правила: ${i + batch.length}/${conditionalRequests.length}`);
        
        if (batch.length === condBatchSize && i + condBatchSize < conditionalRequests.length) {
          Utilities.sleep(200);
        }
      } catch (e) {
        console.log(`    ⚠️ Ошибка условных правил ${i}-${i + batch.length}: ${e.toString()}`);
      }
    }
    
  } catch (e) {
    console.log('⚠️ Ошибка детального условного форматирования:', e.toString());
  }
  
  console.log('✅ Детальное условное форматирование завершено');
}

function applyBatchGrouping(spreadsheetId, sheetId, groupingData) {
  console.log(`Применение пакетной группировки: ${groupingData.length} групп...`);
  
  if (groupingData.length === 0) {
    console.log('Группы отсутствуют, пропускаем');
    return;
  }
  
  try {
    console.log('  Создание групп...');
    const groupRequests = [];
    
    groupingData.forEach((group, index) => {
      if (group.rowCount > 0) {
        console.log(`    Группа ${index + 1}: строки ${group.startRow}-${group.startRow + group.rowCount - 1} (${group.type})`);
        groupRequests.push({
          addDimensionGroup: {
            range: {
              sheetId: sheetId,
              dimension: 'ROWS',
              startIndex: group.startRow,
              endIndex: group.startRow + group.rowCount
            }
          }
        });
      }
    });
    
    if (groupRequests.length > 0) {
      console.log(`  Отправка ${groupRequests.length} запросов создания групп...`);
      const batchSize = 50;
      
      for (let i = 0; i < groupRequests.length; i += batchSize) {
        const batch = groupRequests.slice(i, i + batchSize);
        
        try {
          Sheets.Spreadsheets.batchUpdate({
            requests: batch
          }, spreadsheetId);
          
          console.log(`    Создано групп: ${i + batch.length}/${groupRequests.length}`);
          
          if (batch.length === batchSize && i + batchSize < groupRequests.length) {
            Utilities.sleep(200);
          }
        } catch (e) {
          console.log(`    ⚠️ Ошибка создания групп ${i}-${i + batch.length}: ${e.toString()}`);
        }
      }
      
      console.log('  Сворачивание групп...');
      const collapseRequests = [];
      groupingData.forEach(group => {
        if (group.rowCount > 0) {
          collapseRequests.push({
            updateDimensionGroup: {
              dimensionGroup: {
                range: {
                  sheetId: sheetId,
                  dimension: 'ROWS',
                  startIndex: group.startRow,
                  endIndex: group.startRow + group.rowCount
                },
                depth: group.depth,
                collapsed: true
              },
              fields: 'collapsed'
            }
          });
        }
      });
      
      if (collapseRequests.length > 0) {
        console.log(`  Отправка ${collapseRequests.length} запросов сворачивания...`);
        
        for (let i = 0; i < collapseRequests.length; i += batchSize) {
          const batch = collapseRequests.slice(i, i + batchSize);
          
          try {
            Sheets.Spreadsheets.batchUpdate({
              requests: batch
            }, spreadsheetId);
            
            console.log(`    Свернуто групп: ${i + batch.length}/${collapseRequests.length}`);
            
            if (batch.length === batchSize && i + batchSize < collapseRequests.length) {
              Utilities.sleep(200);
            }
          } catch (e) {
            console.log(`    ⚠️ Ошибка сворачивания групп ${i}-${i + batch.length}: ${e.toString()}`);
          }
        }
      }
    }
    
    console.log('✅ Пакетная группировка завершена');
    
  } catch (e) {
    console.log('⚠️ Общая ошибка группировки (не критично):', e.toString());
  }
}

function initTrickyOptimizedCache() {
  try {
    const appsDb = new AppsDatabase('TRICKY');
    appsDb.ensureCacheUpToDate();
    const appsDbCache = appsDb.loadFromCache();
    console.log(`TRICKY кеш инициализирован: ${Object.keys(appsDbCache).length} приложений`);
    return { appsDbCache };
  } catch (e) {
    console.log('Ошибка инициализации TRICKY кеша:', e);
    return { appsDbCache: {} };
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

function calculateWeekTotals(campaigns) {
  if (!campaigns || campaigns.length === 0) {
    return {
      totalSpend: 0, totalInstalls: 0, avgCpi: 0, avgRoas: 0, avgIpm: 0, 
      avgRrD1: 0, avgRrD7: 0, avgArpu: 0, avgERoas: 0, avgEROASD730: 0, totalProfit: 0
    };
  }
  
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

function createCampaignRow(campaign, campaignIdValue, spendPct, profitPct, growthStatus) {
  return [
    'CAMPAIGN', campaign.sourceApp, campaignIdValue, campaign.geo,
    campaign.spend.toFixed(2), spendPct, campaign.installs, campaign.cpi ? campaign.cpi.toFixed(3) : '0.000',
    campaign.roas.toFixed(2), campaign.ipm.toFixed(1), `${campaign.rrD1.toFixed(1)}%`, `${campaign.rrD7.toFixed(1)}%`,
    campaign.eArpuForecast.toFixed(3), `${campaign.eRoasForecast.toFixed(0)}%`, `${campaign.eRoasForecastD730.toFixed(0)}%`,
    campaign.eProfitForecast.toFixed(2), profitPct, growthStatus, ''
  ];
}

function getColumnLetter(columnIndex) {
  let letter = '';
  let tempIndex = columnIndex;
  
  while (tempIndex >= 0) {
    letter = String.fromCharCode(65 + (tempIndex % 26)) + letter;
    tempIndex = Math.floor(tempIndex / 26) - 1;
  }
  
  return letter;
}

function getSheetByName(spreadsheetId, sheetName) {
  console.log(`Получение листа: ${sheetName} из таблицы ${spreadsheetId}`);
  try {
    const spreadsheet = Sheets.Spreadsheets.get(spreadsheetId);
    console.log(`Таблица получена, листов: ${spreadsheet.sheets.length}`);
    
    const sheet = spreadsheet.sheets.find(s => s.properties.title === sheetName);
    if (sheet) {
      console.log(`✅ Лист найден: ${sheetName} (ID: ${sheet.properties.sheetId})`);
      return sheet;
    } else {
      console.log(`❌ Лист не найден: ${sheetName}`);
      const sheetNames = spreadsheet.sheets.map(s => s.properties.title).join(', ');
      console.log(`Доступные листы: ${sheetNames}`);
      return null;
    }
  } catch (e) {
    console.error('❌ Ошибка получения листа:', e);
    return null;
  }
}