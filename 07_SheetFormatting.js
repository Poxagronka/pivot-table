function createEnhancedPivotTable(appData) {
  console.log('=== СОЗДАНИЕ ENHANCED PIVOT TABLE ===');
  console.log(`Получено приложений: ${Object.keys(appData).length}`);
  
  const config = getCurrentConfig();
  console.log(`Конфигурация: Sheet ID = ${config.SHEET_ID}, Sheet Name = ${config.SHEET_NAME}`);
  
  console.log('Этап 1: Расчет WoW метрик...');
  const wow = calculateWoWMetrics(appData);
  console.log(`WoW метрики рассчитаны: ${Object.keys(wow.campaignWoW).length} кампаний, ${Object.keys(wow.appWeekWoW).length} недель приложений`);
  
  console.log('Этап 2: Подготовка заголовков и данных...');
  const headers = getUnifiedHeaders();
  console.log(`Заголовков: ${headers.length}`);
  
  const tableData = [headers];
  const formatData = [];
  let totalRows = 1;

  const appKeys = Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName));
  console.log(`Обрабатываем приложения: ${appKeys.length} штук`);

  appKeys.forEach((appKey, appIndex) => {
    const app = appData[appKey];
    console.log(`Приложение ${appIndex + 1}/${appKeys.length}: ${app.appName}`);
    
    formatData.push({ row: tableData.length + 1, type: 'APP' });
    const emptyRow = new Array(headers.length).fill('');
    emptyRow[0] = 'APP';
    emptyRow[1] = app.appName;
    tableData.push(emptyRow);
    totalRows++;

    const weekKeys = Object.keys(app.weeks).sort();
    console.log(`  Недель для ${app.appName}: ${weekKeys.length}`);

    weekKeys.forEach((weekKey, weekIndex) => {
      const week = app.weeks[weekKey];
      console.log(`  Неделя ${weekIndex + 1}/${weekKeys.length}: ${weekKey}`);
      
      formatData.push({ row: tableData.length + 1, type: 'WEEK' });
      
      if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
        console.log(`    TRICKY проект - обрабатываем sourceApps: ${Object.keys(week.sourceApps).length}`);
        
        const allCampaigns = [];
        Object.values(week.sourceApps).forEach(sourceApp => {
          allCampaigns.push(...sourceApp.campaigns);
        });
        console.log(`    Всего кампаний в неделе: ${allCampaigns.length}`);
        
        const weekTotals = calculateWeekTotals(allCampaigns);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = wow.appWeekWoW[appWeekKey] || {};
        
        const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
        const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
        const status = weekWoW.growthStatus || '';
        
        const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status);
        tableData.push(weekRow);
        totalRows++;
        
        console.log(`    Добавляем source app строки...`);
        const sourceAppRowsAdded = addSourceAppRows(tableData, week.sourceApps, weekKey, wow, formatData);
        totalRows += sourceAppRowsAdded;
        console.log(`    Добавлено source app строк: ${sourceAppRowsAdded}`);
        
      } else {
        console.log(`    Стандартный проект - обрабатываем кампании: ${week.campaigns?.length || 0}`);
        
        const weekTotals = calculateWeekTotals(week.campaigns || []);
        const appWeekKey = `${app.appName}_${weekKey}`;
        const weekWoW = wow.appWeekWoW[appWeekKey] || {};
        
        const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
        const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
        const status = weekWoW.growthStatus || '';
        
        const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status);
        tableData.push(weekRow);
        totalRows++;
        
        console.log(`    Добавляем кампании...`);
        const campaignRowsAdded = addCampaignRows(tableData, week.campaigns || [], week, weekKey, wow, formatData);
        totalRows += campaignRowsAdded;
        console.log(`    Добавлено строк кампаний: ${campaignRowsAdded}`);
      }
    });
  });

  console.log(`Этап 3: Подготовка данных завершена. Всего строк: ${totalRows}`);
  console.log(`Размер tableData: ${tableData.length} строк`);
  console.log(`Размер formatData: ${formatData.length} элементов форматирования`);

  console.log('Этап 4: Запись таблицы через Sheets API...');
  writeTableWithSheetsAPI(config, tableData, formatData, headers.length, appData);
  
  console.log('=== ENHANCED PIVOT TABLE СОЗДАНА ===');
}

function createOverallPivotTable(appData) {
  console.log('=== СОЗДАНИЕ OVERALL PIVOT TABLE ===');
  console.log(`Получено приложений: ${Object.keys(appData).length}`);
  
  const config = getCurrentConfig();
  console.log(`Конфигурация: Sheet ID = ${config.SHEET_ID}, Sheet Name = ${config.SHEET_NAME}`);
  
  console.log('Этап 1: Расчет WoW метрик...');
  const wow = calculateWoWMetrics(appData);
  console.log(`WoW метрики рассчитаны: ${Object.keys(wow.appWeekWoW).length} недель приложений`);
  
  console.log('Этап 2: Подготовка заголовков и данных...');
  const headers = getUnifiedHeaders();
  console.log(`Заголовков: ${headers.length}`);
  
  const tableData = [headers];
  const formatData = [];
  let totalRows = 1;

  const appKeys = Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName));
  console.log(`Обрабатываем приложения: ${appKeys.length} штук`);

  appKeys.forEach((appKey, appIndex) => {
    const app = appData[appKey];
    console.log(`Приложение ${appIndex + 1}/${appKeys.length}: ${app.appName}`);
    
    formatData.push({ row: tableData.length + 1, type: 'APP' });
    const emptyRow = new Array(headers.length).fill('');
    emptyRow[0] = 'APP';
    emptyRow[1] = app.appName;
    tableData.push(emptyRow);
    totalRows++;

    const weekKeys = Object.keys(app.weeks).sort();
    console.log(`  Недель для ${app.appName}: ${weekKeys.length}`);

    weekKeys.forEach((weekKey, weekIndex) => {
      const week = app.weeks[weekKey];
      console.log(`  Неделя ${weekIndex + 1}/${weekKeys.length}: ${weekKey}`);
      
      const weekTotals = calculateWeekTotals(week.campaigns || []);
      const appWeekKey = `${app.appName}_${weekKey}`;
      const weekWoW = wow.appWeekWoW[appWeekKey] || {};
      
      const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = weekWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'WEEK' });
      const weekRow = createWeekRow(week, weekTotals, spendWoW, profitWoW, status);
      tableData.push(weekRow);
      totalRows++;
    });
  });

  console.log(`Этап 3: Подготовка данных завершена. Всего строк: ${totalRows}`);
  console.log(`Размер tableData: ${tableData.length} строк`);
  console.log(`Размер formatData: ${formatData.length} элементов форматирования`);

  console.log('Этап 4: Запись таблицы через Sheets API...');
  writeTableWithSheetsAPI(config, tableData, formatData, headers.length, appData);
  
  console.log('=== OVERALL PIVOT TABLE СОЗДАНА ===');
}

function writeTableWithSheetsAPI(config, tableData, formatData, numCols, appData, knownSheetId = null) {
  console.log('=== ЗАПИСЬ ТАБЛИЦЫ ЧЕРЕЗ SHEETS API ===');
  const numRows = tableData.length;
  const sheetName = config.SHEET_NAME;
  
  console.log(`Записываем таблицу: ${numRows} строк x ${numCols} колонок`);
  console.log(`Лист: ${sheetName}`);
  
  console.log('Этап 1: Получение Sheet ID...');
  const sheetId = knownSheetId || getSheetIdOnce(config.SHEET_ID, sheetName);
  console.log(`Sheet ID: ${sheetId}`);
  
  console.log('Этап 2: Запись данных батчами...');
  writeDataInBatches(config.SHEET_ID, sheetId, tableData, numCols);
  console.log('✅ Данные записаны');
  
  console.log('Этап 3: Применение форматирования...');
  applyAllFormatting(config.SHEET_ID, sheetId, formatData, numRows, numCols);
  console.log('✅ Форматирование применено');
  
  console.log('Этап 4: Создание группировки...');
  applyGrouping(config.SHEET_ID, sheetId, formatData);
  console.log('✅ Группировка создана');
  
  console.log('=== ЗАПИСЬ ТАБЛИЦЫ ЗАВЕРШЕНА ===');
}

function getSheetIdOnce(spreadsheetId, sheetName) {
  console.log(`Получение Sheet ID для листа: ${sheetName}`);
  try {
    Utilities.sleep(500);
    const spreadsheet = Sheets.Spreadsheets.get(spreadsheetId);
    const sheet = spreadsheet.sheets.find(s => s.properties.title === sheetName);
    if (sheet) {
      console.log(`✅ Sheet ID найден: ${sheet.properties.sheetId}`);
      return sheet.properties.sheetId;
    } else {
      console.log('❌ Sheet не найден, используем ID = 0');
      return 0;
    }
  } catch (e) {
    console.log('⚠️ Ошибка получения Sheet ID, повторная попытка через 1 сек...');
    Utilities.sleep(1000);
    try {
      const spreadsheet = Sheets.Spreadsheets.get(spreadsheetId);
      const sheet = spreadsheet.sheets.find(s => s.properties.title === sheetName);
      if (sheet) {
        console.log(`✅ Sheet ID найден при повторе: ${sheet.properties.sheetId}`);
        return sheet.properties.sheetId;
      } else {
        console.log('❌ Sheet не найден при повторе, используем ID = 0');
        return 0;
      }
    } catch (e2) {
      console.error('❌ Критическая ошибка получения Sheet ID:', e2);
      return 0;
    }
  }
}

function writeDataInBatches(spreadsheetId, sheetId, tableData, numCols) {
  const batchSize = 100;
  const totalRows = tableData.length;
  
  console.log(`Запись данных батчами: ${totalRows} строк, батч размер ${batchSize}`);
  
  for (let i = 0; i < totalRows; i += batchSize) {
    const endRow = Math.min(i + batchSize, totalRows);
    const batchData = tableData.slice(i, endRow);
    
    console.log(`  Записываем батч: строки ${i + 1}-${endRow}`);
    
    const request = {
      updateCells: {
        range: {
          sheetId: sheetId,
          startRowIndex: i,
          endRowIndex: endRow,
          startColumnIndex: 0,
          endColumnIndex: numCols
        },
        rows: batchData.map(row => ({
          values: row.map(cell => ({
            userEnteredValue: { stringValue: cell?.toString() || '' }
          }))
        })),
        fields: 'userEnteredValue'
      }
    };
    
    try {
      Sheets.Spreadsheets.batchUpdate({ requests: [request] }, spreadsheetId);
      console.log(`  ✅ Батч ${i + 1}-${endRow} записан успешно`);
    } catch (e) {
      console.error(`  ❌ Ошибка записи батча ${i + 1}-${endRow}:`, e);
      throw e;
    }
    
    if (endRow < totalRows) {
      Utilities.sleep(100);
    }
  }
  
  console.log('✅ Все данные записаны');
}

function applyAllFormatting(spreadsheetId, sheetId, formatData, numRows, numCols) {
  console.log('Применение всего форматирования...');
  const requests = [];
  
  console.log('  Добавляем форматирование заголовков...');
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
  
  console.log('  Добавляем ширину колонок...');
  const widths = [80, 300, 40, 40, 75, 55, 55, 55, 55, 55, 55, 55, 55, 55, 55, 75, 85, 160, 250];
  widths.forEach((width, index) => {
    requests.push({
      updateDimensionProperties: {
        range: {
          sheetId: sheetId,
          dimension: 'COLUMNS',
          startIndex: index,
          endIndex: index + 1
        },
        properties: {
          pixelSize: width
        },
        fields: 'pixelSize'
      }
    });
  });
  
  console.log('  Добавляем форматирование строк...');
  const appRows = formatData.filter(f => f.type === 'APP').map(f => f.row - 1);
  const weekRows = formatData.filter(f => f.type === 'WEEK').map(f => f.row - 1);
  const sourceAppRows = formatData.filter(f => f.type === 'SOURCE_APP').map(f => f.row - 1);
  const campaignRows = formatData.filter(f => f.type === 'CAMPAIGN').map(f => f.row - 1);
  
  console.log(`    APP строк: ${appRows.length}`);
  console.log(`    WEEK строк: ${weekRows.length}`);
  console.log(`    SOURCE_APP строк: ${sourceAppRows.length}`);
  console.log(`    CAMPAIGN строк: ${campaignRows.length}`);
  
  appRows.forEach(rowIndex => {
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
            backgroundColor: { red: 0.82, green: 0.91, blue: 0.996 },
            textFormat: { bold: true, fontSize: 10 }
          }
        },
        fields: 'userEnteredFormat'
      }
    });
  });
  
  weekRows.forEach(rowIndex => {
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
            backgroundColor: { red: 0.91, green: 0.94, blue: 0.996 },
            textFormat: { fontSize: 10 }
          }
        },
        fields: 'userEnteredFormat'
      }
    });
  });
  
  sourceAppRows.forEach(rowIndex => {
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
            backgroundColor: { red: 0.94, green: 0.97, blue: 1 },
            textFormat: { fontSize: 9 }
          }
        },
        fields: 'userEnteredFormat'
      }
    });
  });
  
  campaignRows.forEach(rowIndex => {
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
            backgroundColor: { red: 1, green: 1, blue: 1 },
            textFormat: { fontSize: 9 }
          }
        },
        fields: 'userEnteredFormat'
      }
    });
  });
  
  console.log('  Добавляем условное форматирование...');
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
              type: 'TEXT_CONTAINS',
              values: [{ userEnteredValue: '%' }]
            },
            format: {
              backgroundColor: { red: 0.82, green: 0.95, blue: 0.92 }
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
            startColumnIndex: 14,
            endColumnIndex: 15
          }],
          booleanRule: {
            condition: {
              type: 'CUSTOM_FORMULA',
              values: [{ userEnteredValue: '=AND(NOT(ISBLANK(O2)), VALUE(SUBSTITUTE(O2,"%","")) >= 150)' }]
            },
            format: {
              backgroundColor: { red: 0.82, green: 0.95, blue: 0.92 }
            }
          }
        },
        index: 1
      }
    });
  }
  
  console.log('  Добавляем настройки листа...');
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
  
  console.log(`  Всего запросов форматирования: ${requests.length}`);
  
  try {
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
    console.log('✅ Форматирование применено успешно');
  } catch (e) {
    console.error('❌ Ошибка применения форматирования:', e);
    throw e;
  }
}

function applyGrouping(spreadsheetId, sheetId, formatData) {
  console.log('Создание группировки строк...');
  const requests = [];
  
  const appRows = formatData.filter(f => f.type === 'APP').map(f => f.row - 1);
  console.log(`Создаем группировку для ${appRows.length} приложений`);
  
  appRows.forEach((appRowIndex, index) => {
    const nextAppIndex = appRows[index + 1];
    const endRow = nextAppIndex ? nextAppIndex : formatData[formatData.length - 1].row;
    
    if (endRow > appRowIndex + 1) {
      console.log(`  Группа ${index + 1}: строки ${appRowIndex + 2}-${endRow}`);
      requests.push({
        addDimensionGroup: {
          range: {
            sheetId: sheetId,
            dimension: 'ROWS',
            startIndex: appRowIndex + 1,
            endIndex: endRow
          }
        }
      });
    }
  });
  
  if (requests.length > 0) {
    console.log(`Применяем ${requests.length} групп...`);
    try {
      Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
      console.log('✅ Группировка создана успешно');
    } catch (e) {
      console.error('❌ Ошибка создания группировки:', e);
      throw e;
    }
  } else {
    console.log('Нет групп для создания');
  }
}

function addSourceAppRows(tableData, sourceApps, weekKey, wow, formatData) {
  console.log(`    Добавление source app строк для недели ${weekKey}`);
  let addedRows = 0;
  
  const sourceAppKeys = Object.keys(sourceApps).sort((a, b) => {
    const totalSpendA = sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
    const totalSpendB = sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
    return totalSpendB - totalSpendA;
  });
  
  console.log(`    Source Apps: ${sourceAppKeys.length} штук`);
  
  sourceAppKeys.forEach((sourceAppKey, index) => {
    const sourceApp = sourceApps[sourceAppKey];
    console.log(`      Source App ${index + 1}/${sourceAppKeys.length}: ${sourceApp.sourceAppName} (${sourceApp.campaigns.length} кампаний)`);
    
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
          console.log(`        Добавлена гиперссылка для ${sourceApp.sourceAppName}`);
        }
      } catch (e) {
        console.log('        Ошибка получения ссылки на store:', e);
      }
    }
    
    const sourceAppRow = createSourceAppRow(sourceAppDisplayName, sourceAppTotals, spendWoW, profitWoW, status);
    tableData.push(sourceAppRow);
    addedRows++;
    
    console.log(`        Добавляем кампании для ${sourceApp.sourceAppName}...`);
    const campaignRowsAdded = addCampaignRows(tableData, sourceApp.campaigns, { weekStart: weekKey.split('-').join('/'), weekEnd: '' }, weekKey, wow, formatData);
    addedRows += campaignRowsAdded;
    console.log(`        Добавлено кампаний: ${campaignRowsAdded}`);
  });
  
  console.log(`    Всего добавлено source app строк: ${addedRows}`);
  return addedRows;
}

function addCampaignRows(tableData, campaigns, week, weekKey, wow, formatData) {
  if (CURRENT_PROJECT === 'OVERALL') {
    console.log('        OVERALL проект - кампании не добавляются');
    return 0;
  }
  
  if (!campaigns || campaigns.length === 0) {
    console.log('        Нет кампаний для добавления');
    return 0;
  }
  
  console.log(`        Добавление ${campaigns.length} кампаний`);
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
    
    formatData.push({ row: tableData.length + 1, type: 'CAMPAIGN' });
    
    const campaignRow = createCampaignRow(campaign, campaignIdValue, spendPct, profitPct, growthStatus);
    tableData.push(campaignRow);
    addedRows++;
    
    if (index < 3) {
      console.log(`          Кампания ${index + 1}: ${campaign.campaignName?.substring(0, 50) || 'Unknown'}... (spend: ${campaign.spend})`);
    }
  });
  
  console.log(`        Добавлено строк кампаний: ${addedRows}`);
  return addedRows;
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

function createProjectPivotTable(projectName, appData) {
  console.log(`=== СОЗДАНИЕ ТАБЛИЦЫ ДЛЯ ПРОЕКТА ${projectName} ===`);
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
  
  console.log(`=== ТАБЛИЦА ДЛЯ ПРОЕКТА ${projectName} СОЗДАНА ===`);
}