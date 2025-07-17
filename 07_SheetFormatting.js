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

  console.log('Этап 4: Запись таблицы...');
  writeTableSafely(config, tableData, formatData, headers.length, appData);
  
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

  console.log('Этап 4: Запись таблицы...');
  writeTableSafely(config, tableData, formatData, headers.length, appData);
  
  console.log('=== OVERALL PIVOT TABLE СОЗДАНА ===');
}

function writeTableSafely(config, tableData, formatData, numCols, appData) {
  console.log('=== БЕЗОПАСНАЯ ЗАПИСЬ ТАБЛИЦЫ ===');
  const numRows = tableData.length;
  const sheetName = config.SHEET_NAME;
  
  console.log(`Записываем таблицу: ${numRows} строк x ${numCols} колонок`);
  console.log(`Лист: ${sheetName}`);
  
  try {
    console.log('Этап 1: Безопасное получение листа...');
    const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
    
    let sheet = null;
    try {
      sheet = spreadsheet.getSheetByName(sheetName);
      console.log(`✅ Лист найден: ${sheetName}`);
    } catch (e) {
      console.log(`Лист не найден, создаем новый: ${sheetName}`);
    }
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      console.log(`✅ Лист создан: ${sheetName}`);
    }
    
    console.log('Этап 2: Очистка листа...');
    try {
      sheet.clear();
      console.log('✅ Лист очищен');
    } catch (e) {
      console.log('⚠️ Не удалось очистить лист:', e);
    }
    
    console.log('Этап 3: Запись данных...');
    const range = sheet.getRange(1, 1, numRows, numCols);
    range.setValues(tableData);
    console.log('✅ Данные записаны');
    
    console.log('Этап 4: Применение форматирования...');
    applyEnhancedFormatting(sheet, numRows, numCols, formatData, appData);
    console.log('✅ Форматирование применено');
    
    console.log('Этап 5: Создание группировки...');
    createRowGrouping(sheet, formatData, appData);
    console.log('✅ Группировка создана');
    
    console.log('Этап 6: Финальные настройки...');
    sheet.setFrozenRows(1);
    sheet.hideColumns(1);
    console.log('✅ Финальные настройки применены');
    
  } catch (e) {
    console.error('❌ Ошибка при записи таблицы:', e);
    throw e;
  }
  
  console.log('=== БЕЗОПАСНАЯ ЗАПИСЬ ЗАВЕРШЕНА ===');
}

function applyEnhancedFormatting(sheet, numRows, numCols, formatData, appData) {
  console.log('Применение форматирования...');
  
  console.log('  Заголовки...');
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground(COLORS.HEADER.background)
    .setFontColor(COLORS.HEADER.fontColor)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setWrap(true);

  console.log('  Ширина колонок...');
  const columnWidths = TABLE_CONFIG.COLUMN_WIDTHS;
  columnWidths.forEach(col => {
    try {
      sheet.setColumnWidth(col.c, col.w);
    } catch (e) {
      console.log(`Ошибка установки ширины колонки ${col.c}:`, e);
    }
  });

  if (numRows > 1) {
    console.log('  Общее выравнивание...');
    try {
      const allDataRange = sheet.getRange(2, 1, numRows - 1, numCols);
      allDataRange.setVerticalAlignment('middle');
      
      const commentsRange = sheet.getRange(2, numCols, numRows - 1, 1);
      commentsRange.setWrap(true).setHorizontalAlignment('left');
      
      const growthStatusRange = sheet.getRange(2, numCols - 1, numRows - 1, 1);
      growthStatusRange.setWrap(true).setHorizontalAlignment('left');
    } catch (e) {
      console.log('Ошибка общего выравнивания:', e);
    }
  }

  console.log('  Форматирование строк по типам...');
  const rowsByType = {
    app: [],
    week: [],
    sourceApp: [],
    campaign: []
  };
  
  formatData.forEach(item => {
    if (item.type === 'APP') rowsByType.app.push(item.row);
    if (item.type === 'WEEK') rowsByType.week.push(item.row);
    if (item.type === 'SOURCE_APP') rowsByType.sourceApp.push(item.row);
    if (item.type === 'CAMPAIGN') rowsByType.campaign.push(item.row);
  });

  console.log(`    APP строк: ${rowsByType.app.length}`);
  console.log(`    WEEK строк: ${rowsByType.week.length}`);
  console.log(`    SOURCE_APP строк: ${rowsByType.sourceApp.length}`);
  console.log(`    CAMPAIGN строк: ${rowsByType.campaign.length}`);

  try {
    rowsByType.app.forEach(r => {
      sheet.getRange(r, 1, 1, numCols)
           .setBackground(COLORS.APP_ROW.background)
           .setFontColor(COLORS.APP_ROW.fontColor)
           .setFontWeight('bold')
           .setFontSize(10);
    });

    rowsByType.week.forEach(r => {
      sheet.getRange(r, 1, 1, numCols)
           .setBackground(COLORS.WEEK_ROW.background)
           .setFontSize(10);
    });

    rowsByType.sourceApp.forEach(r => {
      sheet.getRange(r, 1, 1, numCols)
           .setBackground(COLORS.SOURCE_APP_ROW.background)
           .setFontSize(9);
    });

    rowsByType.campaign.forEach(r => {
      sheet.getRange(r, 1, 1, numCols)
           .setBackground(COLORS.CAMPAIGN_ROW.background)
           .setFontSize(9);
    });
  } catch (e) {
    console.log('Ошибка форматирования строк:', e);
  }

  if (numRows > 1) {
    console.log('  Числовые форматы...');
    try {
      sheet.getRange(2, 5, numRows - 1, 1).setNumberFormat('$0.00');
      sheet.getRange(2, 8, numRows - 1, 1).setNumberFormat('$0.000');
      sheet.getRange(2, 9, numRows - 1, 1).setNumberFormat('0.00');
      sheet.getRange(2, 10, numRows - 1, 1).setNumberFormat('0.0');
      sheet.getRange(2, 13, numRows - 1, 1).setNumberFormat('$0.000');
      sheet.getRange(2, 16, numRows - 1, 1).setNumberFormat('$0.00');
    } catch (e) {
      console.log('Ошибка числовых форматов:', e);
    }
  }

  console.log('  Условное форматирование...');
  try {
    applyConditionalFormatting(sheet, numRows, appData);
  } catch (e) {
    console.log('Ошибка условного форматирования:', e);
  }
  
  console.log('✅ Форматирование завершено');
}

function applyConditionalFormatting(sheet, numRows, appData) {
  if (numRows <= 1) return;
  
  const rules = [];
  
  console.log('    Форматирование WoW изменений...');
  try {
    const spendWoWRange = sheet.getRange(2, 6, numRows - 1, 1);
    const profitWoWRange = sheet.getRange(2, 17, numRows - 1, 1);
    
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(NOT(ISBLANK(F2)), ISNUMBER(VALUE(SUBSTITUTE(F2,"%",""))), VALUE(SUBSTITUTE(F2,"%","")) > 0)')
        .setBackground(COLORS.POSITIVE.background)
        .setFontColor(COLORS.POSITIVE.fontColor)
        .setRanges([spendWoWRange])
        .build()
    );
    
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(NOT(ISBLANK(F2)), ISNUMBER(VALUE(SUBSTITUTE(F2,"%",""))), VALUE(SUBSTITUTE(F2,"%","")) < 0)')
        .setBackground(COLORS.NEGATIVE.background)
        .setFontColor(COLORS.NEGATIVE.fontColor)
        .setRanges([spendWoWRange])
        .build()
    );
    
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(NOT(ISBLANK(Q2)), ISNUMBER(VALUE(SUBSTITUTE(Q2,"%",""))), VALUE(SUBSTITUTE(Q2,"%","")) > 0)')
        .setBackground(COLORS.POSITIVE.background)
        .setFontColor(COLORS.POSITIVE.fontColor)
        .setRanges([profitWoWRange])
        .build()
    );
    
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(NOT(ISBLANK(Q2)), ISNUMBER(VALUE(SUBSTITUTE(Q2,"%",""))), VALUE(SUBSTITUTE(Q2,"%","")) < 0)')
        .setBackground(COLORS.NEGATIVE.background)
        .setFontColor(COLORS.NEGATIVE.fontColor)
        .setRanges([profitWoWRange])
        .build()
    );

    console.log('    Форматирование eROAS D730...');
    const eroasRange = sheet.getRange(2, 15, numRows - 1, 1);
    
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(NOT(ISBLANK(O2)), ISNUMBER(VALUE(SUBSTITUTE(O2,"%",""))), VALUE(SUBSTITUTE(O2,"%","")) >= 250)')
        .setBackground('#d4edda')
        .setFontColor('#155724')
        .setRanges([eroasRange])
        .build()
    );
    
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(NOT(ISBLANK(O2)), ISNUMBER(VALUE(SUBSTITUTE(O2,"%",""))), VALUE(SUBSTITUTE(O2,"%","")) >= 150, VALUE(SUBSTITUTE(O2,"%","")) < 250)')
        .setBackground('#d1f2eb')
        .setFontColor('#0c5460')
        .setRanges([eroasRange])
        .build()
    );
    
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(NOT(ISBLANK(O2)), ISNUMBER(VALUE(SUBSTITUTE(O2,"%",""))), VALUE(SUBSTITUTE(O2,"%","")) >= 140, VALUE(SUBSTITUTE(O2,"%","")) < 150)')
        .setBackground('#fff3cd')
        .setFontColor('#856404')
        .setRanges([eroasRange])
        .build()
    );
    
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(NOT(ISBLANK(O2)), ISNUMBER(VALUE(SUBSTITUTE(O2,"%",""))), VALUE(SUBSTITUTE(O2,"%","")) < 140)')
        .setBackground('#f8d7da')
        .setFontColor('#721c24')
        .setRanges([eroasRange])
        .build()
    );

    console.log('    Форматирование Growth Status...');
    const growthRange = sheet.getRange(2, 18, numRows - 1, 1);
    
    const statusFormats = [
      { text: '🟢 Healthy Growth', bg: '#d4edda', color: '#155724' },
      { text: '🟢 Efficiency Improvement', bg: '#d1f2eb', color: '#0c5460' },
      { text: '🔴 Inefficient Growth', bg: '#f8d7da', color: '#721c24' },
      { text: '🟠 Declining Efficiency', bg: '#fff3cd', color: '#856404' },
      { text: '🔵 Scaling Down', bg: '#cce7ff', color: '#004085' },
      { text: '🟡 Moderate Growth', bg: '#fff3cd', color: '#856404' },
      { text: '🟡 Moderate Decline', bg: '#fff3cd', color: '#856404' },
      { text: '⚪ Stable', bg: '#f5f5f5', color: '#616161' }
    ];
    
    statusFormats.forEach(format => {
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains(format.text)
          .setBackground(format.bg)
          .setFontColor(format.color)
          .setRanges([growthRange])
          .build()
      );
    });
    
    console.log(`    Применяем ${rules.length} правил условного форматирования...`);
    sheet.setConditionalFormatRules(rules);
    console.log('    ✅ Условное форматирование применено');
  } catch (e) {
    console.log('Ошибка в условном форматировании:', e);
  }
}

function createRowGrouping(sheet, formatData, appData) {
  console.log('Создание группировки строк...');
  
  try {
    let currentRow = 2;
    const appKeys = Object.keys(appData).sort((a, b) => appData[a].appName.localeCompare(appData[b].appName));
    
    console.log(`Создаем группы для ${appKeys.length} приложений`);
    
    appKeys.forEach((appKey, appIndex) => {
      const app = appData[appKey];
      const appStartRow = currentRow;
      currentRow++;
      
      const weekKeys = Object.keys(app.weeks).sort();
      let appContentRows = 0;
      
      weekKeys.forEach(weekKey => {
        const week = app.weeks[weekKey];
        const weekStartRow = currentRow;
        currentRow++;
        let weekContentRows = 0;
        
        if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
          const sourceAppKeys = Object.keys(week.sourceApps).sort((a, b) => {
            const spendA = week.sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
            const spendB = week.sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
            return spendB - spendA;
          });
          
          sourceAppKeys.forEach(sourceAppKey => {
            const sourceApp = week.sourceApps[sourceAppKey];
            const sourceAppStartRow = currentRow;
            currentRow++;
            
            const campaignCount = sourceApp.campaigns.length;
            currentRow += campaignCount;
            weekContentRows += 1 + campaignCount;
            
            if (campaignCount > 0) {
              try {
                const campaignRange = sheet.getRange(sourceAppStartRow + 1, 1, campaignCount, 1);
                campaignRange.shiftRowGroupDepth(1);
                campaignRange.collapseGroups();
                console.log(`      Группа кампаний: строки ${sourceAppStartRow + 1}-${sourceAppStartRow + campaignCount}`);
              } catch (e) {
                console.log(`      Ошибка группировки кампаний: ${e}`);
              }
            }
          });
          
        } else if (CURRENT_PROJECT !== 'OVERALL') {
          const campaignCount = week.campaigns?.length || 0;
          currentRow += campaignCount;
          weekContentRows = campaignCount;
          
          if (campaignCount > 0) {
            try {
              const campaignRange = sheet.getRange(weekStartRow + 1, 1, campaignCount, 1);
              campaignRange.shiftRowGroupDepth(1);
              campaignRange.collapseGroups();
              console.log(`    Группа кампаний недели: строки ${weekStartRow + 1}-${weekStartRow + campaignCount}`);
            } catch (e) {
              console.log(`    Ошибка группировки кампаний недели: ${e}`);
            }
          }
        }
        
        appContentRows += 1 + weekContentRows;
        
        if (weekContentRows > 0) {
          try {
            const weekRange = sheet.getRange(weekStartRow + 1, 1, weekContentRows, 1);
            weekRange.shiftRowGroupDepth(1);
            weekRange.collapseGroups();
            console.log(`    Группа недели: строки ${weekStartRow + 1}-${weekStartRow + weekContentRows}`);
          } catch (e) {
            console.log(`    Ошибка группировки недели: ${e}`);
          }
        }
      });
      
      if (appContentRows > 0) {
        try {
          const appRange = sheet.getRange(appStartRow + 1, 1, appContentRows, 1);
          appRange.shiftRowGroupDepth(1);
          appRange.collapseGroups();
          console.log(`  Группа приложения ${appIndex + 1}: строки ${appStartRow + 1}-${appStartRow + appContentRows}`);
        } catch (e) {
          console.log(`  Ошибка группировки приложения: ${e}`);
        }
      }
    });
    
    console.log('✅ Группировка завершена успешно');
    
  } catch (e) {
    console.error('❌ Ошибка создания группировки:', e);
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