// ========== ДЕКЛАРАТИВНАЯ КОНФИГУРАЦИЯ ФОРМАТИРОВАНИЯ ==========
// Формат: [fontSize, background, fontWeight, fontColor]
const FORMAT_RULES = {
  INCENT_TRAFFIC: {
    types: {
      APP: [9, '#ffffff'],
      NETWORK: [10, '#d1e7fe', 'bold', '#000000'],
      COUNTRY: [10, '#f0f8ff'],
      CAMPAIGN: [10, '#ffffff'],
      WEEK: [9, '#ffffff']
    },
    hideColumns: [], // Убрали скрытие GEO колонки
    remapping: null
  },
  
  APPLOVIN_TEST: {
  types: {
    APP: [10, '#d1e7fe', 'bold', '#000000'],
    CAMPAIGN: [10, '#e8f0fe'],  // светло-голубой фон для кампаний
    WEEK: [9, '#ffffff'],       // белый фон для недель
    COUNTRY: [9, '#ffffff']      // белый фон для стран
  },
  hideColumns: [4], // GEO
  remapping: null  // Убираем remapping, он только путает
},
  
  OVERALL: {
    types: {
      APP: [10, '#d1e7fe', 'bold', '#000000'],
      WEEK: [10, '#e8f0fe'],
      NETWORK: [9, '#ffffff']
    },
    hideColumns: [],
    remapping: null
  },
  
  TRICKY: {
    types: {
      APP: [10, '#d1e7fe', 'bold', '#000000'],
      WEEK: [10, '#e8f0fe'],
      SOURCE_APP: [10, '#f0f8ff'],
      CAMPAIGN: [9, '#ffffff']
    },
    hideColumns: [],
    remapping: null,
    hyperlinkFormatting: true
  },
  
  MOLOCO: {
    types: {
      APP: [10, '#d1e7fe', 'bold', '#000000'],
      WEEK: [10, '#e8f0fe'],
      CAMPAIGN: [9, '#ffffff']
    },
    hideColumns: [],
    remapping: null
  },
  
  REGULAR: {
    types: {
      APP: [10, '#d1e7fe', 'bold', '#000000'],
      WEEK: [10, '#e8f0fe'],
      CAMPAIGN: [9, '#ffffff']
    },
    hideColumns: [],
    remapping: null,
    hyperlinkFormatting: true
  },
  
  GOOGLE_ADS: {
    types: {
      APP: [10, '#d1e7fe', 'bold', '#000000'],
      WEEK: [10, '#e8f0fe'],
      CAMPAIGN: [9, '#ffffff']
    },
    hideColumns: [],
    remapping: null
  },
  
  APPLOVIN: {
    types: {
      APP: [10, '#d1e7fe', 'bold', '#000000'],
      WEEK: [10, '#e8f0fe'],
      CAMPAIGN: [9, '#ffffff']
    },
    hideColumns: [],
    remapping: null
  },
  
  MINTEGRAL: {
    types: {
      APP: [10, '#d1e7fe', 'bold', '#000000'],
      WEEK: [10, '#e8f0fe'],
      CAMPAIGN: [9, '#ffffff']
    },
    hideColumns: [],
    remapping: null
  },
  
  DEFAULT: {
    types: {
      APP: [10, '#d1e7fe', 'bold', '#000000'],
      WEEK: [10, '#e8f0fe'],
      SOURCE_APP: [10, '#f0f8ff'],
      CAMPAIGN: [9, '#ffffff'],
      NETWORK: [10, '#d1e7fe', 'bold', '#000000'],
      COUNTRY: [9, '#ffffff']
    },
    hideColumns: [],
    remapping: null
  }
};

// Конфигурация условного форматирования
const CONDITIONAL_CONFIG = {
  statusColors: {
    "🟢 Healthy Growth": ["#d4edda", "#155724"],
    "🟢 Efficiency Improvement": ["#d1f2eb", "#0c5460"],
    "🔴 Inefficient Growth": ["#f8d7da", "#721c24"],
    "🟠 Declining Efficiency": ["#ff9800", "white"],
    "🔵 Scaling Down": ["#cce7ff", "#004085"],
    "🔵 Scaling Down - Efficient": ["#b8e6b8", "#2d5a2d"],
    "🔵 Scaling Down - Moderate": ["#d1ecf1", "#0c5460"],
    "🔵 Scaling Down - Problematic": ["#ffcc99", "#cc5500"],
    "🟡 Moderate Growth": ["#fff3cd", "#856404"],
    "🟡 Moderate Decline - Efficiency Drop": ["#ffe0cc", "#cc6600"],
    "🟡 Moderate Decline - Spend Optimization": ["#e6f3ff", "#0066cc"],
    "🟡 Moderate Decline - Proportional": ["#f0f0f0", "#666666"],
    "🟡 Efficiency Improvement": ["#e8f5e8", "#2d5a2d"],
    "🟡 Minimal Growth": ["#fff8e1", "#f57f17"],
    "🟡 Moderate Decline": ["#fff3cd", "#856404"],
    "⚪ Stable": ["#f5f5f5", "#616161"],
    "First Week": ["#e0e0e0", "#757575"]
  },
  columns: {
    spend: COLUMN_CONFIG.COLUMNS.SPEND_WOW,   // 6
    eROAS: COLUMN_CONFIG.COLUMNS.EROAS_730,   // 20 (обновлено)
    eProfit: COLUMN_CONFIG.COLUMNS.EPROFIT_730, // 21 (обновлено)
    profit: COLUMN_CONFIG.COLUMNS.EPROFIT_WOW,  // 22 (обновлено)
    growth: COLUMN_CONFIG.COLUMNS.GROWTH_STATUS  // 23 (обновлено)
  },
  numberFormats: [
    { col: COLUMN_CONFIG.COLUMNS.CPI, format: '$0.0' },           // 8
    { col: COLUMN_CONFIG.COLUMNS.ROAS_COEF_3_1, format: '0.0' },  // NEW
    { col: COLUMN_CONFIG.COLUMNS.ROAS_COEF_7_3, format: '0.0' },  // NEW
    { col: COLUMN_CONFIG.COLUMNS.ROAS_COEF_14_7, format: '0.0' }, // NEW
    { col: COLUMN_CONFIG.COLUMNS.ROAS_COEF_30_7, format: '0.0' }, // NEW
    { col: COLUMN_CONFIG.COLUMNS.IPM, format: '0.0' },            // 14
    { col: COLUMN_CONFIG.COLUMNS.RR_COEF_3_1, format: '0.0' },    // NEW
    { col: COLUMN_CONFIG.COLUMNS.RR_COEF_7_3, format: '0.0' },    // NEW
    { col: COLUMN_CONFIG.COLUMNS.EARPU, format: '$0.0' },         // 18
    { col: COLUMN_CONFIG.COLUMNS.EPROFIT_730, format: '$0.0' }    // 21
  ],
  // Это поле больше не используется - заменено на COLUMN_CONFIG.HIDDEN_COLUMNS
};

// ========== ПУБЛИЧНЫЕ ФУНКЦИИ (для совместимости) ==========
function createEnhancedPivotTable(appData) { 
  createUnifiedPivotTable(appData); 
}

function createOverallPivotTable(appData) { 
  createUnifiedPivotTable(appData); 
}

function createIncentTrafficPivotTable(networkData) { 
  createUnifiedPivotTable(networkData); 
}

function createProjectPivotTable(projectName, appData) {
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  try {
    createUnifiedPivotTable(appData);
  } finally {
    setCurrentProject(originalProject);
  }
}

function createUnifiedPivotTable(data) {
  const startTime = Date.now();
  
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(config.SHEET_NAME);
  else sheet.clear();

  if (!data || Object.keys(data).length === 0) {
    const headers = getUnifiedHeaders();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const initialMetricsCache = new InitialMetricsCache();
  initialMetricsCache.recordInitialValuesFromData(data);

  const wow = CURRENT_PROJECT === 'INCENT_TRAFFIC' ? 
    calculateIncentTrafficWoWMetrics(data) : 
    calculateWoWMetrics(data);
  
  const headers = getUnifiedHeaders();
  const tableData = [headers];
  const formatData = [];
  buildUnifiedTable(data, tableData, formatData, wow, initialMetricsCache);

  const range = sheet.getRange(1, 1, tableData.length, headers.length);
  range.setValues(tableData);
  
  applyOptimizedFormatting(sheet, tableData.length, headers.length, formatData, data);
  
  createUnifiedRowGrouping(sheet, tableData, data);
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  console.log(`Pivot table created in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
}

// ========== ГЛАВНАЯ ФУНКЦИЯ ФОРМАТИРОВАНИЯ ==========
function applyOptimizedFormatting(sheet, numRows, numCols, formatData, appData) {
  const startTime = Date.now();
  
  try {
    const spreadsheetId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();
    
    // 1. Базовое форматирование
    applyBaseFormatting(sheet, numRows, numCols);
    
    // 2. Форматирование по типам строк
    applyTypeFormatting(sheet, formatData, numCols);
    
    // 3. Условное форматирование
    applyConditionalFormats(sheet, sheetId, numRows, appData, spreadsheetId);
    
    // 4. eROAS/eProfit форматирование
    applyArrowFormatting(sheet, sheetId, numRows, spreadsheetId);
    
    // 5. Скрытие колонок
    applyColumnHiding(sheet);
    
    console.log(`Formatting completed in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
    
  } catch (e) {
    console.error('Error in applyOptimizedFormatting:', e);
    throw e;
  }
}

// ========== БАЗОВОЕ ФОРМАТИРОВАНИЕ ==========
function applyBaseFormatting(sheet, numRows, numCols) {
  // Ширина колонок
  TABLE_CONFIG.COLUMN_WIDTHS.forEach(col => {
    sheet.setColumnWidth(col.c, col.w);
  });
  
  // Заголовок
  sheet.getRange(1, 1, 1, numCols)
    .setBackground('#4285f4')
    .setFontColor('white')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setWrap(true);
  
  if (numRows <= 1) return;
  
  // Базовое форматирование данных
  sheet.getRange(2, 1, numRows - 1, numCols).setVerticalAlignment('middle');
  
  // Специальные колонки
  sheet.getRange(2, COLUMN_CONFIG.COLUMNS.ROAS_COMBINED, numRows - 1, 1).setWrap(true).setHorizontalAlignment('center'); // ROAS
  sheet.getRange(2, COLUMN_CONFIG.COLUMNS.RR_COMBINED, numRows - 1, 1).setWrap(true).setHorizontalAlignment('center'); // RR
  sheet.getRange(2, COLUMN_CONFIG.COLUMNS.COMMENTS, numRows - 1, 1).setWrap(true).setHorizontalAlignment('left'); // Comments
  sheet.getRange(2, COLUMN_CONFIG.COLUMNS.GROWTH_STATUS, numRows - 1, 1).setWrap(true).setHorizontalAlignment('left'); // Growth Status
  sheet.getRange(2, COLUMN_CONFIG.COLUMNS.EROAS_730, numRows - 1, 1).setHorizontalAlignment('right'); // eROAS
  sheet.getRange(2, COLUMN_CONFIG.COLUMNS.EPROFIT_730, numRows - 1, 1).setHorizontalAlignment('right'); // eProfit
  
  // Числовое форматирование
  CONDITIONAL_CONFIG.numberFormats.forEach(({ col, format }) => {
    sheet.getRange(2, col, numRows - 1, 1).setNumberFormat(format);
  });
}

// ========== ФОРМАТИРОВАНИЕ ПО ТИПАМ ==========
function applyTypeFormatting(sheet, formatData, numCols) {
  const rules = FORMAT_RULES[CURRENT_PROJECT] || FORMAT_RULES.DEFAULT;
  const typeMap = {};
  const hyperlinkRows = [];
  
  // Группировка строк по типам с учетом ремаппинга
  formatData.forEach(item => {
    if (item.type === 'HYPERLINK') {
      hyperlinkRows.push(item.row);
      return;
    }
    
    let type = item.type;
    if (rules.remapping && rules.remapping[type]) {
      type = rules.remapping[type];
    } else {
      type = type.toLowerCase();
    }
    
    if (!typeMap[type]) typeMap[type] = [];
    typeMap[type].push(item.row);
  });
  
  // Применение форматирования для каждого типа
  Object.entries(typeMap).forEach(([type, rows]) => {
    if (rows.length === 0) return;
    
    const typeUpper = type.toUpperCase();
    const config = rules.types[typeUpper];
    if (!config) return;
    
    const [fontSize, background, fontWeight, fontColor] = config;
    
    createOptimizedRanges(sheet, rows, numCols).forEach(range => {
      if (fontSize) range.setFontSize(fontSize);
      if (background) range.setBackground(background);
      if (fontWeight) range.setFontWeight(fontWeight);
      if (fontColor) range.setFontColor(fontColor);
    });
  });
  
  // Обработка гиперссылок для TRICKY и REGULAR
  if (rules.hyperlinkFormatting && hyperlinkRows.length > 0) {
    hyperlinkRows.filter(row => row >= 2).forEach(row => {
      try {
        sheet.getRange(row, 2, 1, 1)
          .setFontColor('#000000')
          .setFontLine('none');
      } catch (e) {
        console.error(`Error formatting hyperlink row ${row}:`, e);
      }
    });
  }
}

// ========== УСЛОВНОЕ ФОРМАТИРОВАНИЕ ==========
function applyConditionalFormats(sheet, sheetId, numRows, appData, spreadsheetId) {
  if (numRows <= 1) return;
  
  const requests = [];
  const cols = CONDITIONAL_CONFIG.columns;
  
  // 1. Spend WoW (отрицательные/положительные) - используем A1 формат с правильными индексами
  const spendCol = cols.spend; // 6
  const spendLetter = columnIndexToLetter(spendCol); // F
  requests.push(
    createFormatRule(sheetId, spendCol, 1, numRows, 
      `=AND(NOT(ISBLANK($${spendLetter}2)), LEFT($${spendLetter}2,1)="-")`, 
      '#f8d7da', '#721c24', 0),
    createFormatRule(sheetId, spendCol, 1, numRows,
      `=AND(NOT(ISBLANK($${spendLetter}2)), $${spendLetter}2<>"", LEFT($${spendLetter}2,1)<>"-")`,
      '#d1f2eb', '#0c5460', 1)
  );
  
  // 2. eROAS правила (с учетом таргетов)
  const targetGroups = groupRowsByEROASTarget(sheet.getDataRange().getValues());
  let ruleIndex = requests.length;
  
  targetGroups.forEach((rows, targetEROAS) => {
    const ranges = rows.map(row => ({
      sheetId: sheetId,
      startRowIndex: row - 1,
      endRowIndex: row,
      startColumnIndex: cols.eROAS - 1,
      endColumnIndex: cols.eROAS
    }));
    
    // Формулы для eROAS
    const formulas = [
      [`>=`, targetEROAS, '#d1f2eb', '#0c5460'],
      [`>=`, 120, `<`, targetEROAS, '#fff3cd', '#856404'],
      [`<`, 120, '#f8d7da', '#721c24']
    ];
    
    formulas.forEach(formula => {
      requests.push(createEROASRule(ranges, formula, ruleIndex++));
    });
  });
  
  // 3. Profit WoW - используем A1 формат с правильными индексами
  const profitCol = cols.profit; // 16 (колонка P)
  const profitLetter = columnIndexToLetter(profitCol); // P
  requests.push(
    createFormatRule(sheetId, profitCol, 1, numRows,
      `=AND(ISNUMBER($${profitLetter}2), $${profitLetter}2>0)`,
      '#d1f2eb', '#0c5460', ruleIndex++),
    createFormatRule(sheetId, profitCol, 1, numRows,
      `=AND(ISNUMBER($${profitLetter}2), $${profitLetter}2<0)`,
      '#f8d7da', '#721c24', ruleIndex++)
  );
  
  // 4. Growth Status
  Object.entries(CONDITIONAL_CONFIG.statusColors).forEach(([status, [bg, text]]) => {
    requests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: [{
            sheetId: sheetId,
            startRowIndex: 1,
            endRowIndex: numRows,
            startColumnIndex: cols.growth - 1,
            endColumnIndex: cols.growth
          }],
          booleanRule: {
            condition: {
              type: 'TEXT_CONTAINS',
              values: [{ userEnteredValue: status }]
            },
            format: {
              backgroundColor: hexToRgb(bg),
              textFormat: { foregroundColor: hexToRgb(text) }
            }
          }
        },
        index: ruleIndex++
      }
    });
  });
  
  // Применяем все правила одним батчем
  if (requests.length > 0) {
    Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);
  }
}

// ========== ФОРМАТИРОВАНИЕ СТРЕЛОК В eROAS/eProfit ==========
function applyArrowFormatting(sheet, sheetId, numRows, spreadsheetId) {
  if (numRows <= 1) return;
  
  const rules = FORMAT_RULES[CURRENT_PROJECT] || FORMAT_RULES.DEFAULT;
  
  // Используем максимальный индекс колонки для чтения всех нужных данных
  const maxCol = Math.max(
    COLUMN_CONFIG.COLUMNS.EROAS_730,
    COLUMN_CONFIG.COLUMNS.EPROFIT_730
  );
  
  const data = sheet.getRange(2, 1, numRows - 1, maxCol).getValues();
  const requests = [];
  
  data.forEach((row, index) => {
    const level = row[0];
    // Используем правильные индексы из конфига (минус 1 для 0-based массива)
    const eroasValue = row[COLUMN_CONFIG.COLUMNS.EROAS_730 - 1];
    const eprofitValue = row[COLUMN_CONFIG.COLUMNS.EPROFIT_730 - 1];
    const rowIndex = index + 1;
    
    // Получаем размер шрифта из конфигурации
    const config = rules.types[level];
    const baseFontSize = config ? config[0] : 10;
    const smallerFontSize = baseFontSize - 1;
    
    // Обрабатываем стрелки в eROAS
    if (eroasValue && typeof eroasValue === 'string' && eroasValue.includes('→')) {
      requests.push(createArrowFormat(sheetId, rowIndex, COLUMN_CONFIG.COLUMNS.EROAS_730 - 1, 
        eroasValue, smallerFontSize, baseFontSize));
    }
    
    // Обрабатываем стрелки в eProfit
    if (eprofitValue && typeof eprofitValue === 'string' && eprofitValue.includes('→')) {
      requests.push(createArrowFormat(sheetId, rowIndex, COLUMN_CONFIG.COLUMNS.EPROFIT_730 - 1,
        eprofitValue, smallerFontSize, baseFontSize));
    }
  });
  
  // Применяем батчами
  if (requests.length > 0) {
    const batchSize = 500;
    for (let i = 0; i < requests.length; i += batchSize) {
      Sheets.Spreadsheets.batchUpdate({
        requests: requests.slice(i, i + batchSize)
      }, spreadsheetId);
      
      if (i + batchSize < requests.length) {
        Utilities.sleep(50);
      }
    }
  }
}

// ========== СКРЫТИЕ КОЛОНОК ==========
function applyColumnHiding(sheet) {
  // Стандартные скрытые колонки
  COLUMN_CONFIG.HIDDEN_COLUMNS.forEach(col => {
    if (col === 13 || col === 14) {
      sheet.hideColumns(col, 1);
    } else {
      sheet.hideColumns(col);
    }
  });
  
  // Проектно-специфичные скрытые колонки
  const rules = FORMAT_RULES[CURRENT_PROJECT] || FORMAT_RULES.DEFAULT;
  if (rules.hideColumns) {
    rules.hideColumns.forEach(col => sheet.hideColumns(col));
  }
}

// ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========

// Преобразование индекса колонки в букву (1 -> A, 14 -> N, и т.д.)
function columnIndexToLetter(index) {
  let letter = '';
  while (index > 0) {
    const mod = (index - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

function createOptimizedRanges(sheet, rowNumbers, numCols) {
  if (rowNumbers.length === 0) return [];
  
  const ranges = [];
  const sortedRows = [...rowNumbers].sort((a, b) => a - b);
  
  let start = sortedRows[0];
  let end = start;
  
  for (let i = 1; i < sortedRows.length; i++) {
    if (sortedRows[i] === end + 1) {
      end = sortedRows[i];
    } else {
      ranges.push(sheet.getRange(start, 1, end - start + 1, numCols));
      start = sortedRows[i];
      end = start;
    }
  }
  
  ranges.push(sheet.getRange(start, 1, end - start + 1, numCols));
  return ranges;
}

function hexToRgb(hex) {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    red: parseInt(result[1], 16) / 255,
    green: parseInt(result[2], 16) / 255,
    blue: parseInt(result[3], 16) / 255
  } : { red: 1, green: 1, blue: 1 };
}

function createFormatRule(sheetId, column, startRow, endRow, formula, bgColor, textColor, index) {
  // Используем формулу как есть - она уже в правильном A1 формате
  return {
    addConditionalFormatRule: {
      rule: {
        ranges: [{
          sheetId: sheetId,
          startRowIndex: startRow,
          endRowIndex: endRow,
          startColumnIndex: column - 1,
          endColumnIndex: column
        }],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: formula }]
          },
          format: {
            backgroundColor: hexToRgb(bgColor),
            textFormat: { foregroundColor: hexToRgb(textColor) }
          }
        }
      },
      index: index
    }
  };
}

function createEROASRule(ranges, formula, index) {
  const eroasCol = COLUMN_CONFIG.COLUMNS.EROAS_730; // 14 (колонка N)
  const eroasLetter = columnIndexToLetter(eroasCol); // N
  let formulaStr;
  
  if (formula.length === 4) {
    // Зеленый (>= target)
    formulaStr = `=AND(NOT(ISBLANK(INDIRECT("${eroasLetter}" & ROW()))), IF(ISERROR(SEARCH("→", INDIRECT("${eroasLetter}" & ROW()))), VALUE(SUBSTITUTE(INDIRECT("${eroasLetter}" & ROW()), "%", "")), VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(INDIRECT("${eroasLetter}" & ROW()), "→", REPT(" ", 100)), 100)), "%", ""))) ${formula[0]} ${formula[1]})`;
  } else if (formula.length === 6) {
    // Желтый (между 120 и target)
    formulaStr = `=AND(NOT(ISBLANK(INDIRECT("${eroasLetter}" & ROW()))), IF(ISERROR(SEARCH("→", INDIRECT("${eroasLetter}" & ROW()))), VALUE(SUBSTITUTE(INDIRECT("${eroasLetter}" & ROW()), "%", "")), VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(INDIRECT("${eroasLetter}" & ROW()), "→", REPT(" ", 100)), 100)), "%", ""))) ${formula[0]} ${formula[1]}, IF(ISERROR(SEARCH("→", INDIRECT("${eroasLetter}" & ROW()))), VALUE(SUBSTITUTE(INDIRECT("${eroasLetter}" & ROW()), "%", "")), VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(INDIRECT("${eroasLetter}" & ROW()), "→", REPT(" ", 100)), 100)), "%", ""))) ${formula[2]} ${formula[3]})`;
  } else {
    // Красный (< 120)
    formulaStr = `=AND(NOT(ISBLANK(INDIRECT("${eroasLetter}" & ROW()))), IF(ISERROR(SEARCH("→", INDIRECT("${eroasLetter}" & ROW()))), VALUE(SUBSTITUTE(INDIRECT("${eroasLetter}" & ROW()), "%", "")), VALUE(SUBSTITUTE(TRIM(RIGHT(SUBSTITUTE(INDIRECT("${eroasLetter}" & ROW()), "→", REPT(" ", 100)), 100)), "%", ""))) ${formula[0]} ${formula[1]})`;
  }
  
  const colors = formula[formula.length - 2];
  const textColor = formula[formula.length - 1];
  
  return {
    addConditionalFormatRule: {
      rule: {
        ranges: ranges,
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: formulaStr }]
          },
          format: {
            backgroundColor: hexToRgb(colors),
            textFormat: { foregroundColor: hexToRgb(textColor) }
          }
        }
      },
      index: index
    }
  };
}

function createArrowFormat(sheetId, rowIndex, columnIndex, value, smallFont, largeFont) {
  const arrowIndex = value.indexOf('→');
  
  return {
    updateCells: {
      range: {
        sheetId: sheetId,
        startRowIndex: rowIndex,
        endRowIndex: rowIndex + 1,
        startColumnIndex: columnIndex,
        endColumnIndex: columnIndex + 1
      },
      rows: [{
        values: [{
          userEnteredValue: { stringValue: value },
          textFormatRuns: [
            {
              startIndex: 0,
              format: {
                foregroundColor: { red: 0.5, green: 0.5, blue: 0.5 },
                fontSize: smallFont
              }
            },
            {
              startIndex: arrowIndex,
              format: { fontSize: largeFont }
            }
          ]
        }]
      }],
      fields: 'userEnteredValue,textFormatRuns'
    }
  };
}

function groupRowsByEROASTarget(data) {
  const targetGroups = new Map();
  
  for (let i = 1; i < data.length; i++) {
    const level = data[i][0];
    let appName = '';
    let targetEROAS = 150;
    
    if (level === 'APP') {
      appName = data[i][1];
      targetEROAS = getTargetEROAS(CURRENT_PROJECT, appName);
    } else {
      // Ищем родительское приложение
      for (let j = i - 1; j >= 1; j--) {
        if (data[j][0] === 'APP') {
          appName = data[j][1];
          targetEROAS = getTargetEROAS(CURRENT_PROJECT, appName);
          break;
        }
      }
    }
    
    if (!targetGroups.has(targetEROAS)) {
      targetGroups.set(targetEROAS, []);
    }
    targetGroups.get(targetEROAS).push(i + 1);
  }
  
  return targetGroups;
}