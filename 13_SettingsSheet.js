/**
 * Settings Sheet Management - ОБНОВЛЕНО: таргеты eROAS D730 вместо D365
 */

var SETTINGS_SHEET_NAME = 'Settings';
var SETTINGS_CACHE = null;
var SETTINGS_CACHE_TIME = null;

function getOrCreateSettingsSheet() {
  const spreadsheet = SpreadsheetApp.openById(MAIN_SHEET_ID);
  let sheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SETTINGS_SHEET_NAME);
    createSettingsLayout(sheet);
    populateDefaultSettings(sheet);
  }
  
  return sheet;
}

function createSettingsLayout(sheet) {
  sheet.clear();
  
  // Заголовок
  sheet.getRange('A1:F1').merge().setValue('⚙️ CAMPAIGN REPORT SETTINGS');
  sheet.getRange('A1:F1').setBackground('#1c4587').setFontColor('white').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  
  // API Settings
  sheet.getRange('A3').setValue('🔐 API SETTINGS').setBackground('#4285f4').setFontColor('white').setFontWeight('bold');
  sheet.getRange('A4').setValue('Bearer Token:');
  sheet.getRange('B4').setValue('[ENTER_YOUR_TOKEN_HERE]');
  sheet.getRange('B4:F4').merge();
  sheet.getRange('A4:A4').setBackground('#e8f0fe');
  
  // Target eROAS D730
  sheet.getRange('A6').setValue('🎯 TARGET eROAS D730 (%)').setBackground('#34a853').setFontColor('white').setFontWeight('bold');
  const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
  projects.forEach((proj, i) => {
    const row = 7 + i;
    sheet.getRange(`A${row}`).setValue(proj + ':');
    // ОБНОВЛЕНО: новые дефолтные значения для eROAS D730
    const defaultValue = proj === 'TRICKY' ? 250 : 150;
    sheet.getRange(`B${row}`).setValue(defaultValue);
    sheet.getRange(`A${row}:A${row}`).setBackground('#e8f5e8');
  });
  
  // Automation
  sheet.getRange('A16').setValue('🤖 AUTOMATION').setBackground('#ff9800').setFontColor('white').setFontWeight('bold');
  sheet.getRange('A17').setValue('Auto Cache Enabled:');
  sheet.getRange('B17').setValue('FALSE');
  sheet.getRange('A18').setValue('Auto Update Enabled:');
  sheet.getRange('B18').setValue('FALSE');
  sheet.getRange('A17:A18').setBackground('#fff3e0');
  
  // Advanced Growth Thresholds
  sheet.getRange('A20').setValue('📊 GROWTH THRESHOLDS (Advanced)').setBackground('#9c27b0').setFontColor('white').setFontWeight('bold');
  sheet.getRange('A21').setValue('Project');
  sheet.getRange('B21').setValue('Healthy Growth');
  sheet.getRange('C21').setValue('Efficiency Improvement');
  sheet.getRange('D21').setValue('Inefficient Growth');
  sheet.getRange('E21').setValue('Scaling Down');
  sheet.getRange('F21').setValue('Other Thresholds');
  sheet.getRange('A21:F21').setBackground('#f3e5f5').setFontWeight('bold');
  
  projects.forEach((proj, i) => {
    const row = 22 + i;
    sheet.getRange(`A${row}`).setValue(proj + ':');
    sheet.getRange(`B${row}`).setValue('spend:10,profit:5');
    sheet.getRange(`C${row}`).setValue('spendDrop:-5,profitGain:8');
    sheet.getRange(`D${row}`).setValue('profitDrop:-8');
    sheet.getRange(`E${row}`).setValue('spendDrop:-15,efficientProfit:0,moderateMin:-1,moderateMax:-10');
    sheet.getRange(`F${row}`).setValue('modSpend:3,modProfit:2,stable:2');
    sheet.getRange(`A${row}:A${row}`).setBackground('#fce4ec');
  });
  
  // Detailed Instructions
  sheet.getRange('A32').setValue('📖 DETAILED INSTRUCTIONS').setBackground('#607d8b').setFontColor('white').setFontWeight('bold');
  
  // API Instructions
  sheet.getRange('A34').setValue('🔐 API Settings:');
  sheet.getRange('A34').setFontWeight('bold');
  sheet.getRange('A35:F37').merge();
  sheet.getRange('A35').setValue(
    '• Bearer Token: Получите из app.appodeal.com → Settings → API\n' +
    '• Токен должен начинаться с "eyJ" и быть длиной 300+ символов\n' +
    '• Один токен работает для всех проектов'
  );
  sheet.getRange('A35:F37').setBackground('#f5f5f5').setWrap(true);
  
  // Target eROAS Instructions - ОБНОВЛЕНО для D730
  sheet.getRange('A39').setValue('🎯 Target eROAS D730:');
  sheet.getRange('A39').setFontWeight('bold');
  sheet.getRange('A40:F42').merge();
  sheet.getRange('A40').setValue(
    '• Целевые значения eROAS D730 для цветового кодирования в отчетах\n' +
    '• Зеленый: ≥ вашего значения, Желтый: 120-ваше значение, Красный: <120%\n' +
    '• Динамические таргеты: TRICKY=250%, Business apps=140%, остальные=150%\n' +
    '• Настройки влияют на условное форматирование столбца eROAS D730'
  );
  sheet.getRange('A40:F42').setBackground('#f5f5f5').setWrap(true);
  
  // Growth Thresholds Instructions
  sheet.getRange('A44').setValue('📊 Growth Thresholds (Пороги роста):');
  sheet.getRange('A44').setFontWeight('bold');
  sheet.getRange('A45:F55').merge();
  sheet.getRange('A45').setValue(
    '🟢 HEALTHY GROWTH (spend:X,profit:Y):\n' +
    '• spend:10 = спенд вырос на 10%+, profit:5 = профит вырос на 5%+\n' +
    '• Оба условия должны выполняться = 🟢 Healthy Growth\n\n' +
    
    '🟢 EFFICIENCY IMPROVEMENT (spendDrop:X,profitGain:Y):\n' +
    '• spendDrop:-5 = спенд упал на 5%+, profitGain:8 = профит вырос на 8%+\n' +
    '• Тратим меньше, зарабатываем больше = 🟢 Efficiency Improvement\n\n' +
    
    '🔴 INEFFICIENT GROWTH (profitDrop:X):\n' +
    '• profitDrop:-8 = профит упал на 8%+\n' +
    '• Критическое падение прибыли = 🔴 Inefficient Growth\n\n' +
    
    '🔵 SCALING DOWN (spendDrop:X,efficientProfit:Y,moderateMin:Z,moderateMax:W):\n' +
    '• spendDrop:-15 = спенд упал на 15%+\n' +
    '• efficientProfit:0 = если профит не упал = 🔵 Efficient\n' +
    '• moderateMin:-1, moderateMax:-10 = профит упал от 1% до 10% = 🔵 Moderate\n' +
    '• профит упал >10% = 🔵 Problematic\n\n' +
    
    '🟡 OTHER THRESHOLDS (modSpend:X,modProfit:Y,stable:Z):\n' +
    '• modSpend:3,modProfit:2 = умеренный рост (3%+ спенд, 2%+ профит)\n' +
    '• stable:2 = стабильные изменения (±2%)\n\n' +
    
    '💡 ПОЧЕМУ МИНУСЫ: Отрицательные значения = падение (spendDrop:-15 = "спенд упал на 15%")\n' +
    'Положительные значения = рост (profitGain:8 = "профит вырос на 8%")'
  );
  sheet.getRange('A45:F55').setBackground('#f5f5f5').setWrap(true);
  
  // Automation Instructions
  sheet.getRange('A57').setValue('🤖 Automation:');
  sheet.getRange('A57').setFontWeight('bold');
  sheet.getRange('A58:F60').merge();
  sheet.getRange('A58').setValue(
    '• Auto Cache: TRUE = автосохранение комментариев в 2:00 каждый день по CET\n' +
    '• Auto Update: TRUE = автообновление всех проектов в 5:00 каждый день по CET\n' +
    '• После изменения используйте "🔄 Refresh Settings" в меню для синхронизации'
  );
  sheet.getRange('A58:F60').setBackground('#f5f5f5').setWrap(true);
  
  // Unified Metrics Info - НОВОЕ
  sheet.getRange('A62').setValue('📊 Unified Metrics (New):');
  sheet.getRange('A62').setFontWeight('bold');
  sheet.getRange('A63:F65').merge();
  sheet.getRange('A63').setValue(
    '• Все проекты теперь используют унифицированный набор метрик\n' +
    '• Добавлены: IPM, RR D-1, RR D-7, eROAS D730\n' +
    '• Цветовое кодирование переведено на eROAS D730 вместо D365\n' +
    '• Фильтрация: spend > 0 применяется на уровне API (havingFilters)'
  );
  sheet.getRange('A63:F65').setBackground('#f5f5f5').setWrap(true);
  
  // Настройка ширины колонок
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 250);
  sheet.setColumnWidth(6, 180);
  
  // Валидация
  sheet.getRange('B17:B18').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TRUE', 'FALSE']).build());
}

function loadSettingsFromSheet() {
  const now = new Date().getTime();
  
  if (SETTINGS_CACHE && SETTINGS_CACHE_TIME && (now - SETTINGS_CACHE_TIME) < 30000) {
    return SETTINGS_CACHE;
  }
  
  const sheet = getOrCreateSettingsSheet();
  const data = sheet.getDataRange().getValues();
  
  const settings = {
    bearerToken: '',
    targetEROAS: {},
    automation: { autoCache: false, autoUpdate: false },
    growthThresholds: {}
  };
  
  console.log('Loading settings from sheet, total rows:', data.length);
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const label = row[0] ? row[0].toString().trim() : '';
    const value = row[1] ? row[1].toString().trim() : '';
    
    console.log(`Row ${i}: "${label}" = "${value}"`);
    
    if (label === 'Bearer Token:' && value && value !== '[ENTER_YOUR_TOKEN_HERE]') {
      settings.bearerToken = value;
      console.log('Bearer token found');
    }
    
    if (label === 'Auto Cache Enabled:') {
      settings.automation.autoCache = value.toUpperCase() === 'TRUE';
      console.log('Auto cache setting:', settings.automation.autoCache);
    }
    
    if (label === 'Auto Update Enabled:') {
      settings.automation.autoUpdate = value.toUpperCase() === 'TRUE';
      console.log('Auto update setting:', settings.automation.autoUpdate);
    }
    
    // Target eROAS D730 - проверяем что мы в правильной секции (строки 7-14)
    const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    projects.forEach(proj => {
      if (label === `${proj}:` && i >= 6 && i <= 15) { // eROAS section
        const numValue = parseInt(value);
        if (!isNaN(numValue) && numValue >= 100 && numValue <= 500) {
          settings.targetEROAS[proj] = numValue;
        } else {
          settings.targetEROAS[proj] = proj === 'TRICKY' ? 250 : 150; // ОБНОВЛЕНО: новые дефолты
        }
        console.log(`Target eROAS D730 ${proj}:`, settings.targetEROAS[proj]);
      }
    });
    
    // Advanced Growth Thresholds - проверяем что мы в правильной секции (строки 22-29)
    projects.forEach(proj => {
      if (label === `${proj}:` && i >= 21 && i <= 30) { // Advanced Growth thresholds section
        const healthyValue = row[1] ? row[1].toString() : 'spend:10,profit:5';
        const efficiencyValue = row[2] ? row[2].toString() : 'spendDrop:-5,profitGain:8';
        const inefficientValue = row[3] ? row[3].toString() : 'profitDrop:-8';
        const scalingValue = row[4] ? row[4].toString() : 'spendDrop:-15,efficientProfit:0,moderateMin:-1,moderateMax:-10';
        const otherValue = row[5] ? row[5].toString() : 'modSpend:3,modProfit:2,stable:2';
        
        settings.growthThresholds[proj] = parseAdvancedGrowthThresholds(
          healthyValue, efficiencyValue, inefficientValue, scalingValue, otherValue
        );
        console.log(`Growth thresholds ${proj}:`, settings.growthThresholds[proj]);
      }
    });
  }
  
  console.log('Final settings loaded:', settings);
  
  SETTINGS_CACHE = settings;
  SETTINGS_CACHE_TIME = now;
  
  return settings;
}

function parseAdvancedGrowthThresholds(healthyStr, efficiencyStr, inefficientStr, scalingStr, otherStr) {
  function parseCompactFormat(str, defaults = {}) {
    const result = { ...defaults };
    if (!str) return result;
    
    str.split(',').forEach(pair => {
      const [key, value] = pair.split(':');
      if (key && value !== undefined) {
        const numValue = parseFloat(value.trim());
        if (!isNaN(numValue)) {
          result[key.trim()] = numValue;
        }
      }
    });
    return result;
  }
  
  const healthy = parseCompactFormat(healthyStr, { spend: 10, profit: 5 });
  const efficiency = parseCompactFormat(efficiencyStr, { spendDrop: -5, profitGain: 8 });
  const inefficient = parseCompactFormat(inefficientStr, { profitDrop: -8 });
  const scaling = parseCompactFormat(scalingStr, { 
    spendDrop: -15, efficientProfit: 0, moderateMin: -1, moderateMax: -10 
  });
  const other = parseCompactFormat(otherStr, { modSpend: 3, modProfit: 2, stable: 2 });
  
  return {
    healthyGrowth: { 
      minSpendChange: healthy.spend || 10, 
      minProfitChange: healthy.profit || 5 
    },
    efficiencyImprovement: { 
      maxSpendDecline: efficiency.spendDrop || -5, 
      minProfitGrowth: efficiency.profitGain || 8 
    },
    inefficientGrowth: { 
      minSpendChange: 0, 
      maxProfitChange: inefficient.profitDrop || -8 
    },
    decliningEfficiency: { 
      minSpendStable: -2, 
      maxSpendGrowth: 10, 
      maxProfitDecline: -4, 
      minProfitDecline: -7 
    },
    scalingDown: { 
      maxSpendChange: scaling.spendDrop || -15,
      efficient: { minProfitChange: scaling.efficientProfit || 0 },
      moderate: { 
        maxProfitDecline: scaling.moderateMax || -10, 
        minProfitDecline: scaling.moderateMin || -1 
      },
      problematic: { maxProfitDecline: -15 }
    },
    moderateGrowthSpend: other.modSpend || 3,
    moderateGrowthProfit: other.modProfit || 2,
    minimalGrowth: { maxSpendChange: 2, maxProfitChange: 1 },
    moderateDecline: { 
      maxSpendDecline: -3, maxProfitDecline: -3, spendOptimizationRatio: 1.5,
      efficiencyDropRatio: 1.5, proportionalRatio: 1.3
    },
    stable: { maxAbsoluteChange: other.stable || 2 }
  };
}

function saveAdvancedGrowthThresholds(projectName, thresholds) {
  const sheet = getOrCreateSettingsSheet();
  const data = sheet.getDataRange().getValues();
  
  const healthyStr = `spend:${thresholds.healthyGrowth.minSpendChange},profit:${thresholds.healthyGrowth.minProfitChange}`;
  const efficiencyStr = `spendDrop:${thresholds.efficiencyImprovement.maxSpendDecline},profitGain:${thresholds.efficiencyImprovement.minProfitGrowth}`;
  const inefficientStr = `profitDrop:${thresholds.inefficientGrowth.maxProfitChange}`;
  const scalingStr = `spendDrop:${thresholds.scalingDown.maxSpendChange},efficientProfit:${thresholds.scalingDown.efficient.minProfitChange},moderateMin:${thresholds.scalingDown.moderate.minProfitDecline},moderateMax:${thresholds.scalingDown.moderate.maxProfitDecline}`;
  const otherStr = `modSpend:${thresholds.moderateGrowthSpend},modProfit:${thresholds.moderateGrowthProfit},stable:${thresholds.stable.maxAbsoluteChange}`;
  
  for (let i = 0; i < data.length; i++) {
    const label = data[i][0] ? data[i][0].toString().trim() : '';
    
    if (label === `${projectName}:` && i >= 21 && i <= 30) {
      sheet.getRange(i + 1, 2).setValue(healthyStr);
      sheet.getRange(i + 1, 3).setValue(efficiencyStr);
      sheet.getRange(i + 1, 4).setValue(inefficientStr);
      sheet.getRange(i + 1, 5).setValue(scalingStr);
      sheet.getRange(i + 1, 6).setValue(otherStr);
      
      clearSettingsCache();
      break;
    }
  }
}

function populateDefaultSettings(sheet) {
  try {
    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty('BEARER_TOKEN');
    if (token) {
      sheet.getRange('B4').setValue(token);
    }
    
    // ОБНОВЛЕНО: новые дефолты для eROAS D730
    const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    projects.forEach((proj, i) => {
      const value = props.getProperty(`TARGET_EROAS_${proj}`);
      if (value) {
        sheet.getRange(`B${7 + i}`).setValue(parseInt(value));
      } else {
        // Установить новые дефолты
        const defaultValue = proj === 'TRICKY' ? 250 : 150;
        sheet.getRange(`B${7 + i}`).setValue(defaultValue);
      }
    });
    
    const autoCache = props.getProperty('AUTO_CACHE_ENABLED') === 'true';
    const autoUpdate = props.getProperty('AUTO_UPDATE_ENABLED') === 'true';
    sheet.getRange('B17').setValue(autoCache ? 'TRUE' : 'FALSE');
    sheet.getRange('B18').setValue(autoUpdate ? 'TRUE' : 'FALSE');
    
    projects.forEach((proj, i) => {
      const row = 22 + i;
      
      let existingThresholds = null;
      try {
        const savedThresholds = props.getProperty(`GROWTH_THRESHOLDS_${proj}`);
        if (savedThresholds) {
          existingThresholds = JSON.parse(savedThresholds);
        }
      } catch (e) {
        // Игнорируем ошибки парсинга
      }
      
      if (existingThresholds) {
        const healthyStr = `spend:${existingThresholds.healthyGrowth?.minSpendChange || 10},profit:${existingThresholds.healthyGrowth?.minProfitChange || 5}`;
        const efficiencyStr = `spendDrop:${existingThresholds.efficiencyImprovement?.maxSpendDecline || -5},profitGain:${existingThresholds.efficiencyImprovement?.minProfitGrowth || 8}`;
        const inefficientStr = `profitDrop:${existingThresholds.inefficientGrowth?.maxProfitChange || -8}`;
        const scalingStr = `spendDrop:${existingThresholds.scalingDown?.maxSpendChange || -15},efficientProfit:${existingThresholds.scalingDown?.efficient?.minProfitChange || 0},moderateMin:${existingThresholds.scalingDown?.moderate?.minProfitDecline || -1},moderateMax:${existingThresholds.scalingDown?.moderate?.maxProfitDecline || -10}`;
        const otherStr = `modSpend:${existingThresholds.moderateGrowthSpend || 3},modProfit:${existingThresholds.moderateGrowthProfit || 2},stable:${existingThresholds.stable?.maxAbsoluteChange || 2}`;
        
        sheet.getRange(`B${row}`).setValue(healthyStr);
        sheet.getRange(`C${row}`).setValue(efficiencyStr);
        sheet.getRange(`D${row}`).setValue(inefficientStr);
        sheet.getRange(`E${row}`).setValue(scalingStr);
        sheet.getRange(`F${row}`).setValue(otherStr);
      }
    });
    
  } catch (e) {
    console.log('No existing settings to migrate or migration error:', e);
  }
}

function saveSettingToSheet(settingPath, value) {
  const sheet = getOrCreateSettingsSheet();
  const data = sheet.getDataRange().getValues();
  
  SETTINGS_CACHE = null;
  
  console.log(`Saving setting: ${settingPath} = ${value}`);
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const label = row[0] ? row[0].toString().trim() : '';
    
    if (settingPath === 'bearerToken' && label === 'Bearer Token:') {
      sheet.getRange(i + 1, 2).setValue(value);
      console.log(`Bearer token saved at row ${i + 1}`);
      return;
    }
    
    if (settingPath.startsWith('targetEROAS.')) {
      const project = settingPath.split('.')[1];
      if (label === `${project}:` && i >= 6 && i <= 15) { // eROAS section
        sheet.getRange(i + 1, 2).setValue(value);
        console.log(`Target eROAS D730 ${project} saved at row ${i + 1}`);
        return;
      }
    }
    
    if (settingPath === 'automation.autoCache' && label === 'Auto Cache Enabled:') {
      sheet.getRange(i + 1, 2).setValue(value ? 'TRUE' : 'FALSE');
      console.log(`Auto cache saved at row ${i + 1}: ${value ? 'TRUE' : 'FALSE'}`);
      return;
    }
    
    if (settingPath === 'automation.autoUpdate' && label === 'Auto Update Enabled:') {
      sheet.getRange(i + 1, 2).setValue(value ? 'TRUE' : 'FALSE');
      console.log(`Auto update saved at row ${i + 1}: ${value ? 'TRUE' : 'FALSE'}`);
      return;
    }
  }
  
  console.log(`Setting ${settingPath} not found in sheet`);
}

function refreshSettingsFromSheet() {
  clearSettingsCache();
  const settings = loadSettingsFromSheet();
  console.log('Settings refreshed:', settings);
  return settings;
}

function updateAutomationInSheet(autoCache, autoUpdate) {
  const sheet = getOrCreateSettingsSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 0; i < data.length; i++) {
    const label = data[i][0] ? data[i][0].toString().trim() : '';
    
    if (label === 'Auto Cache Enabled:') {
      sheet.getRange(i + 1, 2).setValue(autoCache ? 'TRUE' : 'FALSE');
    }
    
    if (label === 'Auto Update Enabled:') {
      sheet.getRange(i + 1, 2).setValue(autoUpdate ? 'TRUE' : 'FALSE');
    }
  }
  
  clearSettingsCache();
  
  console.log(`Automation updated: cache=${autoCache}, update=${autoUpdate}`);
}

function clearSettingsCache() {
  SETTINGS_CACHE = null;
  SETTINGS_CACHE_TIME = null;
}

function openSettingsSheet() {
  const sheet = getOrCreateSettingsSheet();
  const spreadsheet = SpreadsheetApp.openById(MAIN_SHEET_ID);
  spreadsheet.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('Settings Sheet', 'Лист Settings открыт. Вы можете редактировать настройки прямо в таблице.\n\nИспользуйте "🔄 Refresh Settings" после изменений.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function validateSettings() {
  const settings = loadSettingsFromSheet();
  const issues = [];
  
  if (!settings.bearerToken || settings.bearerToken.length < 50) {
    issues.push('Bearer Token не настроен или слишком короткий');
  }
  
  Object.keys(settings.targetEROAS).forEach(proj => {
    const value = settings.targetEROAS[proj];
    if (value < 100 || value > 500) {
      issues.push(`${proj}: Target eROAS D730 вне допустимого диапазона (100-500%)`);
    }
  });
  
  Object.keys(settings.growthThresholds).forEach(proj => {
    const thresholds = settings.growthThresholds[proj];
    if (!thresholds.healthyGrowth || !thresholds.scalingDown) {
      issues.push(`${proj}: Некорректные настройки Growth Thresholds`);
    }
  });
  
  return {
    valid: issues.length === 0,
    issues: issues
  };
}

function createExampleGrowthThresholds(projectName) {
  const sheet = getOrCreateSettingsSheet();
  const ui = SpreadsheetApp.getUi();
  
  const examples = {
    conservative: {
      healthy: 'spend:5,profit:3',
      efficiency: 'spendDrop:-3,profitGain:5',
      inefficient: 'profitDrop:-5',
      scaling: 'spendDrop:-10,efficientProfit:0,moderateMin:-1,moderateMax:-5',
      other: 'modSpend:2,modProfit:1,stable:1'
    },
    standard: {
      healthy: 'spend:10,profit:5',
      efficiency: 'spendDrop:-5,profitGain:8',
      inefficient: 'profitDrop:-8',
      scaling: 'spendDrop:-15,efficientProfit:0,moderateMin:-1,moderateMax:-10',
      other: 'modSpend:3,modProfit:2,stable:2'
    },
    aggressive: {
      healthy: 'spend:20,profit:10',
      efficiency: 'spendDrop:-10,profitGain:15',
      inefficient: 'profitDrop:-15',
      scaling: 'spendDrop:-25,efficientProfit:5,moderateMin:-5,moderateMax:-20',
      other: 'modSpend:5,modProfit:3,stable:3'
    }
  };
  
  const choice = ui.alert(`📊 ${projectName} Growth Thresholds Examples`, 
    'Выберите пример настроек:\n\nYES = Conservative (осторожные)\nNO = Standard (стандартные)\nCANCEL = Aggressive (агрессивные)', 
    ui.ButtonSet.YES_NO_CANCEL);
  
  let selectedExample;
  if (choice === ui.Button.YES) selectedExample = examples.conservative;
  else if (choice === ui.Button.NO) selectedExample = examples.standard;
  else if (choice === ui.Button.CANCEL) selectedExample = examples.aggressive;
  else return;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    const label = data[i][0] ? data[i][0].toString().trim() : '';
    
    if (label === `${projectName}:` && i >= 21 && i <= 30) {
      sheet.getRange(i + 1, 2).setValue(selectedExample.healthy);
      sheet.getRange(i + 1, 3).setValue(selectedExample.efficiency);
      sheet.getRange(i + 1, 4).setValue(selectedExample.inefficient);
      sheet.getRange(i + 1, 5).setValue(selectedExample.scaling);
      sheet.getRange(i + 1, 6).setValue(selectedExample.other);
      
      clearSettingsCache();
      ui.alert('✅ Применено', `${projectName} настройки обновлены примером.`, ui.ButtonSet.OK);
      break;
    }
  }
}

function exportSettings() {
  const settings = loadSettingsFromSheet();
  const safeSettings = {
    targetEROAS: settings.targetEROAS,
    automation: settings.automation,
    growthThresholds: settings.growthThresholds
  };
  
  return JSON.stringify(safeSettings, null, 2);
}