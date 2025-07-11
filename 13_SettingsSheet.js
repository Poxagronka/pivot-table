/**
 * Settings Sheet Management - ОБНОВЛЕНО: таргеты по типам приложений вместо проектов
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
  } else {
    migrateExistingSettings(sheet);
  }
  
  return sheet;
}

function migrateExistingSettings(sheet) {
  const data = sheet.getDataRange().getValues();
  let needsUpdate = false;
  
  // Проверяем есть ли старая структура по проектам или плохое форматирование
  let hasOldStructure = false;
  let hasOldFormatting = false;
  
  for (let i = 0; i < data.length; i++) {
    const cellValue = data[i][0] ? data[i][0].toString() : '';
    if (cellValue === 'TRICKY:' || cellValue === 'MOLOCO:') {
      hasOldStructure = true;
      break;
    }
  }
  
  // Проверяем количество колонок - если меньше 8, то старое форматирование
  if (data.length > 0 && data[0].length < 8) {
    hasOldFormatting = true;
  }
  
  // Проверяем есть ли заголовок с правильным форматированием
  let hasProperFormatting = false;
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const cellValue = data[i][0] ? data[i][0].toString() : '';
    if (cellValue === '⚙️ CAMPAIGN REPORT SETTINGS') {
      hasProperFormatting = true;
      break;
    }
  }
  
  if (hasOldStructure || hasOldFormatting || !hasProperFormatting) {
    // Сохраняем существующие значения
    let savedToken = '';
    let savedAutoCache = false;
    let savedAutoUpdate = false;
    
    try {
      for (let i = 0; i < data.length; i++) {
        const label = data[i][0] ? data[i][0].toString().trim() : '';
        const value = data[i][1] ? data[i][1].toString().trim() : '';
        
        if (label === 'Bearer Token:' && value && value !== '[ENTER_YOUR_TOKEN_HERE]') {
          savedToken = value;
        }
        if (label === 'Auto Cache Enabled:') {
          savedAutoCache = value.toUpperCase() === 'TRUE';
        }
        if (label === 'Auto Update Enabled:') {
          savedAutoUpdate = value.toUpperCase() === 'TRUE';
        }
      }
    } catch (e) {
      console.log('Error reading old settings:', e);
    }
    
    // Полностью пересоздаем лист с новой структурой
    sheet.clear();
    createSettingsLayout(sheet);
    
    // Восстанавливаем сохраненные значения
    if (savedToken) {
      sheet.getRange('B4:H4').setValue(savedToken);
    }
    sheet.getRange('B12').setValue(savedAutoCache ? 'TRUE' : 'FALSE');
    sheet.getRange('B13').setValue(savedAutoUpdate ? 'TRUE' : 'FALSE');
    
    console.log('Settings migrated to new UX-friendly structure');
  }
}

function createSettingsLayout(sheet) {
  sheet.clear();
  
  // Заголовок
  sheet.getRange('A1:H1').merge().setValue('⚙️ CAMPAIGN REPORT SETTINGS');
  sheet.getRange('A1:H1').setBackground('#1c4587').setFontColor('white').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);
  
  // Пустая строка
  sheet.setRowHeight(2, 20);
  
  // API Settings
  sheet.getRange('A3:H3').merge().setValue('🔐 API SETTINGS').setBackground('#4285f4').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(3, 30);
  
  sheet.getRange('A4').setValue('Bearer Token:').setFontWeight('bold');
  sheet.getRange('B4:H4').merge().setValue('[ENTER_YOUR_TOKEN_HERE]');
  sheet.getRange('A4:A4').setBackground('#e8f0fe');
  sheet.getRange('B4:H4').setBackground('#f8f9fa').setBorder(true, true, true, true, false, false);
  sheet.setRowHeight(4, 25);
  
  // Пустая строка
  sheet.setRowHeight(5, 15);
  
  // Target eROAS D730 
  sheet.getRange('A6:H6').merge().setValue('🎯 TARGET eROAS D730 (%)').setBackground('#34a853').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(6, 30);
  
  const appTypes = [
    { name: 'Tricky Apps:', value: 250, desc: 'Word games, puzzles' },
    { name: 'Business Empire:', value: 140, desc: 'Business simulation games' },
    { name: 'CEG Apps:', value: 150, desc: 'All other apps' }
  ];
  
  appTypes.forEach((appType, i) => {
    const row = 7 + i;
    sheet.getRange(`A${row}`).setValue(appType.name).setFontWeight('bold');
    sheet.getRange(`B${row}`).setValue(appType.value).setHorizontalAlignment('center').setFontWeight('bold');
    sheet.getRange(`C${row}:H${row}`).merge().setValue(appType.desc).setFontStyle('italic');
    sheet.getRange(`A${row}:A${row}`).setBackground('#e8f5e8');
    sheet.getRange(`B${row}:B${row}`).setBackground('#d4edda');
    sheet.getRange(`C${row}:H${row}`).setBackground('#f8f9fa');
    sheet.setRowHeight(row, 25);
  });
  
  // Пустая строка
  sheet.setRowHeight(10, 15);
  
  // Automation
  sheet.getRange('A11:H11').merge().setValue('🤖 AUTOMATION').setBackground('#ff9800').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(11, 30);
  
  sheet.getRange('A12').setValue('Auto Cache Enabled:').setFontWeight('bold');
  sheet.getRange('B12').setValue('FALSE');
  sheet.getRange('C12:H12').merge().setValue('Daily at 2:00 AM - saves comments automatically').setFontStyle('italic');
  sheet.getRange('A12:A12').setBackground('#fff3e0');
  sheet.getRange('B12:B12').setBackground('#f8f9fa');
  sheet.getRange('C12:H12').setBackground('#f8f9fa');
  sheet.setRowHeight(12, 25);
  
  sheet.getRange('A13').setValue('Auto Update Enabled:').setFontWeight('bold');
  sheet.getRange('B13').setValue('FALSE');
  sheet.getRange('C13:H13').merge().setValue('Daily at 5:00 AM - updates all projects data').setFontStyle('italic');
  sheet.getRange('A13:A13').setBackground('#fff3e0');
  sheet.getRange('B13:B13').setBackground('#f8f9fa');
  sheet.getRange('C13:H13').setBackground('#f8f9fa');
  sheet.setRowHeight(13, 25);
  
  // Пустая строка
  sheet.setRowHeight(14, 15);
  
  // Advanced Growth Thresholds
  sheet.getRange('A15:H15').merge().setValue('📊 GROWTH THRESHOLDS (Advanced)').setBackground('#9c27b0').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(15, 30);
  
  // Заголовки с переносами
  sheet.getRange('A16').setValue('Project').setFontWeight('bold').setWrap(true);
  sheet.getRange('B16').setValue('Healthy\nGrowth').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('C16').setValue('Efficiency\nImprovement').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('D16').setValue('Inefficient\nGrowth').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('E16').setValue('Scaling\nDown').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('F16').setValue('Other\nThresholds').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('G16').setValue('Status').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('H16').setValue('Modified').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('A16:H16').setBackground('#f3e5f5');
  sheet.setRowHeight(16, 35);
  
  const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
  projects.forEach((proj, i) => {
    const row = 17 + i;
    sheet.getRange(`A${row}`).setValue(proj).setFontWeight('bold');
    sheet.getRange(`B${row}`).setValue('spend:10,profit:5').setWrap(true);
    sheet.getRange(`C${row}`).setValue('spendDrop:-5,profitGain:8').setWrap(true);
    sheet.getRange(`D${row}`).setValue('profitDrop:-8').setWrap(true);
    sheet.getRange(`E${row}`).setValue('spendDrop:-15,efficientProfit:0,moderateMin:-1,moderateMax:-10').setWrap(true);
    sheet.getRange(`F${row}`).setValue('modSpend:3,modProfit:2,stable:2').setWrap(true);
    sheet.getRange(`G${row}`).setValue('✅ Active').setHorizontalAlignment('center').setFontColor('#28a745');
    sheet.getRange(`H${row}`).setValue('Default').setHorizontalAlignment('center').setFontStyle('italic');
    sheet.getRange(`A${row}:A${row}`).setBackground('#fce4ec');
    sheet.setRowHeight(row, 30);
    
    // Добавляем границы
    sheet.getRange(`A${row}:H${row}`).setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  });
  
  // Пустая строка
  sheet.setRowHeight(25, 20);
  
  // Detailed Instructions
  sheet.getRange('A26:H26').merge().setValue('📖 DETAILED INSTRUCTIONS').setBackground('#607d8b').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(26, 30);
  
  // API Instructions
  sheet.getRange('A28').setValue('🔐 API Settings:').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A29:H31').merge();
  sheet.getRange('A29').setValue(
    '• Bearer Token: Получите из app.appodeal.com → Settings → API\n' +
    '• Токен должен начинаться с "eyJ" и быть длиной 300+ символов\n' +
    '• Один токен работает для всех проектов'
  );
  sheet.getRange('A29:H31').setBackground('#f5f5f5').setWrap(true).setBorder(true, true, true, true, false, false);
  
  // Target eROAS Instructions
  sheet.getRange('A33').setValue('🎯 Target eROAS D730:').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A34:H36').merge();
  sheet.getRange('A34').setValue(
    '• Taргеты по типам приложений для цветового кодирования\n' +
    '• Tricky: 250% | Business Empire: 140% | CEG: 150%\n' +
    '• Зеленый: ≥ таргета, Желтый: 120-таргет, Красный: <120%'
  );
  sheet.getRange('A34:H36').setBackground('#f5f5f5').setWrap(true).setBorder(true, true, true, true, false, false);
  
  // Growth Thresholds Instructions
  sheet.getRange('A38').setValue('📊 Growth Thresholds:').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A39:H45').merge();
  sheet.getRange('A39').setValue(
    '🟢 HEALTHY GROWTH (spend:X,profit:Y): оба условия выполняются\n' +
    '🟢 EFFICIENCY IMPROVEMENT (spendDrop:X,profitGain:Y): тратим меньше, зарабатываем больше\n' +
    '🔴 INEFFICIENT GROWTH (profitDrop:X): критическое падение прибыли\n' +
    '🔵 SCALING DOWN (spendDrop:X): значительное сокращение спенда\n' +
    '🟡 УМЕРЕННЫЕ: различные паттерны умеренного роста/спада\n' +
    '⚪ STABLE: минимальные изменения'
  );
  sheet.getRange('A39:H45').setBackground('#f5f5f5').setWrap(true).setBorder(true, true, true, true, false, false);
  
  // Настройка ширины колонок для лучшего отображения
  sheet.setColumnWidth(1, 140);  // Project
  sheet.setColumnWidth(2, 120);  // Healthy Growth
  sheet.setColumnWidth(3, 140);  // Efficiency Improvement  
  sheet.setColumnWidth(4, 100);  // Inefficient Growth
  sheet.setColumnWidth(5, 200);  // Scaling Down
  sheet.setColumnWidth(6, 140);  // Other Thresholds
  sheet.setColumnWidth(7, 80);   // Status
  sheet.setColumnWidth(8, 100);  // Modified
  
  // Валидация для automation
  sheet.getRange('B12:B13').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TRUE', 'FALSE']).build());
  
  // Замораживаем верхние строки для удобства
  sheet.setFrozenRows(3);
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
    targetEROAS: { tricky: 250, business: 140, ceg: 150 },
    automation: { autoCache: false, autoUpdate: false },
    growthThresholds: {}
  };
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const label = row[0] ? row[0].toString().trim() : '';
    const value = row[1] ? row[1].toString().trim() : '';
    
    if (label === 'Bearer Token:' && value && value !== '[ENTER_YOUR_TOKEN_HERE]') {
      settings.bearerToken = value;
    }
    
    if (label === 'Auto Cache Enabled:') {
      settings.automation.autoCache = value.toUpperCase() === 'TRUE';
    }
    
    if (label === 'Auto Update Enabled:') {
      settings.automation.autoUpdate = value.toUpperCase() === 'TRUE';
    }
    
    // Target eROAS D730 по типам приложений с новыми названиями
    if (label === 'Tricky Apps:' && i >= 6 && i <= 10) {
      const numValue = parseInt(value);
      settings.targetEROAS.tricky = (!isNaN(numValue) && numValue >= 100 && numValue <= 500) ? numValue : 250;
    }
    
    if (label === 'Business Empire:' && i >= 6 && i <= 10) {
      const numValue = parseInt(value);
      settings.targetEROAS.business = (!isNaN(numValue) && numValue >= 100 && numValue <= 500) ? numValue : 140;
    }
    
    if (label === 'CEG Apps:' && i >= 6 && i <= 10) {
      const numValue = parseInt(value);
      settings.targetEROAS.ceg = (!isNaN(numValue) && numValue >= 100 && numValue <= 500) ? numValue : 150;
    }
    
    // Advanced Growth Thresholds по проектам (номера строк изменились)
    const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    projects.forEach(proj => {
      if (label === proj && i >= 16 && i <= 25) {
        const healthyValue = row[1] ? row[1].toString() : 'spend:10,profit:5';
        const efficiencyValue = row[2] ? row[2].toString() : 'spendDrop:-5,profitGain:8';
        const inefficientValue = row[3] ? row[3].toString() : 'profitDrop:-8';
        const scalingValue = row[4] ? row[4].toString() : 'spendDrop:-15,efficientProfit:0,moderateMin:-1,moderateMax:-10';
        const otherValue = row[5] ? row[5].toString() : 'modSpend:3,modProfit:2,stable:2';
        
        settings.growthThresholds[proj] = parseAdvancedGrowthThresholds(
          healthyValue, efficiencyValue, inefficientValue, scalingValue, otherValue
        );
      }
    });
  }
  
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

function populateDefaultSettings(sheet) {
  try {
    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty('BEARER_TOKEN');
    if (token) {
      sheet.getRange('B4:H4').setValue(token);
    }
    
    // Дефолтные значения уже установлены в createSettingsLayout
    
    const autoCache = props.getProperty('AUTO_CACHE_ENABLED') === 'true';
    const autoUpdate = props.getProperty('AUTO_UPDATE_ENABLED') === 'true';
    sheet.getRange('B12').setValue(autoCache ? 'TRUE' : 'FALSE');
    sheet.getRange('B13').setValue(autoUpdate ? 'TRUE' : 'FALSE');
    
    console.log('Default settings populated successfully');
  } catch (e) {
    console.log('Error populating default settings:', e);
  }
}

function saveSettingToSheet(settingPath, value) {
  const sheet = getOrCreateSettingsSheet();
  const data = sheet.getDataRange().getValues();
  
  SETTINGS_CACHE = null;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const label = row[0] ? row[0].toString().trim() : '';
    
    if (settingPath === 'bearerToken' && label === 'Bearer Token:') {
      sheet.getRange(i + 1, 2, 1, 7).merge().setValue(value);
      return;
    }
    
    if (settingPath === 'targetEROAS.tricky' && label === 'Tricky Apps:') {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
    
    if (settingPath === 'targetEROAS.business' && label === 'Business Empire:') {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
    
    if (settingPath === 'targetEROAS.ceg' && label === 'CEG Apps:') {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
    
    if (settingPath === 'automation.autoCache' && label === 'Auto Cache Enabled:') {
      sheet.getRange(i + 1, 2).setValue(value ? 'TRUE' : 'FALSE');
      return;
    }
    
    if (settingPath === 'automation.autoUpdate' && label === 'Auto Update Enabled:') {
      sheet.getRange(i + 1, 2).setValue(value ? 'TRUE' : 'FALSE');
      return;
    }
  }
}

function refreshSettingsFromSheet() {
  clearSettingsCache();
  const settings = loadSettingsFromSheet();
  return settings;
}

function clearSettingsCache() {
  SETTINGS_CACHE = null;
  SETTINGS_CACHE_TIME = null;
}

function openSettingsSheet() {
  const sheet = getOrCreateSettingsSheet();
  const spreadsheet = SpreadsheetApp.openById(MAIN_SHEET_ID);
  spreadsheet.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('Settings Sheet', 'Лист Settings открыт с новым UX-дизайном!\n\n✨ Улучшенное форматирование\n📊 Четкая структура по разделам\n🎯 Таргеты по типам приложений\n\nИспользуйте "🔄 Refresh Settings" после изменений.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function forceUpdateSettingsSheet() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('🔄 Force Update Settings', 'Принудительно обновить лист Settings?\n\nЭто создаст новую UX-структуру с улучшенным форматированием.', ui.ButtonSet.YES_NO);
  
  if (result === ui.Button.YES) {
    const spreadsheet = SpreadsheetApp.openById(MAIN_SHEET_ID);
    let sheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
    
    if (sheet) {
      spreadsheet.deleteSheet(sheet);
    }
    
    sheet = spreadsheet.insertSheet(SETTINGS_SHEET_NAME);
    createSettingsLayout(sheet);
    populateDefaultSettings(sheet);
    clearSettingsCache();
    
    ui.alert('✅ Updated', 'Лист Settings обновлен с новым UX!\n\n📊 Улучшенное форматирование\n🎯 Таргеты по типам:\n• Tricky Apps: 250%\n• Business Empire: 140%\n• CEG Apps: 150%\n\n💡 Лучшая читаемость и навигация', ui.ButtonSet.OK);
  }
}