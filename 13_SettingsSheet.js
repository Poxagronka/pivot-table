/**
 * Settings Sheet Management - ОБНОВЛЕНО: улучшенное кеширование и обработка таймаутов + INCENT_TRAFFIC
 */

var SETTINGS_SHEET_NAME = 'Settings';
var SETTINGS_CACHE = null;
var SETTINGS_CACHE_TIME = null;
var SETTINGS_CACHE_DURATION = 300000; // 5 минут кеша вместо 30 секунд

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
  
  let hasOldStructure = false;
  let hasOldFormatting = false;
  
  for (let i = 0; i < data.length; i++) {
    const cellValue = data[i][0] ? data[i][0].toString() : '';
    if (cellValue === 'TRICKY:' || cellValue === 'MOLOCO:') {
      hasOldStructure = true;
      break;
    }
  }
  
  if (data.length > 0 && data[0].length < 8) {
    hasOldFormatting = true;
  }
  
  let hasProperFormatting = false;
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const cellValue = data[i][0] ? data[i][0].toString() : '';
    if (cellValue === '⚙️ CAMPAIGN REPORT SETTINGS') {
      hasProperFormatting = true;
      break;
    }
  }
  
  if (hasOldStructure || hasOldFormatting || !hasProperFormatting) {
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
    
    sheet.clear();
    createSettingsLayout(sheet);
    
    if (savedToken) {
      const tokenRow = findTokenRow(sheet);
      if (tokenRow > 0) {
        sheet.getRange(tokenRow, 2, 1, 7).merge().setValue(savedToken);
      }
    }
    sheet.getRange('B4').setValue(savedAutoCache ? 'TRUE' : 'FALSE');
    sheet.getRange('B5').setValue(savedAutoUpdate ? 'TRUE' : 'FALSE');
    
    console.log('Settings migrated to fixed targets structure');
  }
}

function findTokenRow(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === 'Bearer Token:') {
      return i + 1;
    }
  }
  return -1;
}

function createSettingsLayout(sheet) {
  sheet.clear();
  
  // Заголовок
  sheet.getRange('A1:H1').merge().setValue('⚙️ CAMPAIGN REPORT SETTINGS');
  sheet.getRange('A1:H1').setBackground('#1c4587').setFontColor('white').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);
  
  sheet.setRowHeight(2, 20);
  
  // Automation
  sheet.getRange('A3:H3').merge().setValue('🤖 AUTOMATION').setBackground('#ff9800').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(3, 30);
  
  sheet.getRange('A4').setValue('Auto Cache Enabled:').setFontWeight('bold');
  sheet.getRange('B4').setValue('FALSE');
  sheet.getRange('C4:H4').merge().setValue('Every hour - saves comments automatically').setFontStyle('italic');
  sheet.getRange('A4:A4').setBackground('#fff3e0');
  sheet.getRange('B4:B4').setBackground('#f8f9fa');
  sheet.getRange('C4:H4').setBackground('#f8f9fa');
  sheet.setRowHeight(4, 25);
  
  sheet.getRange('A5').setValue('Auto Update Enabled:').setFontWeight('bold');
  sheet.getRange('B5').setValue('FALSE');
  sheet.getRange('C5:H5').merge().setValue('Daily at 5:00 AM - updates all projects data').setFontStyle('italic');
  sheet.getRange('A5:A5').setBackground('#fff3e0');
  sheet.getRange('B5:B5').setBackground('#f8f9fa');
  sheet.getRange('C5:H5').setBackground('#f8f9fa');
  sheet.setRowHeight(5, 25);
  
  sheet.setRowHeight(6, 15);
  
  // Target eROAS D730 
  sheet.getRange('A7:H7').merge().setValue('🎯 TARGET eROAS D730 (%)').setBackground('#34a853').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(7, 30);
  
  const appTypes = [
    { name: 'TRICKY Project:', value: 250, desc: 'Весь лист Tricky' },
    { name: 'Business Apps:', value: 140, desc: 'Приложения со словом "Business"' },
    { name: 'Other Apps:', value: 150, desc: 'Все остальные приложения' }
  ];
  
  appTypes.forEach((appType, i) => {
    const row = 8 + i;
    sheet.getRange(`A${row}`).setValue(appType.name).setFontWeight('bold');
    sheet.getRange(`B${row}`).setValue(appType.value).setHorizontalAlignment('center').setFontWeight('bold');
    sheet.getRange(`C${row}:H${row}`).merge().setValue(appType.desc).setFontStyle('italic');
    sheet.getRange(`A${row}:A${row}`).setBackground('#e8f5e8');
    sheet.getRange(`B${row}:B${row}`).setBackground('#d4edda');
    sheet.getRange(`C${row}:H${row}`).setBackground('#f8f9fa');
    sheet.setRowHeight(row, 25);
  });
  
  sheet.setRowHeight(11, 15);
  
  // Advanced Growth Thresholds
  sheet.getRange('A12:H12').merge().setValue('📊 GROWTH THRESHOLDS (Advanced)').setBackground('#9c27b0').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(12, 30);
  
  // Заголовки с увеличенной шириной
  sheet.getRange('A13').setValue('Project').setFontWeight('bold').setWrap(true);
  sheet.getRange('B13').setValue('Healthy Growth').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('C13').setValue('Efficiency').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('D13').setValue('Inefficient').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('E13').setValue('Scaling Down').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('F13').setValue('Other').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('G13').setValue('Status').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('H13').setValue('Modified').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  sheet.getRange('A13:H13').setBackground('#f3e5f5');
  sheet.setRowHeight(13, 25);
  
  // ОБНОВЛЕНО: добавлен INCENT_TRAFFIC в список проектов
  const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'INCENT_TRAFFIC', 'OVERALL'];
  projects.forEach((proj, i) => {
    const row = 14 + i;
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
    
    sheet.getRange(`A${row}:H${row}`).setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  });
  
  sheet.setRowHeight(23, 20);
  
  // Detailed Instructions
  sheet.getRange('A24:H24').merge().setValue('📖 INSTRUCTIONS').setBackground('#607d8b').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(24, 30);
  
  // Target eROAS Instructions
  sheet.getRange('A25').setValue('🎯 Target eROAS Logic:').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A26:H28').merge();
  sheet.getRange('A26').setValue(
    '• TRICKY проект: всегда 250% (весь лист)\n' +
    '• Business приложения: 140% (со словом "Business" в любом проекте)\n' +
    '• Все остальные: 150% (по умолчанию)'
  );
  sheet.getRange('A26:H28').setBackground('#f5f5f5').setWrap(true).setBorder(true, true, true, true, false, false);
  
  // Growth Thresholds Instructions
  sheet.getRange('A30').setValue('📊 Growth Thresholds:').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A31:H34').merge();
  sheet.getRange('A31').setValue(
    '🟢 HEALTHY: spend:X,profit:Y - оба условия выполняются\n' +
    '🟢 EFFICIENCY: spendDrop:X,profitGain:Y - тратим меньше, зарабатываем больше\n' +
    '🔴 INEFFICIENT: profitDrop:X - критическое падение прибыли\n' +
    '🔵 SCALING DOWN: spendDrop:X - значительное сокращение спенда'
  );
  sheet.getRange('A31:H34').setBackground('#f5f5f5').setWrap(true).setBorder(true, true, true, true, false, false);
  
  sheet.setRowHeight(36, 20);
  
  // API Settings в конце
  sheet.getRange('A37:H37').merge().setValue('🔐 API SETTINGS').setBackground('#4285f4').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(37, 30);
  
  sheet.getRange('A38').setValue('Bearer Token:').setFontWeight('bold');
  sheet.getRange('B38:H38').merge().setValue('[ENTER_YOUR_TOKEN_HERE]');
  sheet.getRange('A38:A38').setBackground('#e8f0fe');
  sheet.getRange('B38:H38').setBackground('#f8f9fa').setBorder(true, true, true, true, false, false);
  sheet.setRowHeight(38, 25);
  
  sheet.getRange('A40:H42').merge();
  sheet.getRange('A40').setValue(
    '• Bearer Token: Получите из app.appodeal.com → Settings → API\n' +
    '• Токен должен начинаться с "eyJ" и быть длиной 300+ символов\n' +
    '• Один токен работает для всех проектов'
  );
  sheet.getRange('A40:H42').setBackground('#f5f5f5').setWrap(true).setBorder(true, true, true, true, false, false);
  
  // Настройка ширины колонок
  sheet.setColumnWidth(1, 140);  // Project
  sheet.setColumnWidth(2, 160);  // Healthy Growth  
  sheet.setColumnWidth(3, 160);  // Efficiency
  sheet.setColumnWidth(4, 120);  // Inefficient Growth
  sheet.setColumnWidth(5, 220);  // Scaling Down
  sheet.setColumnWidth(6, 160);  // Other
  sheet.setColumnWidth(7, 80);   // Status
  sheet.setColumnWidth(8, 100);  // Modified
  
  // Валидация для automation
  sheet.getRange('B4:B5').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TRUE', 'FALSE']).build());
  
  sheet.setFrozenRows(3);
}

function loadSettingsFromSheet() {
  const now = new Date().getTime();
  
  // Увеличиваем время кеширования до 5 минут
  if (SETTINGS_CACHE && SETTINGS_CACHE_TIME && (now - SETTINGS_CACHE_TIME) < SETTINGS_CACHE_DURATION) {
    return SETTINGS_CACHE;
  }
  
  // Попытка загрузить настройки с обработкой таймаутов
  let retries = 3;
  let lastError = null;
  
  while (retries > 0) {
    try {
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
        
        // Target eROAS D730 по новой логике
        if (label === 'TRICKY Project:' && i >= 7 && i <= 11) {
          const numValue = parseInt(value);
          settings.targetEROAS.tricky = (!isNaN(numValue) && numValue >= 100 && numValue <= 500) ? numValue : 250;
        }
        
        if (label === 'Business Apps:' && i >= 7 && i <= 11) {
          const numValue = parseInt(value);
          settings.targetEROAS.business = (!isNaN(numValue) && numValue >= 100 && numValue <= 500) ? numValue : 140;
        }
        
        if (label === 'Other Apps:' && i >= 7 && i <= 11) {
          const numValue = parseInt(value);
          settings.targetEROAS.ceg = (!isNaN(numValue) && numValue >= 100 && numValue <= 500) ? numValue : 150;
        }
        
        // ОБНОВЛЕНО: Advanced Growth Thresholds по проектам (включая INCENT_TRAFFIC)
        const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'INCENT_TRAFFIC', 'OVERALL'];
        projects.forEach(proj => {
          if (label === proj && i >= 13 && i <= 23) {
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
      
    } catch (e) {
      lastError = e;
      retries--;
      
      if (e.toString().includes('timed out') || e.toString().includes('Service Spreadsheets')) {
        console.log('Timeout loading settings, retries left:', retries);
        if (retries > 0) {
          Utilities.sleep(3000 * (4 - retries)); // Экспоненциальная задержка
          SpreadsheetApp.flush();
        }
      } else {
        // Для других ошибок сразу выходим
        throw e;
      }
    }
  }
  
  // Если все попытки исчерпаны, возвращаем дефолтные настройки
  console.error('Failed to load settings after all retries:', lastError);
  
  // Возвращаем последний успешный кеш или дефолтные значения
  if (SETTINGS_CACHE) {
    console.log('Returning cached settings despite timeout');
    return SETTINGS_CACHE;
  }
  
  // Дефолтные настройки
  return {
    bearerToken: '',
    targetEROAS: { tricky: 250, business: 140, ceg: 150 },
    automation: { autoCache: false, autoUpdate: false },
    growthThresholds: getDefaultGrowthThresholdsForAllProjects()
  };
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
      const tokenRow = findTokenRow(sheet);
      if (tokenRow > 0) {
        sheet.getRange(tokenRow, 2, 1, 7).merge().setValue(token);
      }
    }
    
    const autoCache = props.getProperty('AUTO_CACHE_ENABLED') === 'true';
    const autoUpdate = props.getProperty('AUTO_UPDATE_ENABLED') === 'true';
    sheet.getRange('B4').setValue(autoCache ? 'TRUE' : 'FALSE');
    sheet.getRange('B5').setValue(autoUpdate ? 'TRUE' : 'FALSE');
    
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
    
    if (settingPath === 'targetEROAS.tricky' && label === 'TRICKY Project:') {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
    
    if (settingPath === 'targetEROAS.business' && label === 'Business Apps:') {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
    
    if (settingPath === 'targetEROAS.ceg' && label === 'Other Apps:') {
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
  SpreadsheetApp.getUi().alert('Settings Sheet', 'Лист Settings с исправленной логикой таргетов!\n\n🎯 TRICKY: 250% (весь проект)\n💼 Business: 140% (приложения с "Business")\n📱 Остальные: 150% (по умолчанию)\n\nИспользуйте "🔄 Refresh Settings" после изменений.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ОБНОВЛЕНО: включен INCENT_TRAFFIC в список проектов для получения дефолтных порогов
function getDefaultGrowthThresholdsForAllProjects() {
  const defaultThresholds = getDefaultGrowthThresholds();
  const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'INCENT_TRAFFIC', 'OVERALL'];
  const result = {};
  
  projects.forEach(proj => {
    result[proj] = defaultThresholds;
  });
  
  return result;
}

// Функция для предварительной загрузки настроек
function preloadSettings() {
  try {
    loadSettingsFromSheet();
    console.log('Settings preloaded successfully');
  } catch (e) {
    console.error('Error preloading settings:', e);
  }
}