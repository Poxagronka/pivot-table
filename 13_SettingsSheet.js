var SETTINGS_SHEET_NAME = 'Settings';
var SETTINGS_CACHE = null;
var SETTINGS_CACHE_TIME = null;
var SETTINGS_LOAD_LOCK = LockService.getScriptLock();

function getOrCreateSettingsSheet() {
  try {
    SETTINGS_LOAD_LOCK.waitLock(10000);
    
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
  } catch (e) {
    console.error('Error getting settings sheet:', e);
    throw e;
  } finally {
    SETTINGS_LOAD_LOCK.releaseLock();
  }
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
        sheet.getRange(tokenRow, 2, 1, 8).merge().setValue(savedToken);
      }
    }
    sheet.getRange('B4').setValue(savedAutoCache ? 'TRUE' : 'FALSE');
    sheet.getRange('B5').setValue(savedAutoUpdate ? 'TRUE' : 'FALSE');
    
    console.log('Settings migrated to separate triggers structure');
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
  
  sheet.getRange('A1:I1').merge().setValue('⚙️ CAMPAIGN REPORT SETTINGS');
  sheet.getRange('A1:I1').setBackground('#1c4587').setFontColor('white').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);
  
  sheet.setRowHeight(2, 20);
  
  sheet.getRange('A3:I3').merge().setValue('🤖 AUTOMATION').setBackground('#ff9800').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(3, 30);
  
  sheet.getRange('A4').setValue('Auto Cache Enabled:').setFontWeight('bold');
  sheet.getRange('B4').setValue('FALSE');
  sheet.getRange('C4:I4').merge().setValue('Daily at 2:00 AM - saves comments automatically').setFontStyle('italic');
  sheet.getRange('A4:A4').setBackground('#fff3e0');
  sheet.getRange('B4:B4').setBackground('#f8f9fa');
  sheet.getRange('C4:I4').setBackground('#f8f9fa');
  sheet.setRowHeight(4, 25);
  
  sheet.getRange('A5').setValue('Auto Update Enabled:').setFontWeight('bold');
  sheet.getRange('B5').setValue('FALSE');
  sheet.getRange('C5:I5').merge().setValue('Staggered updates 4:00-7:30 AM (30min intervals)').setFontStyle('italic');
  sheet.getRange('A5:A5').setBackground('#fff3e0');
  sheet.getRange('B5:B5').setBackground('#f8f9fa');
  sheet.getRange('C5:I5').setBackground('#f8f9fa');
  sheet.setRowHeight(5, 25);
  
  sheet.setRowHeight(6, 15);
  
  sheet.getRange('A7:I7').merge().setValue('🕐 UPDATE SCHEDULE').setBackground('#4caf50').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(7, 30);
  
  const schedule = [
    { project: 'TRICKY', time: '4:00 AM', desc: 'Heaviest project first' },
    { project: 'MOLOCO', time: '4:30 AM', desc: 'APD campaigns' },
    { project: 'REGULAR', time: '5:00 AM', desc: 'Non-tricky campaigns' },
    { project: 'GOOGLE_ADS', time: '5:30 AM', desc: 'Google Ads network' },
    { project: 'APPLOVIN', time: '6:00 AM', desc: 'AppLovin network' },
    { project: 'MINTEGRAL', time: '6:30 AM', desc: 'Mintegral network' },
    { project: 'INCENT', time: '7:00 AM', desc: 'Incentivized traffic' },
    { project: 'OVERALL', time: '7:30 AM', desc: 'Aggregated app-level data' }
  ];
  
  schedule.forEach((item, i) => {
    const row = 8 + i;
    sheet.getRange(`A${row}`).setValue(`${item.project}:`).setFontWeight('bold');
    sheet.getRange(`B${row}`).setValue(item.time).setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#1976d2');
    sheet.getRange(`C${row}:I${row}`).merge().setValue(item.desc).setFontStyle('italic');
    sheet.getRange(`A${row}:A${row}`).setBackground('#e8f5e8');
    sheet.getRange(`B${row}:B${row}`).setBackground('#e3f2fd');
    sheet.getRange(`C${row}:I${row}`).setBackground('#f8f9fa');
    sheet.setRowHeight(row, 25);
  });
  
  sheet.setRowHeight(16, 15);
  
  sheet.getRange('A17:I17').merge().setValue('🎯 TARGET eROAS D730 (%)').setBackground('#34a853').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(17, 30);
  
  const appTypes = [
    { name: 'TRICKY Project:', value: 250, desc: 'All data in TRICKY sheet' },
    { name: 'Business Apps:', value: 140, desc: 'Apps with "Business" in the name' },
    { name: 'Other Apps:', value: 150, desc: 'All other applications' }
  ];
  
  appTypes.forEach((appType, i) => {
    const row = 18 + i;
    sheet.getRange(`A${row}`).setValue(appType.name).setFontWeight('bold');
    sheet.getRange(`B${row}`).setValue(appType.value).setHorizontalAlignment('center').setFontWeight('bold');
    sheet.getRange(`C${row}:I${row}`).merge().setValue(appType.desc).setFontStyle('italic');
    sheet.getRange(`A${row}:A${row}`).setBackground('#e8f5e8');
    sheet.getRange(`B${row}:B${row}`).setBackground('#d4edda');
    sheet.getRange(`C${row}:I${row}`).setBackground('#f8f9fa');
    sheet.setRowHeight(row, 25);
  });
  
  sheet.setRowHeight(21, 15);
  
  sheet.getRange('A22:I22').merge().setValue('📊 GROWTH THRESHOLDS').setBackground('#9c27b0').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(22, 30);
  
  sheet.getRange('A23').setValue('Project').setFontWeight('bold').setWrap(false).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
  sheet.getRange('B23').setValue('🟢 Healthy Growth (Spend/Profit %)').setFontWeight('bold').setWrap(false).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
  sheet.getRange('C23').setValue('🟢 Efficiency (Spend/Profit %)').setFontWeight('bold').setWrap(false).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
  sheet.getRange('D23').setValue('🔴 Warning (Profit %)').setFontWeight('bold').setWrap(false).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
  sheet.getRange('E23').setValue('🔵 Scaling Down (Spend %)').setFontWeight('bold').setWrap(false).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
  sheet.getRange('F23').setValue('🟡 Moderate (Spend/Profit %)').setFontWeight('bold').setWrap(false).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
  sheet.getRange('G23').setValue('⚪ Stable (%)').setFontWeight('bold').setWrap(false).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
  sheet.getRange('H23').setValue('Status').setFontWeight('bold').setWrap(false).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
  sheet.getRange('I23').setValue('Modified').setFontWeight('bold').setWrap(false).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
  sheet.getRange('A23:I23').setBackground('#f3e5f5');
  sheet.setRowHeight(23, 30);
  
  const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
  
  projects.forEach((proj, i) => {
    const row = 24 + i;
    sheet.getRange(`A${row}`).setValue(proj).setFontWeight('bold');
    
    sheet.getRange(`B${row}`).setValue('+10\n+5').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFontSize(12).setFontColor('#2e7d32').setFontWeight('bold').setBackground('#e8f5e9').setWrap(true);
    
    sheet.getRange(`C${row}`).setValue('-5\n+8').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFontSize(12).setFontColor('#1b5e20').setFontWeight('bold').setBackground('#e8f5e9').setWrap(true);
    
    sheet.getRange(`D${row}`).setValue('-8').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFontSize(14).setFontWeight('bold').setFontColor('#c62828').setBackground('#ffebee');
    
    sheet.getRange(`E${row}`).setValue('-15').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFontSize(14).setFontWeight('bold').setFontColor('#1565c0').setBackground('#e3f2fd');
    
    sheet.getRange(`F${row}`).setValue('±3\n±2').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFontSize(12).setFontColor('#f57f17').setFontWeight('bold').setBackground('#fff9c4').setWrap(true);
    
    sheet.getRange(`G${row}`).setValue('±2').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFontSize(14).setFontWeight('bold').setFontColor('#424242').setBackground('#f5f5f5');
    
    sheet.getRange(`H${row}`).setValue('✅ Active').setHorizontalAlignment('center').setFontColor('#28a745');
    sheet.getRange(`I${row}`).setValue('Default').setHorizontalAlignment('center').setFontStyle('italic');
    sheet.getRange(`A${row}:A${row}`).setBackground('#fce4ec');
    sheet.setRowHeight(row, 40);
    
    sheet.getRange(`A${row}:I${row}`).setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  });
  
  sheet.setRowHeight(32, 20);
  
  sheet.getRange('A33:I33').merge().setValue('📖 INSTRUCTIONS').setBackground('#607d8b').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(33, 30);
  
  sheet.getRange('A34').setValue('🤖 Automation Benefits:').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A35:I37').merge();
  sheet.getRange('A35').setValue(
    '• SEPARATE TRIGGERS: Each project updates independently\n' +
    '• NO TIMEOUTS: 30-minute intervals prevent API overload\n' +
    '• RELIABILITY: Failed project won\'t affect others'
  );
  sheet.getRange('A35:I37').setBackground('#f5f5f5').setWrap(true).setBorder(true, true, true, true, false, false);
  
  sheet.getRange('A39').setValue('🎯 How Target eROAS Works:').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A40:I42').merge();
  sheet.getRange('A40').setValue(
    '• TRICKY Project: Always 250% (entire sheet)\n' +
    '• Business Apps: 140% (apps with "Business" in name across all projects)\n' +
    '• Other Apps: 150% (default for everything else)'
  );
  sheet.getRange('A40:I42').setBackground('#f5f5f5').setWrap(true).setBorder(true, true, true, true, false, false);
  
  sheet.getRange('A44').setValue('📊 Understanding Growth Thresholds:').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A45:I53').merge();
  sheet.getRange('A45').setValue(
    '🎯 HOW TO USE:\n' +
    '• Just edit the numbers directly in the cells!\n' +
    '• Numbers represent percentage thresholds\n' +
    '• First line = Spend change %, Second line = Profit change %\n\n' +
    '🟢 HEALTHY GROWTH: Both metrics positive (e.g. +10/+5)\n' +
    '🟢 EFFICIENCY: Spend down, profit up (e.g. -5/+8)\n' +
    '🔴 WARNING: Profit drops significantly (e.g. -8)\n' +
    '🔵 SCALING DOWN: Major spend cut (e.g. -15)\n' +
    '🟡 MODERATE: Small changes (e.g. ±3/±2)\n' +
    '⚪ STABLE: Minimal change (e.g. ±2)\n\n' +
    '💡 TIP: After changing numbers, click Menu → 🔄 Refresh Settings'
  );
  sheet.getRange('A45:I53').setBackground('#f5f5f5').setWrap(true).setBorder(true, true, true, true, false, false).setVerticalAlignment('top');
  
  sheet.setRowHeight(55, 20);
  
  sheet.getRange('A56:I56').merge().setValue('🔐 API SETTINGS').setBackground('#4285f4').setFontColor('white').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sheet.setRowHeight(56, 30);
  
  sheet.getRange('A57').setValue('Bearer Token:').setFontWeight('bold');
  sheet.getRange('B57:I57').merge().setValue('[ENTER_YOUR_TOKEN_HERE]');
  sheet.getRange('A57:A57').setBackground('#e8f0fe');
  sheet.getRange('B57:I57').setBackground('#f8f9fa').setBorder(true, true, true, true, false, false);
  sheet.setRowHeight(57, 25);
  
  sheet.getRange('A59:I61').merge();
  sheet.getRange('A59').setValue(
    '• Get your token from: app.appodeal.com → Settings → API\n' +
    '• Token should start with "eyJ" and be 300+ characters long\n' +
    '• One token works for all projects'
  );
  sheet.getRange('A59:I61').setBackground('#f5f5f5').setWrap(true).setBorder(true, true, true, true, false, false);
  
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 210);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 170);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 90);
  sheet.setColumnWidth(8, 80);
  sheet.setColumnWidth(9, 90);
  
  sheet.getRange('B4:B5').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['TRUE', 'FALSE']).build());
  
  sheet.setFrozenRows(3);
}

function loadSettingsFromSheet() {
  const now = new Date().getTime();
  
  if (SETTINGS_CACHE && SETTINGS_CACHE_TIME && (now - SETTINGS_CACHE_TIME) < 60000) {
    return SETTINGS_CACHE;
  }
  
  const cacheKey = 'SETTINGS_CACHE_V3';
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);
  
  if (cachedData) {
    try {
      const cached = JSON.parse(cachedData);
      SETTINGS_CACHE = cached;
      SETTINGS_CACHE_TIME = now;
      return cached;
    } catch (e) {
      console.log('Error parsing cached settings:', e);
    }
  }
  
  let sheet;
  try {
    sheet = getOrCreateSettingsSheet();
  } catch (e) {
    console.error('Error getting settings sheet:', e);
    return getDefaultSettingsValues();
  }
  
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
    
    if (label === 'TRICKY Project:' && i >= 17 && i <= 21) {
      const numValue = parseInt(value);
      settings.targetEROAS.tricky = (!isNaN(numValue) && numValue >= 100 && numValue <= 500) ? numValue : 250;
    }
    
    if (label === 'Business Apps:' && i >= 17 && i <= 21) {
      const numValue = parseInt(value);
      settings.targetEROAS.business = (!isNaN(numValue) && numValue >= 100 && numValue <= 500) ? numValue : 140;
    }
    
    if (label === 'Other Apps:' && i >= 17 && i <= 21) {
      const numValue = parseInt(value);
      settings.targetEROAS.ceg = (!isNaN(numValue) && numValue >= 100 && numValue <= 500) ? numValue : 150;
    }
    
    const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    projects.forEach(proj => {
      if (label === proj && i >= 23 && i <= 32) {
        const healthyCell = row[1] ? row[1].toString() : '+10\n+5';
        const efficiencyCell = row[2] ? row[2].toString() : '-5\n+8';
        const warningCell = row[3] ? row[3].toString() : '-8';
        const scalingCell = row[4] ? row[4].toString() : '-15';
        const moderateCell = row[5] ? row[5].toString() : '±3\n±2';
        
        const healthyParts = healthyCell.split('\n');
        const healthySpend = parseInt(healthyParts[0].replace(/[^0-9-]/g, '')) || 10;
        const healthyProfit = parseInt(healthyParts[1]?.replace(/[^0-9-]/g, '')) || 5;
        
        const efficiencyParts = efficiencyCell.split('\n');
        const efficiencySpend = parseInt(efficiencyParts[0].replace(/[^0-9-]/g, '')) || -5;
        const efficiencyProfit = parseInt(efficiencyParts[1]?.replace(/[^0-9-]/g, '')) || 8;
        
        const warningProfit = parseInt(warningCell.replace(/[^0-9-]/g, '')) || -8;
        const scalingSpend = parseInt(scalingCell.replace(/[^0-9-]/g, '')) || -15;
        
        const moderateParts = moderateCell.split('\n');
        const moderateSpend = parseInt(moderateParts[0].replace(/[^0-9]/g, '')) || 3;
        const moderateProfit = parseInt(moderateParts[1]?.replace(/[^0-9]/g, '')) || 2;
        
        const stableValue = row[6] ? parseInt(row[6].toString().replace(/[^0-9]/g, '')) || 2 : 2;
        
        settings.growthThresholds[proj] = {
          healthyGrowth: { 
            minSpendChange: healthySpend, 
            minProfitChange: healthyProfit 
          },
          efficiencyImprovement: { 
            maxSpendDecline: efficiencySpend, 
            minProfitGrowth: efficiencyProfit 
          },
          inefficientGrowth: { 
            minSpendChange: 0, 
            maxProfitChange: warningProfit 
          },
          decliningEfficiency: { 
            minSpendStable: -2, 
            maxSpendGrowth: 10, 
            maxProfitDecline: -4, 
            minProfitDecline: -7 
          },
          scalingDown: { 
            maxSpendChange: scalingSpend,
            efficient: { minProfitChange: 0 },
            moderate: { 
              maxProfitDecline: -10, 
              minProfitDecline: -1 
            },
            problematic: { maxProfitDecline: -15 }
          },
          moderateGrowthSpend: moderateSpend,
          moderateGrowthProfit: moderateProfit,
          minimalGrowth: { maxSpendChange: 2, maxProfitChange: 1 },
          moderateDecline: { 
            maxSpendDecline: -3, maxProfitDecline: -3, spendOptimizationRatio: 1.5,
            efficiencyDropRatio: 1.5, proportionalRatio: 1.3
          },
          stable: { maxAbsoluteChange: stableValue }
        };
      }
    });
  }
  
  SETTINGS_CACHE = settings;
  SETTINGS_CACHE_TIME = now;
  
  try {
    cache.put(cacheKey, JSON.stringify(settings), 600);
  } catch (e) {
    console.log('Error caching settings:', e);
  }
  
  return settings;
}

function getDefaultSettingsValues() {
  return {
    bearerToken: '',
    targetEROAS: { tricky: 250, business: 140, ceg: 150 },
    automation: { autoCache: false, autoUpdate: false },
    growthThresholds: {
      TRICKY: getDefaultGrowthThresholds(),
      MOLOCO: getDefaultGrowthThresholds(),
      REGULAR: getDefaultGrowthThresholds(),
      GOOGLE_ADS: getDefaultGrowthThresholds(),
      APPLOVIN: getDefaultGrowthThresholds(),
      MINTEGRAL: getDefaultGrowthThresholds(),
      INCENT: getDefaultGrowthThresholds(),
      OVERALL: getDefaultGrowthThresholds()
    }
  };
}

function populateDefaultSettings(sheet) {
  try {
    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty('BEARER_TOKEN');
    if (token) {
      const tokenRow = findTokenRow(sheet);
      if (tokenRow > 0) {
        sheet.getRange(tokenRow, 2, 1, 8).merge().setValue(token);
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
  clearSettingsCache();
  
  const sheet = getOrCreateSettingsSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const label = row[0] ? row[0].toString().trim() : '';
    
    if (settingPath === 'bearerToken' && label === 'Bearer Token:') {
      sheet.getRange(i + 1, 2, 1, 8).merge().setValue(value);
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
    
    if (settingPath.startsWith('growthThresholds.')) {
      const parts = settingPath.split('.');
      const project = parts[1];
      
      if (label === project) {
        const thresholds = value;
        
        const healthyText = `+${thresholds.healthyGrowth.minSpendChange}\n+${thresholds.healthyGrowth.minProfitChange}`;
        const efficiencyText = `${thresholds.efficiencyImprovement.maxSpendDecline}\n+${thresholds.efficiencyImprovement.minProfitGrowth}`;
        const warningText = `${thresholds.inefficientGrowth.maxProfitChange}`;
        const scalingText = `${thresholds.scalingDown.maxSpendChange}`;
        const moderateText = `±${thresholds.moderateGrowthSpend}\n±${thresholds.moderateGrowthProfit}`;
        const stableText = `±${thresholds.stable.maxAbsoluteChange}`;
        
        sheet.getRange(i + 1, 2).setValue(healthyText);
        sheet.getRange(i + 1, 3).setValue(efficiencyText);
        sheet.getRange(i + 1, 4).setValue(warningText);
        sheet.getRange(i + 1, 5).setValue(scalingText);
        sheet.getRange(i + 1, 6).setValue(moderateText);
        sheet.getRange(i + 1, 7).setValue(stableText);
        return;
      }
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
  BEARER_TOKEN_CACHE = null;
  BEARER_TOKEN_CACHE_TIME = null;
  
  const cache = CacheService.getScriptCache();
  cache.remove('SETTINGS_CACHE_V3');
  cache.remove('BEARER_TOKEN');
}

function openSettingsSheet() {
  const sheet = getOrCreateSettingsSheet();
  const spreadsheet = SpreadsheetApp.openById(MAIN_SHEET_ID);
  spreadsheet.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('Settings Sheet', 'Settings sheet opened!\n\n🤖 NEW: Separate triggers prevent timeouts\n🎯 Fixed Target Logic:\n• TRICKY: 250% (entire project)\n• Business: 140% (apps with "Business")\n• Others: 150% (default)\n\nUse "🔄 Refresh Settings" after making changes.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function forceUpdateSettingsSheet() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('🔄 Force Update Settings', 'Force update the Settings sheet?\n\nThis will create the structure with separate triggers logic.', ui.ButtonSet.YES_NO);
  
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
    
    ui.alert('✅ Updated', 'Settings sheet updated with separate triggers!\n\n🤖 NEW FEATURES:\n• Separate triggers for each project\n• 30-minute intervals prevent timeouts\n• More reliable automation\n\n🎯 Schedule: 4:00-7:30 AM', ui.ButtonSet.OK);
  }
}