/**
 * Menu Functions - ОБНОВЛЕНО: добавлен Incent
 */

var MENU_PROJECTS = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent'];
var MENU_DAYS = [30, 60, 90];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📊 Campaign Report');
  var props = PropertiesService.getScriptProperties();
  
  var tokenStatus = isBearerTokenConfigured() ? '🔐✅' : '🔐❌';
  
  menu.addItem('📈 Generate Report...', 'smartReportWizard')
      .addItem('💾 Save All Comments', 'saveAllCommentsToCache')
      .addSeparator()
      .addItem(tokenStatus + ' Bearer Token...', 'showTokenSettings')
      .addSeparator()
      .addItem(props.getProperty('AUTO_CACHE_ENABLED') === 'true' ? '✅ Auto-Cache ON → Turn OFF' : '❌ Auto-Cache OFF → Turn ON', 'toggleAutoCache')
      .addItem(props.getProperty('AUTO_UPDATE_ENABLED') === 'true' ? '✅ Auto-Update ON → Turn OFF' : '❌ Auto-Update OFF → Turn ON', 'toggleAutoUpdate')
      .addSeparator()
      .addItem('⚙️ Settings & Tools...', 'smartSettingsHub')
      .addItem('🐙 GitHub Repository', 'openGitHubRepo')
      .addToUi();
}

function openGitHubRepo() {
  var ui = SpreadsheetApp.getUi();
  var githubUrl = 'https://github.com/Poxagronka/pivot-table';
  
  var htmlOutput = HtmlService.createHtmlOutput(
    '<script>window.open("' + githubUrl + '", "_blank"); google.script.host.close();</script>'
  ).setWidth(400).setHeight(300);
  
  ui.showModalDialog(htmlOutput, 'Opening GitHub Repository...');
}

function smartReportWizard() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    if (ui.alert('🔐 Token Required', 'Bearer token is not configured. Set it now?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
      updateBearerToken();
      return;
    } else {
      ui.alert('❌ Cannot Generate Reports', 'Bearer token is required for API access.', ui.ButtonSet.OK);
      return;
    }
  }
  
  var scope = showChoice('📈 Generate Report - Step 1/3', ['All Projects Together', 'Single Project', 'Custom Selection']);
  if (!scope) return;
  
  var period = showChoice('📅 Select Period - Step 2/3', ['Last 30 days', 'Last 60 days', 'Last 90 days', 'Custom days (120, 360, etc)', 'Date range (from/to)']);
  if (!period) return;
  
  if (scope === 1) {
    var days = period <= 3 ? MENU_DAYS[period-1] : period === 4 ? promptNumber('Enter days:', [120, 360]) : null;
    if (period === 5) {
      var dates = promptDateRange();
      if (dates) runAllProjectsDateRange(dates.start, dates.end);
    } else if (days) {
      quickGenerateAllForDays(days);
    }
  } else if (scope === 2) {
    var project = showChoice('Select Project - Step 3/3', MENU_PROJECTS);
    if (!project) return;
    var projectName = MENU_PROJECTS[project-1].toUpperCase();
    if (period <= 3) {
      generateProjectReport(projectName, MENU_DAYS[period-1]);
    } else if (period === 4) {
      var days = promptNumber('Enter days:', [120, 360]);
      if (days) generateProjectReport(projectName, days);
    } else {
      var dates = promptDateRange();
      if (dates) generateProjectReportForDateRange(projectName, dates.start, dates.end);
    }
  } else {
    var selected = showMultiChoice('Select Projects:', MENU_PROJECTS);
    if (!selected || selected.length === 0) return;
    var days = period <= 3 ? MENU_DAYS[period-1] : period === 4 ? promptNumber('Enter days:', [120, 360]) : null;
    if (days) {
      runSelectedProjects(selected, days);
    } else if (period === 5) {
      var dates = promptDateRange();
      if (dates) runSelectedProjectsDateRange(selected, dates.start, dates.end);
    }
  }
}

function smartSettingsHub() {
  var action = showChoice('⚙️ Settings & Tools', [
    '🔐 Bearer Token Settings',
    '🎯 Target eROAS Settings', 
    '📊 Growth Status Thresholds', 
    '📋 View Project Overview', 
    '💬 Comments Management', 
    '📱 Apps Database (TRICKY)', 
    '🗑️ Clear Data', 
    '🔍 API Health Check', 
    '🐛 Debug Tools', 
    '📊 View System Status'
  ]);
  if (!action) return;
  
  switch(action) {
    case 1: showTokenSettings(); break;
    case 2: targetSettingsWizard(); break;
    case 3: growthThresholdsWizard(); break;
    case 4: projectOverviewWizard(); break;
    case 5: commentsWizard(); break;
    case 6: appsDbWizard(); break;
    case 7: clearDataWizard(); break;
    case 8: apiCheckWizard(); break;
    case 9: debugWizard(); break;
    case 10: showAutomationStatus(); break;
  }
}

function appsDbWizard() {
  var ui = SpreadsheetApp.getUi();
  
  if (CURRENT_PROJECT !== 'TRICKY') {
    var switchResult = ui.alert('Apps Database - TRICKY Only', 
      'Apps Database is only used for TRICKY project.\n\nSwitch to TRICKY project now?', 
      ui.ButtonSet.YES_NO);
    
    if (switchResult !== ui.Button.YES) return;
    setCurrentProject('TRICKY');
  }
  
  var action = showChoice('📱 Apps Database Management', [
    'View Cache Status',
    'Refresh Apps Database', 
    'View Sample Data',
    'Clear Cache',
    'Debug Update Process'
  ]);
  if (!action) return;
  
  switch(action) {
    case 1: showAppsDbStatus(); break;
    case 2: refreshAppsDatabase(); break;
    case 3: showAppsDbSample(); break;
    case 4: clearAppsDbCache(); break;
    case 5: debugAppsDatabase(); break;
  }
}

function showAppsDbStatus() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var appsDb = new AppsDatabase('TRICKY');
    var cache = appsDb.loadFromCache();
    var appCount = Object.keys(cache).length;
    
    var message = '📱 APPS DATABASE STATUS\n\n';
    message += '• Total Apps: ' + appCount + '\n';
    
    if (appCount > 0) {
      var bundleIds = Object.keys(cache);
      var sampleApp = cache[bundleIds[0]];
      message += '• Last Updated: ' + (sampleApp.lastUpdated || 'Unknown') + '\n';
      message += '• Cache Sheet: ' + (appsDb.cacheSheet ? 'Found' : 'Missing') + '\n';
      
      var shouldUpdate = appsDb.shouldUpdateCache();
      message += '• Update Needed: ' + (shouldUpdate ? 'YES (>24h old)' : 'NO') + '\n\n';
      
      message += 'SAMPLE ENTRIES:\n';
      var sampleCount = Math.min(3, bundleIds.length);
      for (var i = 0; i < sampleCount; i++) {
        var bundleId = bundleIds[i];
        var app = cache[bundleId];
        message += '• ' + bundleId + ' → ' + app.publisher + ' ' + app.appName + '\n';
      }
      
      message += '\n💡 TIP: Use "Debug Update Process" for detailed logs';
    } else {
      message += '• Status: Empty cache\n';
      message += '• Action Required: Refresh database\n';
      message += '• Debug Tip: Use "Debug Update Process" for detailed logs\n';
      
      message += '\n🔍 DIAGNOSTIC INFO:\n';
      try {
        var externalSpreadsheet = SpreadsheetApp.openById(APPS_DATABASE_ID);
        var externalSheet = externalSpreadsheet.getSheetByName(APPS_DATABASE_SHEET);
        if (externalSheet) {
          var dataCount = externalSheet.getLastRow() - 1;
          message += '• External table rows: ' + dataCount + '\n';
          
          var headers = externalSheet.getRange(1, 1, 1, externalSheet.getLastColumn()).getValues()[0];
          var linkAppCol = -1;
          for (var j = 0; j < headers.length; j++) {
            if (headers[j].toString().toLowerCase() === 'link app') {
              linkAppCol = j;
              break;
            }
          }
          message += '• Link App column: ' + (linkAppCol !== -1 ? 'Found' : 'NOT FOUND') + '\n';
        } else {
          message += '• External sheet: NOT FOUND\n';
        }
      } catch (e) {
        message += '• External table: ERROR (' + e.toString().substring(0, 50) + '...)\n';
      }
    }
    
    ui.alert('Apps Database Status', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error checking Apps Database status: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function showAppsDbSample() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var appsDb = new AppsDatabase('TRICKY');
    var cache = appsDb.loadFromCache();
    var bundleIds = Object.keys(cache);
    
    if (bundleIds.length === 0) {
      ui.alert('No Data', 'Apps Database cache is empty. Please refresh the database first.', ui.ButtonSet.OK);
      return;
    }
    
    var message = '📱 APPS DATABASE SAMPLE DATA\n\n';
    var sampleCount = Math.min(10, bundleIds.length);
    
    for (var i = 0; i < sampleCount; i++) {
      var bundleId = bundleIds[i];
      var app = cache[bundleId];
      message += bundleId + '\n  → ' + app.publisher + ' ' + app.appName + '\n\n';
    }
    
    if (bundleIds.length > sampleCount) {
      message += '... and ' + (bundleIds.length - sampleCount) + ' more apps';
    }
    
    ui.alert('Apps Database Sample', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error showing sample data: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function clearAppsDbCache() {
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.alert('Clear Apps Database Cache', 
    'This will clear all cached app data for TRICKY project.\n\nThe cache will be automatically rebuilt on next refresh.\n\nContinue?', 
    ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  try {
    var appsDb = new AppsDatabase('TRICKY');
    if (appsDb.cacheSheet) {
      if (appsDb.cacheSheet.getLastRow() > 1) {
        appsDb.cacheSheet.deleteRows(2, appsDb.cacheSheet.getLastRow() - 1);
      }
      ui.alert('Success', 'Apps Database cache cleared successfully.', ui.ButtonSet.OK);
    } else {
      ui.alert('No Cache', 'Apps Database cache sheet not found.', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', 'Error clearing cache: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function showTokenSettings() {
  var ui = SpreadsheetApp.getUi();
  var isConfigured = isBearerTokenConfigured();
  
  var message = '🔐 BEARER TOKEN SETTINGS\n\n';
  if (isConfigured) {
    message += '✅ Token: Configured\n• Length: ' + getBearerToken().length + ' chars\n• Status: Valid\n\n';
    message += 'Options:\n1. Update Token\n2. Clear Token\n3. Test Token\n4. Cancel';
    
    var choice = showChoice('Token Settings', ['Update Token', 'Clear Token', 'Test Token']);
    if (!choice) return;
    
    if (choice === 1) {
      updateBearerToken();
    } else if (choice === 2) {
      if (ui.alert('Clear Token?', 'This will remove the bearer token. All API calls will fail until you set a new token.\n\nContinue?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
        clearBearerToken();
        ui.alert('✅ Token Cleared', 'Bearer token has been removed.', ui.ButtonSet.OK);
      }
    } else if (choice === 3) {
      testBearerToken();
    }
  } else {
    message += '❌ Token: NOT CONFIGURED\n• All API calls will fail\n• Please set token to use the system\n\n';
    message += 'Set Bearer Token now?';
    
    if (ui.alert('Token Required', message, ui.ButtonSet.YES_NO) === ui.Button.YES) {
      updateBearerToken();
    }
  }
}

function updateBearerToken() {
  var ui = SpreadsheetApp.getUi();
  var current = isBearerTokenConfigured() ? 'Current token length: ' + getBearerToken().length + ' chars\n\n' : '';
  
  var result = ui.prompt('🔐 Set Bearer Token', 
    current + 'Enter your Appodeal API Bearer Token:\n\n• Get it from app.appodeal.com\n• Token should start with "eyJ"\n• Must be at least 50 characters', 
    ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK) {
    var token = result.getResponseText().trim();
    
    try {
      setBearerToken(token);
      ui.alert('✅ Token Saved', 'Bearer token has been successfully saved and is ready to use.', ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('❌ Invalid Token', 'Error: ' + e.toString() + '\n\nPlease check:\n• Token is complete\n• No extra spaces\n• Minimum 50 characters', ui.ButtonSet.OK);
    }
  }
}

function testBearerToken() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    ui.alert('❌ No Token', 'Bearer token is not configured. Please set it first.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    setCurrentProject('TRICKY');
    var dateRange = getDateRange(7);
    var raw = fetchCampaignData(dateRange);
    
    if (raw && raw.data) {
      var recordCount = raw.data.analytics?.richStats?.stats?.length || 0;
      ui.alert('✅ Token Valid', 'Bearer token is working correctly!\n\n• API connection: Success\n• Test records: ' + recordCount + '\n• Token length: ' + getBearerToken().length + ' chars', ui.ButtonSet.OK);
    } else {
      ui.alert('⚠️ Token Issues', 'Token accepted but no data returned.\n\nPossible causes:\n• Token has limited permissions\n• No recent campaign data\n• Network configuration issues', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('❌ Token Failed', 'Bearer token test failed:\n\n' + e.toString() + '\n\nPlease check:\n• Token is current and valid\n• Has proper API permissions\n• Network connectivity', ui.ButtonSet.OK);
  }
}

function targetSettingsWizard() {
  var choice = showChoice('🎯 Target eROAS Settings', ['View Current Settings', 'Update Single Project', 'Update All Projects', 'Reset to Defaults']);
  if (!choice) return;
  var ui = SpreadsheetApp.getUi();
  
  if (choice === 1) {
    var message = 'Current Target eROAS:\n';
    for (var i = 0; i < MENU_PROJECTS.length; i++) {
      var p = MENU_PROJECTS[i];
      message += p + ': ' + getTargetEROAS(p.toUpperCase()) + '%\n';
    }
    ui.alert('Current Target eROAS', message, ui.ButtonSet.OK);
  } else if (choice === 2) {
    var project = showChoice('Select Project:', MENU_PROJECTS);
    if (project) {
      var p = MENU_PROJECTS[project-1];
      var current = getTargetEROAS(p.toUpperCase());
      var value = promptNumber(p + ' Target eROAS (current: ' + current + '%):' , [140, 160, 180]);
      if (value && value >= 100 && value <= 500) {
        setTargetEROAS(p.toUpperCase(), value);
        ui.alert('✅ Updated', p + ' target set to ' + value + '%', ui.ButtonSet.OK);
      }
    }
  } else if (choice === 3) {
    var values = {};
    for (var i = 0; i < MENU_PROJECTS.length; i++) {
      var p = MENU_PROJECTS[i];
      var current = getTargetEROAS(p.toUpperCase());
      var value = promptNumber(p + ' (current: ' + current + '%):' , [current]);
      if (value && value >= 100 && value <= 500) values[p] = value;
    }
    var keys = Object.keys(values);
    if (keys.length > 0) {
      for (var i = 0; i < keys.length; i++) {
        setTargetEROAS(keys[i].toUpperCase(), values[keys[i]]);
      }
      ui.alert('✅ Updated', 'All targets have been saved', ui.ButtonSet.OK);
    }
  } else if (choice === 4) {
    if (ui.alert('Reset to Defaults?', 'Tricky: 160%\nMoloco: 140%\nRegular: 140%\nGoogle_Ads: 140%\nApplovin: 140%\nMintegral: 140%\nIncent: 140%', ui.ButtonSet.YES_NO) === ui.Button.YES) {
      setTargetEROAS('TRICKY', 160);
      setTargetEROAS('MOLOCO', 140);
      setTargetEROAS('REGULAR', 140);
      setTargetEROAS('GOOGLE_ADS', 140);
      setTargetEROAS('APPLOVIN', 140);
      setTargetEROAS('MINTEGRAL', 140);
      setTargetEROAS('INCENT', 140);
      ui.alert('✅ Reset', 'All targets reset to defaults', ui.ButtonSet.OK);
    }
  }
}

function projectOverviewWizard() {
  var choice = showChoice('📋 Project Overview', ['View All Projects Summary', 'View Single Project Details', 'Compare Growth Thresholds', 'Export Settings']);
  if (!choice) return;
  var ui = SpreadsheetApp.getUi();
  
  if (choice === 1) {
    showAllProjectsOverview();
  } else if (choice === 2) {
    var project = showChoice('Select Project:', MENU_PROJECTS);
    if (project) {
      var projectName = MENU_PROJECTS[project-1].toUpperCase();
      var overview = getProjectStatusOverview(projectName);
      ui.alert(MENU_PROJECTS[project-1] + ' Overview', overview, ui.ButtonSet.OK);
    }
  } else if (choice === 3) {
    showThresholdsComparison();
  } else if (choice === 4) {
    ui.alert('Export Settings', 'Settings export feature coming soon!', ui.ButtonSet.OK);
  }
}

function getProjectStatusOverview(projectName) {
  var target = getTargetEROAS(projectName);
  var thresholds = getGrowthThresholds(projectName);
  var config = getProjectConfig(projectName);
  var apiConfig = getProjectApiConfig(projectName);
  
  var tokenStatus = isBearerTokenConfigured() ? '✅ Configured' : '❌ Missing';
  
  var overview = '📋 ' + projectName + ' OVERVIEW\n\n';
  overview += '🔐 Bearer Token: ' + tokenStatus + '\n';
  overview += '🎯 Target eROAS: ' + target + '%\n';
  overview += '📊 Sheet: ' + config.SHEET_NAME + '\n';
  overview += '🌐 Network HID: ' + apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.join(', ') + '\n';
  overview += '🔍 Campaign Filter: ' + (apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH || 'NO FILTER') + '\n';
  overview += '👥 Users: ' + apiConfig.FILTERS.USER.length + ' configured\n';
  
  if (projectName === 'TRICKY') {
    try {
      var appsDb = new AppsDatabase('TRICKY');
      var cache = appsDb.loadFromCache();
      var appCount = Object.keys(cache).length;
      overview += '📱 Apps Database: ' + appCount + ' apps cached\n';
    } catch (e) {
      overview += '📱 Apps Database: Error\n';
    }
  }
  
  overview += '\n📈 GROWTH THRESHOLDS:\n';
  overview += '🟢 Healthy: Spend >' + thresholds.healthyGrowth.minSpendChange + '%, Profit >' + thresholds.healthyGrowth.minProfitChange + '%\n';
  overview += '🔴 Inefficient: Profit <' + thresholds.inefficientGrowth.maxProfitChange + '%\n';
  overview += '🔵 Scaling: Spend <' + thresholds.scalingDown.maxSpendChange + '%\n';
  
  return overview;
}

function growthThresholdsWizard() {
  var choice = showChoice('📊 Growth Status Thresholds', ['Quick View Current Settings', 'Update Basic Thresholds', 'Reset to Defaults', 'View Growth Criteria Guide']);
  if (!choice) return;
  
  switch(choice) {
    case 1: viewCurrentThresholds(); break;
    case 2: updateBasicThresholds(); break;
    case 3: resetAllThresholdsToDefaults(); break;
    case 4: showGrowthCriteriaGuide(); break;
  }
}

function updateBasicThresholds() {
  var ui = SpreadsheetApp.getUi();
  var project = showChoice('Select Project:', MENU_PROJECTS);
  if (!project) return;
  var projectName = MENU_PROJECTS[project-1].toUpperCase();
  
  var current;
  try {
    current = getGrowthThresholds(projectName);
  } catch (e) {
    current = { healthyGrowth: { minSpendChange: 10, minProfitChange: 5 }, inefficientGrowth: { maxProfitChange: -8 }, scalingDown: { maxSpendChange: -15 }, moderateGrowthSpend: 3, moderateGrowthProfit: 2 };
  }
  
  if (!current.healthyGrowth) current.healthyGrowth = { minSpendChange: 10, minProfitChange: 5 };
  if (!current.inefficientGrowth) current.inefficientGrowth = { maxProfitChange: -8 };
  if (!current.scalingDown) current.scalingDown = { maxSpendChange: -15 };
  if (!current.moderateGrowthSpend) current.moderateGrowthSpend = 3;
  if (!current.moderateGrowthProfit) current.moderateGrowthProfit = 2;
  
  var currentInfo = '📊 ' + MENU_PROJECTS[project-1] + ' Current Settings:\n\n🟢 Healthy Growth: Spend >' + current.healthyGrowth.minSpendChange + '%, Profit >' + current.healthyGrowth.minProfitChange + '%\n🔴 Inefficient: Profit <' + current.inefficientGrowth.maxProfitChange + '%\n🔵 Scaling: Spend <' + current.scalingDown.maxSpendChange + '%\n🟡 Moderate: Spend >' + current.moderateGrowthSpend + '%, Profit >' + current.moderateGrowthProfit + '%';
  ui.alert('Current Settings', currentInfo, ui.ButtonSet.OK);
  
  var category = showChoice('What to Update?', ['🟢 Healthy Growth Thresholds', '🔴 Inefficient Growth Threshold', '🔵 Scaling Down Threshold', '🟡 Moderate Growth Thresholds']);
  if (!category) return;
  
  var newThresholds = JSON.parse(JSON.stringify(current));
  
  if (category === 1) {
    var spendInput = ui.prompt('🟢 Healthy Growth - Spend Threshold', 'Current: ' + current.healthyGrowth.minSpendChange + '%\n\nEnter minimum spend increase % for healthy growth:', ui.ButtonSet.OK_CANCEL);
    if (spendInput.getSelectedButton() === ui.Button.OK) {
      var spendValue = parseInt(spendInput.getResponseText());
      if (!isNaN(spendValue) && spendValue >= 0 && spendValue <= 100) newThresholds.healthyGrowth.minSpendChange = spendValue;
    }
    var profitInput = ui.prompt('🟢 Healthy Growth - Profit Threshold', 'Current: ' + current.healthyGrowth.minProfitChange + '%\n\nEnter minimum profit increase % for healthy growth:', ui.ButtonSet.OK_CANCEL);
    if (profitInput.getSelectedButton() === ui.Button.OK) {
      var profitValue = parseInt(profitInput.getResponseText());
      if (!isNaN(profitValue) && profitValue >= -50 && profitValue <= 100) newThresholds.healthyGrowth.minProfitChange = profitValue;
    }
  } else if (category === 2) {
    var profitInput = ui.prompt('🔴 Inefficient Growth - Profit Threshold', 'Current: ' + current.inefficientGrowth.maxProfitChange + '%\n\nEnter maximum profit decline % before marking as inefficient:\n(Use negative numbers, e.g., -10 for 10% decline)', ui.ButtonSet.OK_CANCEL);
    if (profitInput.getSelectedButton() === ui.Button.OK) {
      var profitValue = parseInt(profitInput.getResponseText());
      if (!isNaN(profitValue) && profitValue <= 0 && profitValue >= -100) newThresholds.inefficientGrowth.maxProfitChange = profitValue;
    }
  } else if (category === 3) {
    var spendInput = ui.prompt('🔵 Scaling Down - Spend Threshold', 'Current: ' + current.scalingDown.maxSpendChange + '%\n\nEnter maximum spend decline % before marking as scaling down:\n(Use negative numbers, e.g., -20 for 20% decline)', ui.ButtonSet.OK_CANCEL);
    if (spendInput.getSelectedButton() === ui.Button.OK) {
      var spendValue = parseInt(spendInput.getResponseText());
      if (!isNaN(spendValue) && spendValue <= 0 && spendValue >= -100) newThresholds.scalingDown.maxSpendChange = spendValue;
    }
  } else if (category === 4) {
    var spendInput = ui.prompt('🟡 Moderate Growth - Spend Threshold', 'Current: ' + current.moderateGrowthSpend + '%\n\nEnter minimum spend increase % for moderate growth:', ui.ButtonSet.OK_CANCEL);
    if (spendInput.getSelectedButton() === ui.Button.OK) {
      var spendValue = parseInt(spendInput.getResponseText());
      if (!isNaN(spendValue) && spendValue >= 0 && spendValue <= 50) newThresholds.moderateGrowthSpend = spendValue;
    }
    var profitInput = ui.prompt('🟡 Moderate Growth - Profit Threshold', 'Current: ' + current.moderateGrowthProfit + '%\n\nEnter minimum profit increase % for moderate growth:', ui.ButtonSet.OK_CANCEL);
    if (profitInput.getSelectedButton() === ui.Button.OK) {
      var profitValue = parseInt(profitInput.getResponseText());
      if (!isNaN(profitValue) && profitValue >= 0 && profitValue <= 50) newThresholds.moderateGrowthProfit = profitValue;
    }
  }
  
  setGrowthThresholds(projectName, newThresholds);
  ui.alert('✅ Updated', MENU_PROJECTS[project-1] + ' thresholds updated!', ui.ButtonSet.OK);
}

function viewCurrentThresholds() {
  var ui = SpreadsheetApp.getUi();
  var message = '📊 CURRENT GROWTH THRESHOLDS\n\n';
  for (var i = 0; i < MENU_PROJECTS.length; i++) {
    var project = MENU_PROJECTS[i];
    var projectName = project.toUpperCase();
    try {
      var thresholds = getGrowthThresholds(projectName);
      var healthySpend = thresholds.healthyGrowth ? thresholds.healthyGrowth.minSpendChange : 10;
      var healthyProfit = thresholds.healthyGrowth ? thresholds.healthyGrowth.minProfitChange : 5;
      var inefficientProfit = thresholds.inefficientGrowth ? thresholds.inefficientGrowth.maxProfitChange : -8;
      var scalingSpend = thresholds.scalingDown ? thresholds.scalingDown.maxSpendChange : -15;
      var moderateSpend = thresholds.moderateGrowthSpend || 3;
      var moderateProfit = thresholds.moderateGrowthProfit || 2;
      message += project + ':\n🟢 Healthy: Spend >' + healthySpend + '%, Profit >' + healthyProfit + '%\n🔴 Inefficient: Profit <' + inefficientProfit + '%\n🔵 Scaling: Spend <' + scalingSpend + '%\n🟡 Moderate: Spend >' + moderateSpend + '%, Profit >' + moderateProfit + '%\n\n';
    } catch (e) {
      message += project + ': ERROR - ' + e.toString() + '\n\n';
    }
  }
  ui.alert('Growth Thresholds', message, ui.ButtonSet.OK);
}

function resetAllThresholdsToDefaults() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Reset to Defaults?', 'This will reset all growth thresholds to their default values.\n\nContinue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  for (var i = 0; i < MENU_PROJECTS.length; i++) resetGrowthThresholds(MENU_PROJECTS[i].toUpperCase());
  ui.alert('✅ Reset', 'All growth thresholds have been reset to defaults!', ui.ButtonSet.OK);
}

function showAllProjectsOverview() {
  var ui = SpreadsheetApp.getUi();
  var tokenStatus = isBearerTokenConfigured() ? '✅ Configured' : '❌ Missing';
  var message = '📋 ALL PROJECTS OVERVIEW\n\n🔐 Bearer Token: ' + tokenStatus + '\n\n';
  for (var i = 0; i < MENU_PROJECTS.length; i++) {
    var project = MENU_PROJECTS[i];
    var projectName = project.toUpperCase();
    var targetROAS = getTargetEROAS(projectName);
    try {
      var thresholds = getGrowthThresholds(projectName);
      var healthySpend = thresholds.healthyGrowth ? thresholds.healthyGrowth.minSpendChange : 10;
      var healthyProfit = thresholds.healthyGrowth ? thresholds.healthyGrowth.minProfitChange : 5;
      message += project + ': eROAS ' + targetROAS + '%, Healthy ' + healthySpend + '%/' + healthyProfit + '%\n';
    } catch (e) {
      message += project + ': eROAS ' + targetROAS + '%, Thresholds: ERROR\n';
    }
  }
  message += '\nClick "View Single Project Details" for full settings.';
  ui.alert('Projects Overview', message, ui.ButtonSet.OK);
}

function showThresholdsComparison() {
  var ui = SpreadsheetApp.getUi();
  var message = '📊 THRESHOLDS COMPARISON\n\n';
  var categories = ['Healthy Growth', 'Inefficient Growth', 'Scaling Down'];
  for (var c = 0; c < categories.length; c++) {
    var category = categories[c];
    message += category.toUpperCase() + ':\n';
    for (var i = 0; i < MENU_PROJECTS.length; i++) {
      var project = MENU_PROJECTS[i];
      try {
        var thresholds = getGrowthThresholds(project.toUpperCase());
        if (category === 'Healthy Growth' && thresholds.healthyGrowth) {
          message += project + ': ' + thresholds.healthyGrowth.minSpendChange + '%/' + thresholds.healthyGrowth.minProfitChange + '%\n';
        } else if (category === 'Inefficient Growth' && thresholds.inefficientGrowth) {
          message += project + ': ' + thresholds.inefficientGrowth.maxProfitChange + '%\n';
        } else if (category === 'Scaling Down' && thresholds.scalingDown) {
          message += project + ': ' + thresholds.scalingDown.maxSpendChange + '%\n';
        } else {
          message += project + ': N/A\n';
        }
      } catch (e) {
        message += project + ': ERROR\n';
      }
    }
    message += '\n';
  }
  ui.alert('Thresholds Comparison', message, ui.ButtonSet.OK);
}

function showGrowthCriteriaGuide() {
  var project = showChoice('Select Project for Guide:', MENU_PROJECTS);
  if (!project) return;
  var projectName = MENU_PROJECTS[project-1].toUpperCase();
  var explanation = getProjectGrowthStatusExplanation(projectName);
  SpreadsheetApp.getUi().alert('Growth Criteria Guide - ' + MENU_PROJECTS[project-1], explanation, SpreadsheetApp.getUi().ButtonSet.OK);
}

function commentsWizard() {
  var choice = showChoice('💬 Comments Management', ['Save All Comments Now', 'Save Single Project', 'View Auto-Cache Status', 'Configure Auto-Cache']);
  if (!choice) return;
  
  switch(choice) {
    case 1: saveAllCommentsToCache(); break;
    case 2: 
      var p = showChoice('Select Project:', MENU_PROJECTS);
      if (p) {
        var projectName = MENU_PROJECTS[p-1].toUpperCase();
        setCurrentProject(projectName);
        saveProjectCommentsManual(projectName);
        SpreadsheetApp.getUi().alert('✅ Success', MENU_PROJECTS[p-1] + ' comments saved', SpreadsheetApp.getUi().ButtonSet.OK);
      }
      break;
    case 3: showAutomationStatus(); break;
    case 4: showAutoCacheSettings(); break;
  }
}

function clearDataWizard() {
  var choice = showChoice('🗑️ Clear Data', ['Clear All Projects', 'Clear Single Project', 'View What Will Be Cleared']);
  if (!choice) return;
  
  if (choice === 1) {
    clearAllProjectsData();
  } else if (choice === 2) {
    var p = showChoice('Select Project:', MENU_PROJECTS);
    if (p) clearProjectAllData(MENU_PROJECTS[p-1].toUpperCase());
  } else {
    SpreadsheetApp.getUi().alert('Info', 'Clear Data will:\n\n✓ Remove all report data\n✓ Preserve saved comments\n✓ Keep your settings\n\nComments can be restored after clearing.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function apiCheckWizard() {
  var choice = showChoice('🔍 API Health Check', ['Quick Check All Projects', 'Check Single Project', 'Test with Custom Date Range']);
  if (!choice) return;
  
  if (choice === 1) {
    checkAllProjectsAPI();
  } else if (choice === 2) {
    var p = showChoice('Select Project:', MENU_PROJECTS);
    if (p) checkProjectAPI(MENU_PROJECTS[p-1].toUpperCase());
  } else {
    var dates = promptDateRange();
    if (dates) testAPIWithDateRange(dates.start, dates.end);
  }
}

function debugWizard() {
  var p = showChoice('🐛 Debug Tools', MENU_PROJECTS);
  if (p) debugProjectReportGeneration(MENU_PROJECTS[p-1].toUpperCase());
}

function showChoice(title, options) {
  var ui = SpreadsheetApp.getUi();
  var numbered = '';
  for (var i = 0; i < options.length; i++) numbered += (i + 1) + ' - ' + options[i] + '\n';
  var result = ui.prompt(title, numbered + '\nEnter number:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  var choice = parseInt(result.getResponseText());
  return (choice >= 1 && choice <= options.length) ? choice : null;
}

function showMultiChoice(title, options) {
  var ui = SpreadsheetApp.getUi();
  var numbered = '';
  for (var i = 0; i < options.length; i++) numbered += (i + 1) + ' - ' + options[i] + '\n';
  var result = ui.prompt(title, numbered + '\nEnter numbers separated by commas (e.g., 1,3):', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  var selections = result.getResponseText().split(',');
  var validSelections = [];
  for (var i = 0; i < selections.length; i++) {
    var n = parseInt(selections[i].trim());
    if (n >= 1 && n <= options.length) validSelections.push(options[n-1]);
  }
  return validSelections;
}

function promptNumber(prompt, suggestions) {
  suggestions = suggestions || [];
  var ui = SpreadsheetApp.getUi();
  var hint = suggestions.length > 0 ? ' (e.g., ' + suggestions.join(', ') + ')' : '';
  var result = ui.prompt('Input Required', prompt + hint, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  var num = parseInt(result.getResponseText());
  return isNaN(num) ? null : num;
}

function promptDateRange() {
  var ui = SpreadsheetApp.getUi();
  var start = ui.prompt('Start Date', 'Enter date (YYYY-MM-DD):\n\nExample: 2024-01-01', ui.ButtonSet.OK_CANCEL);
  if (start.getSelectedButton() !== ui.Button.OK) return null;
  var end = ui.prompt('End Date', 'Enter date (YYYY-MM-DD):\n\nExample: 2024-12-31', ui.ButtonSet.OK_CANCEL);
  if (end.getSelectedButton() !== ui.Button.OK) return null;
  if (!isValidDate(start.getResponseText()) || !isValidDate(end.getResponseText())) {
    ui.alert('❌ Invalid date format');
    return null;
  }
  return { start: start.getResponseText(), end: end.getResponseText() };
}

function quickGenerateAllForDays(days) {
  var ui = SpreadsheetApp.getUi();
  var success = 0;
  
  try {
    for (var i = 0; i < MENU_PROJECTS.length; i++) {
      var p = MENU_PROJECTS[i];
      try { 
        generateProjectReport(p.toUpperCase(), days); 
        success++; 
      } catch(e) { 
        console.error(e); 
      }
    }
    ui.alert('✅ Complete', 'Generated ' + success + '/' + MENU_PROJECTS.length + ' reports', ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Error', e.toString(), ui.ButtonSet.OK);
  }
}

function runSelectedProjects(projects, days) {
  for (var i = 0; i < projects.length; i++) {
    generateProjectReport(projects[i].toUpperCase(), days);
  }
  SpreadsheetApp.getUi().alert('✅ Complete', 'Generated ' + projects.length + ' reports', SpreadsheetApp.getUi().ButtonSet.OK);
}

function runAllProjectsDateRange(start, end) {
  for (var i = 0; i < MENU_PROJECTS.length; i++) {
    generateProjectReportForDateRange(MENU_PROJECTS[i].toUpperCase(), start, end);
  }
  SpreadsheetApp.getUi().alert('✅ Complete', 'All reports generated', SpreadsheetApp.getUi().ButtonSet.OK);
}

function runSelectedProjectsDateRange(projects, start, end) {
  for (var i = 0; i < projects.length; i++) {
    generateProjectReportForDateRange(projects[i].toUpperCase(), start, end);
  }
  SpreadsheetApp.getUi().alert('✅ Complete', 'Generated ' + projects.length + ' reports', SpreadsheetApp.getUi().ButtonSet.OK);
}

function toggleAutoCache() {
  var props = PropertiesService.getScriptProperties();
  var isOn = props.getProperty('AUTO_CACHE_ENABLED') === 'true';
  isOn ? disableAutoCache() : enableAutoCache();
}

function toggleAutoUpdate() {
  var props = PropertiesService.getScriptProperties();
  var isOn = props.getProperty('AUTO_UPDATE_ENABLED') === 'true';
  isOn ? disableAutoUpdate() : enableAutoUpdate();
}

function generateProjectReport(projectName, days) { setCurrentProject(projectName); generateReport(days); }
function generateProjectReportForDateRange(projectName, startDate, endDate) { setCurrentProject(projectName); generateReportForDateRange(startDate, endDate); }
function debugProjectReportGeneration(projectName) { setCurrentProject(projectName); debugReportGeneration(); }
function isValidDate(dateString) { var regex = /^\d{4}-\d{2}-\d{2}$/; if (!regex.test(dateString)) return false; var date = new Date(dateString); return date instanceof Date && !isNaN(date); }

function generateReport30() { generateProjectReport('TRICKY', 30); }
function generateReport60() { generateProjectReport('TRICKY', 60); }
function generateReport90() { generateProjectReport('TRICKY', 90); }
function saveCommentsToCache() { setCurrentProject('TRICKY'); saveProjectCommentsManual('TRICKY'); }
function showDaysDialog() { smartReportWizard(); }
function showDateRangeDialog() { smartReportWizard(); }
function clearAllData() { clearProjectAllData('TRICKY'); }