/**
 * Menu Functions - СОКРАЩЕННАЯ ВЕРСИЯ - ОБНОВЛЕНО: добавлен Applovin
 */

var MENU_PROJECTS = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin'];
var MENU_DAYS = [30, 60, 90];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📊 Campaign Report');
  var props = PropertiesService.getScriptProperties();
  
  menu.addItem('📈 Generate Report...', 'smartReportWizard')
      .addItem('💾 Save All Comments', 'saveAllCommentsToCache')
      .addSeparator()
      .addItem(props.getProperty('AUTO_CACHE_ENABLED') === 'true' ? '✅ Auto-Cache ON → Turn OFF' : '❌ Auto-Cache OFF → Turn ON', 'toggleAutoCache')
      .addItem(props.getProperty('AUTO_UPDATE_ENABLED') === 'true' ? '✅ Auto-Update ON → Turn OFF' : '❌ Auto-Update OFF → Turn ON', 'toggleAutoUpdate')
      .addSeparator()
      .addItem('⚙️ Settings & Tools...', 'smartSettingsHub')
      .addToUi();
}

function smartReportWizard() {
  var ui = SpreadsheetApp.getUi();
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
  var action = showChoice('⚙️ Settings & Tools', ['🎯 Target eROAS Settings', '📊 Growth Status Thresholds', '📋 View Project Overview', '💬 Comments Management', '🗑️ Clear Data', '🔍 API Health Check', '🐛 Debug Tools', '📊 View System Status']);
  if (!action) return;
  
  switch(action) {
    case 1: targetSettingsWizard(); break;
    case 2: growthThresholdsWizard(); break;
    case 3: projectOverviewWizard(); break;
    case 4: commentsWizard(); break;
    case 5: clearDataWizard(); break;
    case 6: apiCheckWizard(); break;
    case 7: debugWizard(); break;
    case 8: showAutomationStatus(); break;
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
    if (ui.alert('Reset to Defaults?', 'Tricky: 160%\nMoloco: 140%\nRegular: 140%\nGoogle_Ads: 140%\nApplovin: 140%', ui.ButtonSet.YES_NO) === ui.Button.YES) {
      setTargetEROAS('TRICKY', 160);
      setTargetEROAS('MOLOCO', 140);
      setTargetEROAS('REGULAR', 140);
      setTargetEROAS('GOOGLE_ADS', 140);
      setTargetEROAS('APPLOVIN', 140);
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
  ui.alert('✅ Updated', MENU_PROJECTS[project-1] + ' thresholds have been updated!', ui.ButtonSet.OK);
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
  var message = '📋 ALL PROJECTS OVERVIEW\n\n';
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

// Helper UI Functions
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

// Quick Actions
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

// Toggle functions
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

// Core functions
function generateProjectReport(projectName, days) { setCurrentProject(projectName); generateReport(days); }
function generateProjectReportForDateRange(projectName, startDate, endDate) { setCurrentProject(projectName); generateReportForDateRange(startDate, endDate); }
function debugProjectReportGeneration(projectName) { setCurrentProject(projectName); debugReportGeneration(); }
function isValidDate(dateString) { var regex = /^\d{4}-\d{2}-\d{2}$/; if (!regex.test(dateString)) return false; var date = new Date(dateString); return date instanceof Date && !isNaN(date); }

// Legacy support functions
function generateReport30() { generateProjectReport('TRICKY', 30); }
function generateReport60() { generateProjectReport('TRICKY', 60); }
function generateReport90() { generateProjectReport('TRICKY', 90); }
function saveCommentsToCache() { setCurrentProject('TRICKY'); saveProjectCommentsManual('TRICKY'); }
function showDaysDialog() { smartReportWizard(); }
function showDateRangeDialog() { smartReportWizard(); }
function clearAllData() { clearProjectAllData('TRICKY'); }