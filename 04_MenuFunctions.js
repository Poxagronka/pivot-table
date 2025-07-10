/**
 * Menu Functions - ОБНОВЛЕНО: описания ежедневных триггеров (кеширование 3 утра, обновление 5 утра)
 */

var MENU_PROJECTS = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Overall'];
var MENU_DAYS = [30, 60, 90];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📊 Campaign Report');
  
  menu.addItem('📈 Generate Report...', 'smartReportWizard')
      .addItem('🔄 Update All to Current', 'updateAllProjectsToCurrent')
      .addItem('🎯 Update Selected Projects', 'updateSelectedProjectsToCurrent')
      .addItem('💾 Save All Comments', 'saveAllCommentsToCache')
      .addSeparator()
      .addItem('⚙️ Open Settings Sheet', 'openSettingsSheet')
      .addItem('🔄 Refresh Settings', 'refreshSettingsDialog')
      .addItem('✅ Validate Settings', 'validateSettingsDialog')
      .addSeparator()
      .addItem('📊 System Status', 'showQuickStatus')
      .addItem('🔍 Quick API Check', 'quickAPICheckAll')
      .addItem('🐛 Debug Tools...', 'debugWizard')
      .addSeparator()
      .addItem('🐙 GitHub Repository', 'openGitHubRepo')
      .addToUi();
}

function updateSelectedProjectsToCurrent() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    ui.alert('🔐 Token Required', 'Bearer token is not configured. Please set it in Settings sheet first.', ui.ButtonSet.OK);
    return;
  }
  
  var projects = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Overall'];
  var selected = showMultiChoice('Select Projects to Update:', projects);
  
  if (!selected || selected.length === 0) {
    ui.alert('No Selection', 'No projects selected for update.', ui.ButtonSet.OK);
    return;
  }
  
  var result = ui.alert('🔄 Update Selected Projects', 
    `Update ${selected.length} selected projects?\n\n${selected.join(', ')}\n\nThis may take several minutes.`, 
    ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  try {
    var successCount = 0;
    var errors = [];
    
    ui.alert('Processing...', `Updating ${selected.length} projects. Please wait...`, ui.ButtonSet.OK);
    
    selected.forEach(function(proj, index) {
      try {
        var projectName = proj.toUpperCase();
        console.log(`Updating ${projectName} (${index + 1}/${selected.length})...`);
        
        if (index > 0) {
          console.log('Waiting before next project...');
          Utilities.sleep(4000);
        }
        
        updateProjectDataWithRetry(projectName);
        successCount++;
        console.log(`${projectName} updated successfully`);
        
      } catch (e) {
        console.error(`Error updating ${proj}:`, e);
        errors.push(`${proj}: ${e.toString().substring(0, 50)}...`);
        Utilities.sleep(2000);
      }
    });
    
    if (successCount > 0) {
      try {
        console.log('Sorting project sheets...');
        Utilities.sleep(2000);
        sortProjectSheetsWithRetry();
      } catch (e) {
        console.error('Error sorting sheets:', e);
        errors.push(`Sorting: ${e.toString().substring(0, 30)}...`);
      }
    }
    
    var message = `✅ Update completed!\n\n• Successfully updated: ${successCount}/${selected.length} projects`;
    if (errors.length > 0) {
      message += `\n• Errors:\n${errors.join('\n')}`;
      message += '\n\n💡 TIP: Try updating problematic projects individually.';
    }
    
    ui.alert('Update Complete', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error during update: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function refreshSettingsDialog() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var settings = refreshSettingsFromSheet();
    
    var message = '🔄 Settings Refreshed!\n\n';
    message += `🔐 Bearer Token: ${settings.bearerToken ? 'Found' : 'Not Set'}\n`;
    message += `💾 Auto Cache: ${settings.automation.autoCache ? 'Enabled' : 'Disabled'}\n`;
    message += `🔄 Auto Update: ${settings.automation.autoUpdate ? 'Enabled' : 'Disabled'}\n`;
    message += `🎯 Target eROAS: ${Object.keys(settings.targetEROAS).length} projects configured\n`;
    
    try {
      syncTriggersWithSettings();
      message += '\n✅ Triggers synchronized';
    } catch (e) {
      message += '\n⚠️ Error syncing triggers: ' + e.toString();
    }
    
    ui.alert('Settings Refreshed', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error refreshing settings: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function syncTriggersWithSettings() {
  try {
    var settings = loadSettingsFromSheet();
    var triggers = ScriptApp.getProjectTriggers();
    
    var cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
    var updateTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateAllProjects'; });
    
    if (settings.automation.autoCache && !cacheTrigger) {
      ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(3).everyDays(1).create();
      console.log('Created auto cache trigger');
    } else if (!settings.automation.autoCache && cacheTrigger) {
      ScriptApp.deleteTrigger(cacheTrigger);
      console.log('Deleted auto cache trigger');
    }
    
    if (settings.automation.autoUpdate && !updateTrigger) {
      ScriptApp.newTrigger('autoUpdateAllProjects').timeBased().atHour(5).everyDays(1).create();
      console.log('Created auto update trigger');
    } else if (!settings.automation.autoUpdate && updateTrigger) {
      ScriptApp.deleteTrigger(updateTrigger);
      console.log('Deleted auto update trigger');
    }
    
    console.log('Triggers synchronized with Settings sheet');
  } catch (e) {
    console.error('Error syncing triggers with settings:', e);
    throw e;
  }
}

function showQuickStatus() {
  var ui = SpreadsheetApp.getUi();
  
  refreshSettingsFromSheet();
  
  var tokenStatus = isBearerTokenConfigured() ? '✅ Configured' : '❌ Not Set';
  var cacheStatus = isAutoCacheEnabled() ? '✅ Enabled' : '❌ Disabled';
  var updateStatus = isAutoUpdateEnabled() ? '✅ Enabled' : '❌ Disabled';
  
  var message = '📊 SYSTEM STATUS\n\n';
  message += `🔐 Bearer Token: ${tokenStatus}\n`;
  message += `💾 Auto Cache: ${cacheStatus}\n`;
  message += `🔄 Auto Update: ${updateStatus}\n\n`;
  
  var triggers = ScriptApp.getProjectTriggers();
  var cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
  var updateTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateAllProjects'; });
  
  var cacheEnabled = isAutoCacheEnabled();
  var updateEnabled = isAutoUpdateEnabled();
  
  var syncIssues = [];
  if (cacheEnabled && !cacheTrigger) {
    syncIssues.push('• Cache trigger missing (will auto-create)');
  }
  if (!cacheEnabled && cacheTrigger) {
    syncIssues.push('• Cache trigger exists but disabled (will remove)');
  }
  if (updateEnabled && !updateTrigger) {
    syncIssues.push('• Update trigger missing (will auto-create)');
  }
  if (!updateEnabled && updateTrigger) {
    syncIssues.push('• Update trigger exists but disabled (will remove)');
  }
  
  if (syncIssues.length > 0) {
    message += '⚠️ SYNC ISSUES:\n' + syncIssues.join('\n') + '\n\n';
    message += 'Use "🔄 Refresh Settings" to fix.\n\n';
  } else {
    message += '✅ All triggers synchronized\n\n';
  }
  
  message += '💡 TIP: Edit settings directly in Settings sheet';
  
  ui.alert('System Status', message, ui.ButtonSet.OK);
}

function refreshMenu() {
  var ui = SpreadsheetApp.getUi();
  try {
    onOpen();
    ui.alert('Menu Refreshed', 'Menu has been refreshed with current settings.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error refreshing menu: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function validateSettingsDialog() {
  var ui = SpreadsheetApp.getUi();
  var validation = validateSettings();
  
  if (validation.valid) {
    ui.alert('✅ Settings Valid', 'All settings are configured correctly!', ui.ButtonSet.OK);
  } else {
    var message = '❌ Settings Issues Found:\n\n';
    validation.issues.forEach(function(issue) {
      message += '• ' + issue + '\n';
    });
    message += '\nOpen Settings sheet to fix these issues.';
    ui.alert('Settings Validation', message, ui.ButtonSet.OK);
  }
}

function quickAPICheckAll() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    if (ui.alert('🔐 Token Required', 'Bearer token not configured. Open Settings sheet?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
      openSettingsSheet();
    }
    return;
  }
  
  var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
  var results = '🔍 API CHECK RESULTS\n\n';
  
  projects.forEach(function(proj) {
    try {
      setCurrentProject(proj);
      var dateRange = getDateRange(7);
      var raw = fetchCampaignData(dateRange);
      
      if (!raw.data?.analytics?.richStats?.stats?.length) {
        results += `❌ ${proj}: No data\n`;
      } else {
        var count = raw.data.analytics.richStats.stats.length;
        results += `✅ ${proj}: ${count} records\n`;
      }
    } catch (e) {
      results += `❌ ${proj}: ${e.toString().substring(0, 30)}...\n`;
    }
  });
  
  ui.alert('API Check Complete', results, ui.ButtonSet.OK);
}

function updateAllProjectsToCurrent() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    ui.alert('🔐 Token Required', 'Bearer token is not configured. Please set it in Settings sheet first.', ui.ButtonSet.OK);
    return;
  }
  
  var result = ui.alert('🔄 Update All Projects', 
    'This will update all projects with the latest data (up to last complete week).\n\nThis may take several minutes. Continue?', 
    ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  try {
    var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    var successCount = 0;
    var errors = [];
    
    ui.alert('Processing...', 'Starting batch update. Please wait...', ui.ButtonSet.OK);
    
    projects.forEach(function(proj, index) {
      try {
        console.log(`Updating ${proj} (${index + 1}/${projects.length})...`);
        
        if (index > 0) {
          console.log('Waiting before next project...');
          Utilities.sleep(5000);
        }
        
        updateProjectDataWithRetry(proj);
        successCount++;
        console.log(`${proj} updated successfully`);
        
      } catch (e) {
        console.error(`Error updating ${proj}:`, e);
        errors.push(`${proj}: ${e.toString().substring(0, 50)}...`);
        Utilities.sleep(3000);
      }
    });
    
    if (successCount > 0) {
      try {
        console.log('Sorting project sheets...');
        Utilities.sleep(2000);
        sortProjectSheetsWithRetry();
        console.log('Sheets sorted successfully');
      } catch (e) {
        console.error('Error sorting sheets:', e);
        errors.push(`Sorting: ${e.toString().substring(0, 50)}...`);
      }
    }
    
    var message = `✅ Update completed!\n\n• Successfully updated: ${successCount}/${projects.length} projects`;
    if (errors.length > 0) {
      message += `\n• Errors:\n${errors.join('\n')}`;
      message += '\n\n💡 TIP: Try updating projects individually if errors persist.';
    }
    
    ui.alert('Update Complete', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error during update: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function updateProjectDataWithRetry(projectName, maxRetries = 2) {
  var baseDelay = 3000;
  
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      updateProjectData(projectName);
      return;
    } catch (e) {
      console.error(`${projectName} update attempt ${attempt} failed:`, e);
      
      if (attempt === maxRetries) {
        throw e;
      }
      
      var delay = baseDelay * Math.pow(2, attempt - 1);
      console.log(`Waiting ${delay}ms before retry...`);
      Utilities.sleep(delay);
    }
  }
}

function sortProjectSheetsWithRetry(maxRetries = 2) {
  var baseDelay = 2000;
  
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      sortProjectSheets();
      return;
    } catch (e) {
      console.error(`Sheet sorting attempt ${attempt} failed:`, e);
      
      if (attempt === maxRetries) {
        throw e;
      }
      
      var delay = baseDelay * attempt;
      console.log(`Waiting ${delay}ms before retry...`);
      Utilities.sleep(delay);
    }
  }
}

function smartReportWizard() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    if (ui.alert('🔐 Token Required', 'Bearer token is not configured. Open Settings sheet?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
      openSettingsSheet();
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

function debugWizard() {
  var choice = showChoice('🐛 Debug Tools', [
    'Debug Single Project',
    'API Health Check All',
    'Clear All Data',
    'Apps Database (TRICKY)',
    'View Settings Status',
    '📊 Growth Thresholds Examples'
  ]);
  if (!choice) return;
  
  switch(choice) {
    case 1:
      var p = showChoice('Select Project to Debug:', MENU_PROJECTS);
      if (p) debugProjectReportGeneration(MENU_PROJECTS[p-1].toUpperCase());
      break;
    case 2:
      quickAPICheckAll();
      break;
    case 3:
      clearDataWizard();
      break;
    case 4:
      appsDbWizard();
      break;
    case 5:
      showSettingsStatus();
      break;
    case 6:
      growthThresholdsExamplesWizard();
      break;
  }
}

function growthThresholdsExamplesWizard() {
  var ui = SpreadsheetApp.getUi();
  var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
  
  var choice = showChoice('📊 Growth Thresholds Examples', [
    'Apply Example to Single Project',
    'Apply Example to All Projects',
    'View Current Thresholds Summary',
    'Reset All to Defaults'
  ]);
  
  if (!choice) return;
  
  switch(choice) {
    case 1:
      var project = showChoice('Select Project:', projects);
      if (project) {
        createExampleGrowthThresholds(projects[project-1]);
      }
      break;
      
    case 2:
      var exampleChoice = ui.alert('Apply Example to All Projects', 
        'Choose example type:\n\nYES = Conservative (осторожные)\nNO = Standard (стандартные)\nCANCEL = Aggressive (агрессивные)', 
        ui.ButtonSet.YES_NO_CANCEL);
      
      if (exampleChoice !== ui.Button.CANCEL && exampleChoice !== ui.Button.CLOSE) {
        var confirmed = ui.alert('Confirm', 'Apply selected example to ALL projects?', ui.ButtonSet.YES_NO);
        if (confirmed === ui.Button.YES) {
          applyExampleToAllProjects(exampleChoice);
        }
      }
      break;
      
    case 3:
      showCurrentThresholdsSummary();
      break;
      
    case 4:
      if (ui.alert('Reset All Thresholds', 'Reset ALL projects to default thresholds?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
        resetAllThresholdsToDefaults();
      }
      break;
  }
}

function applyExampleToAllProjects(exampleChoice) {
  var ui = SpreadsheetApp.getUi();
  var sheet = getOrCreateSettingsSheet();
  var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
  
  var examples = {
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
  
  var selectedExample;
  if (exampleChoice === ui.Button.YES) selectedExample = examples.conservative;
  else if (exampleChoice === ui.Button.NO) selectedExample = examples.standard;
  else return;
  
  var data = sheet.getDataRange().getValues();
  var updatedCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var label = data[i][0] ? data[i][0].toString().trim() : '';
    
    projects.forEach(function(proj) {
      if (label === `${proj}:` && i >= 21 && i <= 30) {
        sheet.getRange(i + 1, 2).setValue(selectedExample.healthy);
        sheet.getRange(i + 1, 3).setValue(selectedExample.efficiency);
        sheet.getRange(i + 1, 4).setValue(selectedExample.inefficient);
        sheet.getRange(i + 1, 5).setValue(selectedExample.scaling);
        sheet.getRange(i + 1, 6).setValue(selectedExample.other);
        updatedCount++;
      }
    });
  }
  
  clearSettingsCache();
  ui.alert('✅ Applied to All', `Updated growth thresholds for ${updatedCount} projects.`, ui.ButtonSet.OK);
}

function showCurrentThresholdsSummary() {
  var ui = SpreadsheetApp.getUi();
  var settings = loadSettingsFromSheet();
  var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
  
  var message = '📊 CURRENT GROWTH THRESHOLDS SUMMARY\n\n';
  
  projects.forEach(function(proj) {
    var thresholds = settings.growthThresholds[proj];
    if (thresholds) {
      message += `${proj}:\n`;
      message += `🟢 Healthy: ${thresholds.healthyGrowth.minSpendChange}%/${thresholds.healthyGrowth.minProfitChange}%\n`;
      message += `🔴 Inefficient: ${thresholds.inefficientGrowth.maxProfitChange}%\n`;
      message += `🔵 Scaling: ${thresholds.scalingDown.maxSpendChange}%\n`;
      message += `🟡 Moderate: ${thresholds.moderateGrowthSpend}%/${thresholds.moderateGrowthProfit}%\n\n`;
    } else {
      message += `${proj}: Not configured\n\n`;
    }
  });
  
  ui.alert('Growth Thresholds Summary', message, ui.ButtonSet.OK);
}

function resetAllThresholdsToDefaults() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getOrCreateSettingsSheet();
  var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
  
  var defaultExample = {
    healthy: 'spend:10,profit:5',
    efficiency: 'spendDrop:-5,profitGain:8',
    inefficient: 'profitDrop:-8',
    scaling: 'spendDrop:-15,efficientProfit:0,moderateMin:-1,moderateMax:-10',
    other: 'modSpend:3,modProfit:2,stable:2'
  };
  
  var data = sheet.getDataRange().getValues();
  var resetCount = 0;
  
  for (var i = 0; i < data.length; i++) {
    var label = data[i][0] ? data[i][0].toString().trim() : '';
    
    projects.forEach(function(proj) {
      if (label === `${proj}:` && i >= 21 && i <= 30) {
        sheet.getRange(i + 1, 2).setValue(defaultExample.healthy);
        sheet.getRange(i + 1, 3).setValue(defaultExample.efficiency);
        sheet.getRange(i + 1, 4).setValue(defaultExample.inefficient);
        sheet.getRange(i + 1, 5).setValue(defaultExample.scaling);
        sheet.getRange(i + 1, 6).setValue(defaultExample.other);
        resetCount++;
      }
    });
  }
  
  clearSettingsCache();
  ui.alert('✅ Reset Complete', `Reset growth thresholds for ${resetCount} projects to defaults.`, ui.ButtonSet.OK);
}

function showSettingsStatus() {
  var ui = SpreadsheetApp.getUi();
  try {
    var settings = loadSettingsFromSheet();
    var message = '⚙️ SETTINGS STATUS\n\n';
    
    message += '🔐 API:\n';
    message += `• Bearer Token: ${settings.bearerToken ? 'Configured (' + settings.bearerToken.length + ' chars)' : 'Not Set'}\n\n`;
    
    message += '🎯 Target eROAS:\n';
    Object.keys(settings.targetEROAS).forEach(function(proj) {
      message += `• ${proj}: ${settings.targetEROAS[proj]}%\n`;
    });
    
    message += '\n🤖 Automation:\n';
    message += `• Auto Cache: ${settings.automation.autoCache ? 'Enabled (daily 3 AM CET)' : 'Disabled'}\n`;
    message += `• Auto Update: ${settings.automation.autoUpdate ? 'Enabled (daily 5 AM CET)' : 'Disabled'}\n`;
    
    message += '\n📊 Growth Thresholds: Configured for all projects';
    
    ui.alert('Settings Status', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error reading settings: ' + e.toString(), ui.ButtonSet.OK);
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

function clearAllProjectsData() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Confirm Clear All', 'Clear data from ALL projects? Comments preserved.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    var successCount = 0;
    
    projects.forEach(function(proj) {
      try {
        clearProjectDataSilent(proj);
        successCount++;
      } catch (e) {
        console.error(`Error clearing ${proj}:`, e);
      }
    });
    
    ui.alert(successCount === projects.length ? 'Success' : 'Partial Success', 
             `Cleared ${successCount} of ${projects.length} projects. Comments preserved.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error clearing data: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function clearProjectAllData(projectName) {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert(`Clear ${projectName} Data`, `Clear all ${projectName} data? Comments preserved.`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    clearProjectDataSilent(projectName);
    ui.alert('Success', `${projectName} data cleared. Comments preserved.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Error clearing ${projectName}: ${e.toString()}`, ui.ButtonSet.OK);
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
    } else {
      message += '• Status: Empty cache\n';
      message += '• Action Required: Refresh database';
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
      ui.alert('No Data', 'Apps Database cache is empty. Please refresh first.', ui.ButtonSet.OK);
      return;
    }
    
    var message = '📱 APPS DATABASE SAMPLE\n\n';
    var sampleCount = Math.min(5, bundleIds.length);
    
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
  
  if (ui.alert('Clear Apps Database Cache', 'Clear cached app data? Will rebuild on next refresh.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    var appsDb = new AppsDatabase('TRICKY');
    if (appsDb.cacheSheet && appsDb.cacheSheet.getLastRow() > 1) {
      appsDb.cacheSheet.deleteRows(2, appsDb.cacheSheet.getLastRow() - 1);
      ui.alert('Success', 'Apps Database cache cleared.', ui.ButtonSet.OK);
    } else {
      ui.alert('No Cache', 'Apps Database cache sheet not found.', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', 'Error clearing cache: ' + e.toString(), ui.ButtonSet.OK);
  }
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
  var start = ui.prompt('Start Date', 'Enter date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (start.getSelectedButton() !== ui.Button.OK) return null;
  var end = ui.prompt('End Date', 'Enter date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (end.getSelectedButton() !== ui.Button.OK) return null;
  if (!isValidDate(start.getResponseText()) || !isValidDate(end.getResponseText())) {
    ui.alert('❌ Invalid date format');
    return null;
  }
  return { start: start.getResponseText(), end: end.getResponseText() };
}

function isValidDate(dateString) { 
  var regex = /^\d{4}-\d{2}-\d{2}$/; 
  if (!regex.test(dateString)) return false; 
  var date = new Date(dateString); 
  return date instanceof Date && !isNaN(date); 
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
    sortProjectSheets();
    ui.alert('✅ Complete', 'Generated ' + success + '/' + MENU_PROJECTS.length + ' reports', ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Error', e.toString(), ui.ButtonSet.OK);
  }
}

function runSelectedProjects(projects, days) {
  for (var i = 0; i < projects.length; i++) {
    generateProjectReport(projects[i].toUpperCase(), days);
  }
  sortProjectSheets();
  SpreadsheetApp.getUi().alert('✅ Complete', 'Generated ' + projects.length + ' reports', SpreadsheetApp.getUi().ButtonSet.OK);
}

function runAllProjectsDateRange(start, end) {
  for (var i = 0; i < MENU_PROJECTS.length; i++) {
    generateProjectReportForDateRange(MENU_PROJECTS[i].toUpperCase(), start, end);
  }
  sortProjectSheets();
  SpreadsheetApp.getUi().alert('✅ Complete', 'All reports generated', SpreadsheetApp.getUi().ButtonSet.OK);
}

function runSelectedProjectsDateRange(projects, start, end) {
  for (var i = 0; i < projects.length; i++) {
    generateProjectReportForDateRange(projects[i].toUpperCase(), start, end);
  }
  sortProjectSheets();
  SpreadsheetApp.getUi().alert('✅ Complete', 'Generated ' + projects.length + ' reports', SpreadsheetApp.getUi().ButtonSet.OK);
}

function generateProjectReport(projectName, days) { setCurrentProject(projectName); generateReport(days); }
function generateProjectReportForDateRange(projectName, startDate, endDate) { setCurrentProject(projectName); generateReportForDateRange(startDate, endDate); }
function debugProjectReportGeneration(projectName) { setCurrentProject(projectName); debugReportGeneration(); }

function openGitHubRepo() {
  var ui = SpreadsheetApp.getUi();
  var githubUrl = 'https://github.com/Poxagronka/pivot-table';
  
  var htmlOutput = HtmlService.createHtmlOutput(
    '<script>window.open("' + githubUrl + '", "_blank"); google.script.host.close();</script>'
  ).setWidth(400).setHeight(300);
  
  ui.showModalDialog(htmlOutput, 'Opening GitHub Repository...');
}