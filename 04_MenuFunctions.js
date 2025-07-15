var MENU_PROJECTS = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Overall'];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📊 Campaign Report');
  
  menu.addItem('📈 Generate Report...', 'smartReportWizard')
      .addItem('🔄 Update All to Current', 'updateAllProjectsInBatches')
      .addItem('🎯 Update Selected Projects', 'updateSelectedProjectsToCurrent')
      .addSeparator();
  
  var settingsMenu = ui.createMenu('⚙️ Settings');
  settingsMenu.addItem('📝 Open Settings Sheet', 'openSettingsSheet')
            .addItem('🔄 Refresh Settings', 'refreshSettingsDialog')
            .addItem('🔧 Force Update Settings', 'forceUpdateSettingsSheet')
            .addItem('📊 System Status', 'showQuickStatus')
            .addSeparator()
            .addItem('💾 Save All Comments', 'saveAllCommentsToCache')
            .addItem('🔍 Quick API Check', 'quickAPICheckAll')
            .addItem('🗑️ Clear Data...', 'clearDataWizard')
            .addItem('📱 Apps Database (TRICKY)', 'appsDbWizard')
            .addSeparator()
            .addItem('🐙 GitHub Repository', 'openGitHubRepo')
            .addItem('🔄 Recreate Triggers', 'recreateAllTriggers');
  
  menu.addSubMenu(settingsMenu).addToUi();
}

function updateAllProjectsInBatches() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    ui.alert('🔐 Token Required', 'Bearer token is not configured. Please set it in Settings sheet first.', ui.ButtonSet.OK);
    return;
  }
  
  var result = ui.alert('🔄 Update All Projects', 'Update all projects in optimized batches?\n\nThis will be slower but more reliable.', ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  try {
    var batch1 = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS'];
    var batch2 = ['APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    
    var batch1Results = updateProjectBatchOptimized(batch1, 1);
    
    if (batch1Results.successCount > 0) Utilities.sleep(30000);
    
    var batch2Results = updateProjectBatchOptimized(batch2, 2);
    
    var totalSuccess = batch1Results.successCount + batch2Results.successCount;
    var totalErrors = batch1Results.errors.concat(batch2Results.errors);
    
    try {
      Utilities.sleep(5000);
      sortProjectSheetsWithRetry();
    } catch (e) {
      totalErrors.push(`Sorting: ${e.toString().substring(0, 50)}...`);
    }
    
    var message = `✅ Batch update completed!\n\n• Successfully updated: ${totalSuccess}/8 projects`;
    if (totalErrors.length > 0) {
      message += `\n• Errors:\n${totalErrors.join('\n')}\n\n💡 TIP: Try updating failed projects individually.`;
    }
    
    ui.alert('Update Complete', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error during batch update: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function updateProjectBatchOptimized(projects, batchNumber) {
  var successCount = 0;
  var errors = [];
  
  projects.forEach(function(proj, index) {
    try {
      if (proj === 'TRICKY') clearTrickyCaches();
      
      if (index > 0) {
        const waitTime = proj === 'TRICKY' ? 12000 : 8000;
        Utilities.sleep(waitTime);
      }
      
      updateProjectDataOptimized(proj);
      successCount++;
      
    } catch (e) {
      var errorMsg = e.toString();
      if (errorMsg.includes('timeout') || errorMsg.includes('timed out')) {
        errors.push(`${proj}: Timeout - try individually`);
      } else {
        errors.push(`${proj}: ${errorMsg.substring(0, 50)}...`);
      }
      
      Utilities.sleep(5000);
    }
  });
  
  return { successCount: successCount, errors: errors };
}

function updateSelectedProjectsToCurrent() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    ui.alert('🔐 Token Required', 'Bearer token is not configured. Please set it in Settings sheet first.', ui.ButtonSet.OK);
    return;
  }
  
  var selected = showMultiChoice('Select Projects to Update:', MENU_PROJECTS);
  
  if (!selected || selected.length === 0) {
    ui.alert('No Selection', 'No projects selected for update.', ui.ButtonSet.OK);
    return;
  }
  
  var result = ui.alert('🔄 Update Selected Projects', `Update ${selected.length} selected projects?\n\n${selected.join(', ')}\n\nThis may take several minutes.`, ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  try {
    var successCount = 0;
    var errors = [];
    
    selected.forEach(function(proj, index) {
      try {
        var projectName = proj.toUpperCase();
        
        if (projectName === 'TRICKY') clearTrickyCaches();
        
        if (index > 0) {
          const waitTime = projectName === 'TRICKY' ? 12000 : 8000;
          Utilities.sleep(waitTime);
        }
        
        updateProjectDataOptimized(projectName);
        successCount++;
        
      } catch (e) {
        var errorMsg = e.toString();
        if (errorMsg.includes('timeout') || errorMsg.includes('timed out')) {
          errors.push(`${proj}: Timeout - try individually`);
        } else {
          errors.push(`${proj}: ${errorMsg.substring(0, 50)}...`);
        }
        
        Utilities.sleep(5000);
      }
    });
    
    if (successCount > 0) {
      try {
        Utilities.sleep(3000);
        sortProjectSheetsWithRetry();
      } catch (e) {
        errors.push(`Sorting: ${e.toString().substring(0, 30)}...`);
      }
    }
    
    var message = `✅ Update completed!\n\n• Successfully updated: ${successCount}/${selected.length} projects`;
    if (errors.length > 0) {
      message += `\n• Errors:\n${errors.join('\n')}\n\n💡 TIP: Try updating problematic projects individually.`;
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
    message += `🎯 eROAS D730 Targets: Updated\n`;
    
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

function showQuickStatus() {
  var ui = SpreadsheetApp.getUi();
  
  refreshSettingsFromSheet();
  
  var tokenStatus = isBearerTokenConfigured() ? '✅ Configured' : '❌ Not Set';
  var cacheStatus = isAutoCacheEnabled() ? '✅ Enabled' : '❌ Disabled';
  var updateStatus = isAutoUpdateEnabled() ? '✅ Enabled' : '❌ Disabled';
  
  var message = '📊 SYSTEM STATUS\n\n';
  message += `🔐 Bearer Token: ${tokenStatus}\n`;
  message += `💾 Auto Cache: ${cacheStatus}\n`;
  message += `🔄 Auto Update: ${updateStatus}\n`;
  message += `🎯 Metrics: Unified (eROAS D730)\n\n`;
  
  var triggers = ScriptApp.getProjectTriggers();
  var cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
  var updateTriggers = getUpdateTriggers();
  
  var cacheEnabled = isAutoCacheEnabled();
  var updateEnabled = isAutoUpdateEnabled();
  
  var syncIssues = [];
  if (cacheEnabled && !cacheTrigger) syncIssues.push('• Cache trigger missing (will auto-create)');
  if (!cacheEnabled && cacheTrigger) syncIssues.push('• Cache trigger exists but disabled (will remove)');
  if (updateEnabled && updateTriggers.length !== 8) syncIssues.push('• Update triggers incomplete (will auto-create)');
  if (!updateEnabled && updateTriggers.length > 0) syncIssues.push('• Update triggers exist but disabled (will remove)');
  
  if (syncIssues.length > 0) {
    message += '⚠️ SYNC ISSUES:\n' + syncIssues.join('\n') + '\n\nUse "🔄 Refresh Settings" to fix.\n\n';
  } else {
    message += '✅ All triggers synchronized\n\n';
  }
  
  message += '📅 AUTOMATION SCHEDULE:\n';
  message += '• Auto Cache: Daily at 2:00 AM\n';
  message += '• Auto Update: 5:00-6:00 AM (exact times)\n';
  
  ui.alert('System Status', message, ui.ButtonSet.OK);
}

function getUpdateTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var updateFunctions = ['autoUpdateTricky','autoUpdateMoloco','autoUpdateRegular','autoUpdateGoogleAds','autoUpdateApplovin','autoUpdateMintegral','autoUpdateIncent','autoUpdateOverall'];
  return triggers.filter(function(t) { return updateFunctions.includes(t.getHandlerFunction()); });
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
      
      if (proj === 'TRICKY') clearTrickyCaches();
      
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
  
  var period = showChoice('📅 Select Period - Step 2/3', ['Number of weeks', 'Date range (specific dates)']);
  if (!period) return;
  
  if (scope === 1) {
    if (period === 1) {
      var weeks = promptWeeks();
      if (weeks) quickGenerateAllForWeeks(weeks);
    } else {
      var dates = promptDateRange();
      if (dates) runAllProjectsDateRangeOptimized(dates.start, dates.end);
    }
  } else if (scope === 2) {
    var project = showChoice('Select Project - Step 3/3', MENU_PROJECTS);
    if (!project) return;
    var projectName = MENU_PROJECTS[project-1].toUpperCase();
    
    if (period === 1) {
      var weeks = promptWeeks();
      if (weeks) generateProjectReportOptimized(projectName, weeks * 7);
    } else {
      var dates = promptDateRange();
      if (dates) generateProjectReportForDateRangeOptimized(projectName, dates.start, dates.end);
    }
  } else {
    var selected = showMultiChoice('Select Projects:', MENU_PROJECTS);
    if (!selected || selected.length === 0) return;
    
    if (period === 1) {
      var weeks = promptWeeks();
      if (weeks) runSelectedProjectsOptimized(selected, weeks * 7);
    } else {
      var dates = promptDateRange();
      if (dates) runSelectedProjectsDateRangeOptimized(selected, dates.start, dates.end);
    }
  }
}

function clearDataWizard() {
  var choice = showChoice('🗑️ Clear Data', ['Clear All Projects', 'Clear Single Project']);
  if (!choice) return;
  
  if (choice === 1) {
    clearAllProjectsDataOptimized();
  } else {
    var p = showChoice('Select Project:', MENU_PROJECTS);
    if (p) clearProjectAllDataOptimized(MENU_PROJECTS[p-1].toUpperCase());
  }
}

function clearAllProjectsDataOptimized() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Confirm Clear All', 'Clear data from ALL projects? Comments preserved.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    var successCount = 0;
    
    projects.forEach(function(proj) {
      try {
        if (proj === 'TRICKY') clearTrickyCaches();
        clearProjectDataSilent(proj);
        successCount++;
      } catch (e) {}
    });
    
    ui.alert(successCount === projects.length ? 'Success' : 'Partial Success', `Cleared ${successCount} of ${projects.length} projects. Comments preserved.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error clearing data: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function clearProjectAllDataOptimized(projectName) {
  var ui = SpreadsheetApp.getUi();
  
  if (ui.alert(`Clear ${projectName} Data`, `Clear all ${projectName} data? Comments preserved.`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    if (projectName === 'TRICKY') clearTrickyCaches();
    clearProjectDataSilent(projectName);
    ui.alert('Success', `${projectName} data cleared. Comments preserved.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Error clearing ${projectName}: ${e.toString()}`, ui.ButtonSet.OK);
  }
}

function syncTriggersWithSettings() {
  try {
    var settings = loadSettingsFromSheet();
    var triggers = ScriptApp.getProjectTriggers();
    
    var cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
    var updateTriggers = getUpdateTriggers();
    
    if (settings.automation.autoCache && !cacheTrigger) {
      ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
    } else if (!settings.automation.autoCache && cacheTrigger) {
      ScriptApp.deleteTrigger(cacheTrigger);
    }
    
    if (settings.automation.autoUpdate && updateTriggers.length !== 8) {
      clearAllUpdateTriggers();
      createUpdateTriggers();
    } else if (!settings.automation.autoUpdate && updateTriggers.length > 0) {
      clearAllUpdateTriggers();
    }
  } catch (e) {
    throw e;
  }
}

function clearAllUpdateTriggers() {
  var updateFunctions = ['autoUpdateTricky','autoUpdateMoloco','autoUpdateRegular','autoUpdateGoogleAds','autoUpdateApplovin','autoUpdateMintegral','autoUpdateIncent','autoUpdateOverall','autoUpdateAllProjects'];
  
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return updateFunctions.includes(t.getHandlerFunction()); })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
}

function createUpdateTriggers() {
  var schedule = [
    { func: 'autoUpdateTricky', hour: 5, minute: 0 },
    { func: 'autoUpdateMoloco', hour: 5, minute: 10 },
    { func: 'autoUpdateRegular', hour: 5, minute: 20 },
    { func: 'autoUpdateGoogleAds', hour: 5, minute: 30 },
    { func: 'autoUpdateApplovin', hour: 5, minute: 40 },
    { func: 'autoUpdateMintegral', hour: 5, minute: 50 },
    { func: 'autoUpdateIncent', hour: 6, minute: 0 },
    { func: 'autoUpdateOverall', hour: 6, minute: 10 }
  ];
  
  schedule.forEach(function(item) {
    ScriptApp.newTrigger(item.func).timeBased().everyDays(1).atHour(item.hour).nearMinute(item.minute).create();
  });
}

function sortProjectSheetsWithRetry(maxRetries = 2) {
  var baseDelay = 2000;
  
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      sortProjectSheets();
      return;
    } catch (e) {
      if (attempt === maxRetries) throw e;
      Utilities.sleep(baseDelay * attempt);
    }
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

function promptWeeks() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Number of Weeks', 'Enter number of weeks for report (e.g., 4 for last 4 weeks):', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  
  var input = result.getResponseText().trim();
  var weeks = parseInt(input);
  
  if (isNaN(weeks) || weeks < 1) {
    ui.alert('❌ Invalid Input', 'Please enter a valid number of weeks (minimum 1).', ui.ButtonSet.OK);
    return null;
  }
  
  if (weeks > 52) {
    ui.alert('❌ Period Too Large', `${weeks} weeks (over 1 year) is too large.\n\nMaximum recommended: 52 weeks\n\nFor large historical data, use "Date range" option.`, ui.ButtonSet.OK);
    return null;
  }
  
  if (weeks > 26) {
    var confirm = ui.alert('⚠️ Large Period Warning', `You entered ${weeks} weeks (${Math.round(weeks/4)} months).\n\nThis may take longer to process.\n\nContinue?`, ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return null;
  }
  
  return weeks;
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

function quickGenerateAllForWeeks(weeks) {
  var ui = SpreadsheetApp.getUi();
  var success = 0;
  
  try {
    for (var i = 0; i < MENU_PROJECTS.length; i++) {
      var p = MENU_PROJECTS[i];
      try { 
        generateProjectReportOptimized(p.toUpperCase(), weeks * 7); 
        success++; 
      } catch(e) {}
    }
    sortProjectSheets();
    ui.alert('✅ Complete', 'Generated ' + success + '/' + MENU_PROJECTS.length + ' reports', ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Error', e.toString(), ui.ButtonSet.OK);
  }
}

function runSelectedProjectsOptimized(projects, days) {
  for (var i = 0; i < projects.length; i++) {
    generateProjectReportOptimized(projects[i].toUpperCase(), days);
  }
  sortProjectSheets();
  SpreadsheetApp.getUi().alert('✅ Complete', 'Generated ' + projects.length + ' reports', SpreadsheetApp.getUi().ButtonSet.OK);
}

function runAllProjectsDateRangeOptimized(start, end) {
  for (var i = 0; i < MENU_PROJECTS.length; i++) {
    generateProjectReportForDateRangeOptimized(MENU_PROJECTS[i].toUpperCase(), start, end);
  }
  sortProjectSheets();
  SpreadsheetApp.getUi().alert('✅ Complete', 'All reports generated', SpreadsheetApp.getUi().ButtonSet.OK);
}

function runSelectedProjectsDateRangeOptimized(projects, start, end) {
  for (var i = 0; i < projects.length; i++) {
    generateProjectReportForDateRangeOptimized(projects[i].toUpperCase(), start, end);
  }
  sortProjectSheets();
  SpreadsheetApp.getUi().alert('✅ Complete', 'Generated ' + projects.length + ' reports', SpreadsheetApp.getUi().ButtonSet.OK);
}

function generateProjectReportOptimized(projectName, days) { 
  if (projectName === 'TRICKY') clearTrickyCaches();
  setCurrentProject(projectName); 
  generateReport(days); 
}

function generateProjectReportForDateRangeOptimized(projectName, startDate, endDate) { 
  if (projectName === 'TRICKY') clearTrickyCaches();
  setCurrentProject(projectName); 
  generateReportForDateRange(startDate, endDate); 
}

function appsDbWizard() {
  var ui = SpreadsheetApp.getUi();
  
  if (CURRENT_PROJECT !== 'TRICKY') {
    var switchResult = ui.alert('Apps Database - TRICKY Only', 'Apps Database is only used for TRICKY project.\n\nSwitch to TRICKY project now?', ui.ButtonSet.YES_NO);
    
    if (switchResult !== ui.Button.YES) return;
    setCurrentProject('TRICKY');
  }
  
  var action = showChoice('📱 Apps Database Management', ['View Cache Status','Refresh Apps Database','View Sample Data','Clear Cache','Clear Optimization Caches']);
  if (!action) return;
  
  switch(action) {
    case 1: showAppsDbStatusOptimized(); break;
    case 2: refreshAppsDatabase(); break;
    case 3: showAppsDbSample(); break;
    case 4: clearAppsDbCache(); break;
    case 5: clearTrickyOptimizationCaches(); break;
  }
}

function showAppsDbStatusOptimized() {
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

function clearTrickyOptimizationCaches() {
  var ui = SpreadsheetApp.getUi();
  
  if (ui.alert('Clear Optimization Caches', 'Clear all TRICKY optimization caches?\n\nThis will force reloading on next use.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    clearTrickyCaches();
    ui.alert('Success', 'TRICKY optimization caches cleared.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error clearing caches: ' + e.toString(), ui.ButtonSet.OK);
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
      clearTrickyCaches();
      ui.alert('Success', 'Apps Database cache cleared.', ui.ButtonSet.OK);
    } else {
      ui.alert('No Cache', 'Apps Database cache sheet not found.', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', 'Error clearing cache: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function openGitHubRepo() {
  var ui = SpreadsheetApp.getUi();
  var githubUrl = 'https://github.com/Poxagronka/pivot-table';
  
  var htmlOutput = HtmlService.createHtmlOutput('<script>window.open("' + githubUrl + '", "_blank"); google.script.host.close();</script>').setWidth(400).setHeight(300);
  
  ui.showModalDialog(htmlOutput, 'Opening GitHub Repository...');
}

function recreateAllTriggers() {
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.alert('🔄 Recreate Triggers', 'Recreate all automation triggers?\n\n⏰ Schedule:\n• Cache: 2:00 AM\n• Updates: 5:00-6:10 AM', ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  try {
    clearAllUpdateTriggers();
    
    var cacheEnabled = isAutoCacheEnabled();
    var updateEnabled = isAutoUpdateEnabled();
    
    if (cacheEnabled) {
      ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
    }
    
    if (updateEnabled) {
      createUpdateTriggers();
    }
    
    ui.alert('✅ Triggers Recreated', 'All triggers recreated successfully!', ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert('❌ Error', 'Error recreating triggers: ' + e.toString(), ui.ButtonSet.OK);
  }
}