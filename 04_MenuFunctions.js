var MENU_PROJECTS = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Overall'];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📊 Campaign Report');
  
  menu.addItem('📈 Generate Report', 'showReportWizard')
      .addItem('🔄 Update All Projects', 'updateAllProjectsQuick')
      .addItem('🔄 Update Selected Projects', 'updateSelectedProjectsQuick')
      .addSeparator()
      .addSubMenu(ui.createMenu('⚙️ Advanced')
        .addItem('Settings Sheet', 'openSettingsSheet')
        .addItem('System Status', 'showQuickStatus')
        .addItem('API Check', 'quickAPICheckAll')
        .addItem('Clear Data', 'clearDataWizard')
        .addItem('Apps Database', 'appsDbWizard')
        .addItem('Debug', 'debugSingleProject'))
      .addItem('📚 Help', 'showHelp')
      .addItem('🐙 GitHub', 'openGitHubRepo')
      .addToUi();
}

function showReportWizard() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    if (ui.alert('🔐 Token Required', 'Bearer token is not configured. Open Settings sheet?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
      openSettingsSheet();
    }
    return;
  }
  
  var projectsPrompt = '📊 SELECT PROJECTS\n\n' +
    '1. All Projects\n' +
    '2. TRICKY only\n' +
    '3. MOLOCO only\n' +
    '4. REGULAR only\n' +
    '5. GOOGLE_ADS only\n' +
    '6. APPLOVIN only\n' +
    '7. MINTEGRAL only\n' +
    '8. INCENT only\n' +
    '9. OVERALL only\n' +
    '10. Custom selection\n\n' +
    'Enter number(s), comma-separated for multiple:';
  
  var projectResponse = ui.prompt('Step 1/2 - Select Projects', projectsPrompt, ui.ButtonSet.OK_CANCEL);
  if (projectResponse.getSelectedButton() !== ui.Button.OK) return;
  
  var selectedProjects = [];
  var choices = projectResponse.getResponseText().split(',');
  
  choices.forEach(function(choice) {
    var num = parseInt(choice.trim());
    if (num === 1) {
      selectedProjects = MENU_PROJECTS.map(function(p) { return p.toUpperCase(); });
    } else if (num >= 2 && num <= 9) {
      selectedProjects.push(MENU_PROJECTS[num - 2].toUpperCase());
    } else if (num === 10) {
      var customProjects = showMultiChoice('Select Projects:', MENU_PROJECTS);
      if (customProjects) {
        selectedProjects = selectedProjects.concat(customProjects.map(function(p) { return p.toUpperCase(); }));
      }
    }
  });
  
  selectedProjects = [...new Set(selectedProjects)];
  
  if (selectedProjects.length === 0) {
    ui.alert('No Selection', 'No projects selected.', ui.ButtonSet.OK);
    return;
  }
  
  var periodPrompt = '📅 SELECT PERIOD\n\n' +
    '1. Last 1 week\n' +
    '2. Last 2 weeks\n' +
    '3. Last 4 weeks\n' +
    '4. Last 8 weeks\n' +
    '5. Last 12 weeks\n' +
    '6. Custom weeks (enter number)\n' +
    '7. Date range\n\n' +
    'Enter choice:';
  
  var periodResponse = ui.prompt('Step 2/2 - Select Period', periodPrompt, ui.ButtonSet.OK_CANCEL);
  if (periodResponse.getSelectedButton() !== ui.Button.OK) return;
  
  var periodChoice = parseInt(periodResponse.getResponseText());
  var days = 0;
  var useRange = false;
  var startDate, endDate;
  
  switch(periodChoice) {
    case 1: days = 7; break;
    case 2: days = 14; break;
    case 3: days = 28; break;
    case 4: days = 56; break;
    case 5: days = 84; break;
    case 6:
      var weeksResponse = ui.prompt('Custom Weeks', 'Enter number of weeks (1-52):', ui.ButtonSet.OK_CANCEL);
      if (weeksResponse.getSelectedButton() !== ui.Button.OK) return;
      var weeks = parseInt(weeksResponse.getResponseText());
      if (isNaN(weeks) || weeks < 1 || weeks > 52) {
        ui.alert('Invalid', 'Please enter a number between 1 and 52', ui.ButtonSet.OK);
        return;
      }
      days = weeks * 7;
      break;
    case 7:
      var dates = promptDateRange();
      if (!dates) return;
      useRange = true;
      startDate = dates.start;
      endDate = dates.end;
      break;
    default:
      ui.alert('Invalid', 'Invalid period selection', ui.ButtonSet.OK);
      return;
  }
  
  var confirmMsg = '✅ CONFIRM GENERATION\n\n';
  confirmMsg += '📊 Projects: ' + selectedProjects.join(', ') + '\n';
  confirmMsg += '📅 Period: ';
  
  if (useRange) {
    confirmMsg += startDate + ' to ' + endDate;
  } else {
    confirmMsg += 'Last ' + (days / 7) + ' weeks';
  }
  
  confirmMsg += '\n\nGenerate reports?';
  
  if (ui.alert('Confirm', confirmMsg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  ui.alert('Processing', 'Generating reports... This may take a few minutes.', ui.ButtonSet.OK);
  
  try {
    selectedProjects.forEach(function(proj, index) {
      if (proj === 'TRICKY') clearTrickyCaches();
      setCurrentProject(proj);
      
      if (index > 0) Utilities.sleep(5000);
      
      if (useRange) {
        generateReportForDateRange(startDate, endDate);
      } else {
        generateReport(days);
      }
      
      Utilities.sleep(5000);
    });
    
    console.log('Sorting project sheets...');
    Utilities.sleep(5000);
    sortProjectSheets();
    
    ui.alert('✅ Success', 'Reports generated successfully!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Error', 'Error generating reports: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function updateAllProjectsQuick() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    ui.alert('🔐 Token Required', 'Bearer token is not configured.', ui.ButtonSet.OK);
    return;
  }
  
  var result = ui.alert('🔄 Update All Projects', 
    'Update all projects to current date?\n\n• TRICKY: Optimized processing\n• Others: Batch processing\n\nThis may take several minutes.', 
    ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  ui.alert('Processing', 'Updating all projects... This may take several minutes.', ui.ButtonSet.OK);
  
  try {
    console.log('Updating TRICKY separately...');
    try {
      clearTrickyCaches();
      updateProjectDataOptimized('TRICKY');
      console.log('TRICKY updated successfully');
      Utilities.sleep(8000);
    } catch (e) {
      console.error('Error updating TRICKY:', e);
    }
    
    console.log('Updating other projects...');
    var otherProjects = ['MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    var successCount = 0;
    
    otherProjects.forEach(function(proj, index) {
      try {
        if (index > 0) Utilities.sleep(5000);
        updateProjectDataOptimized(proj);
        successCount++;
        console.log(proj + ' updated successfully');
        Utilities.sleep(5000);
      } catch (e) {
        console.error('Error updating ' + proj + ':', e);
      }
    });
    
    console.log('Sorting project sheets...');
    Utilities.sleep(5000);
    sortProjectSheets();
    
    ui.alert('✅ Complete', 'Updated ' + (successCount + 1) + '/8 projects\n\nTRICKY: Optimized processing used', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error during update: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function updateSelectedProjectsQuick() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    ui.alert('🔐 Token Required', 'Bearer token is not configured.', ui.ButtonSet.OK);
    return;
  }
  
  var selected = showMultiChoice('Select Projects to Update:', MENU_PROJECTS);
  if (!selected || selected.length === 0) {
    ui.alert('No Selection', 'No projects selected for update.', ui.ButtonSet.OK);
    return;
  }
  
  var message = 'Update ' + selected.length + ' selected projects?\n\n' + selected.join(', ');
  if (selected.some(function(p) { return p.toUpperCase() === 'TRICKY'; })) {
    message += '\n\n🚀 TRICKY will use optimized processing';
  }
  
  if (ui.alert('🔄 Update Selected Projects', message, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  ui.alert('Processing', 'Updating selected projects...', ui.ButtonSet.OK);
  
  try {
    var successCount = 0;
    
    selected.forEach(function(proj, index) {
      try {
        var projectName = proj.toUpperCase();
        console.log('Updating ' + projectName + '...');
        
        if (projectName === 'TRICKY') {
          clearTrickyCaches();
        }
        
        if (index > 0) Utilities.sleep(5000);
        
        updateProjectDataOptimized(projectName);
        successCount++;
        Utilities.sleep(5000);
        
      } catch (e) {
        console.error('Error updating ' + proj + ':', e);
      }
    });
    
    console.log('Sorting project sheets...');
    Utilities.sleep(5000);
    sortProjectSheets();
    
    ui.alert('✅ Complete', 'Successfully updated ' + successCount + '/' + selected.length + ' projects', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error during update: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function showHelp() {
  var ui = SpreadsheetApp.getUi();
  var help = '📚 CAMPAIGN REPORT HELP\n\n' +
    '📈 GENERATE REPORT:\n' +
    '• Create new reports for selected projects\n' +
    '• Choose from preset periods or custom range\n' +
    '• TRICKY uses optimized processing\n\n' +
    '🔄 UPDATE PROJECTS:\n' +
    '• Refresh existing data to current date\n' +
    '• All Projects - updates everything\n' +
    '• Selected Projects - choose specific ones\n\n' +
    '📅 PERIOD OPTIONS:\n' +
    '• Quick selections: 1, 2, 4, 8, 12 weeks\n' +
    '• Custom weeks: any number 1-52\n' +
    '• Date range: specific start and end dates\n\n' +
    '⚡ TIPS:\n' +
    '• Smaller periods process faster\n' +
    '• Comments are always preserved\n' +
    '• Auto-update runs daily at 5 AM';
  
  ui.alert('Help', help, ui.ButtonSet.OK);
}

function showMultiChoice(title, options) {
  var ui = SpreadsheetApp.getUi();
  var numbered = '';
  for (var i = 0; i < options.length; i++) {
    numbered += (i + 1) + '. ' + options[i] + '\n';
  }
  numbered += '\nEnter numbers separated by commas:';
  
  var result = ui.prompt(title, numbered, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  
  var selections = result.getResponseText().split(',');
  var validSelections = [];
  
  selections.forEach(function(sel) {
    var n = parseInt(sel.trim());
    if (n >= 1 && n <= options.length) {
      validSelections.push(options[n-1]);
    }
  });
  
  return validSelections;
}

function promptDateRange() {
  var ui = SpreadsheetApp.getUi();
  var start = ui.prompt('Start Date', 'Enter date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (start.getSelectedButton() !== ui.Button.OK) return null;
  var end = ui.prompt('End Date', 'Enter date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (end.getSelectedButton() !== ui.Button.OK) return null;
  
  if (!isValidDate(start.getResponseText()) || !isValidDate(end.getResponseText())) {
    ui.alert('Invalid date format');
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

function clearDataWizard() {
  var ui = SpreadsheetApp.getUi();
  var prompt = '🗑️ CLEAR DATA\n\n[A] Clear All Projects\n[S] Clear Single Project\n\nEnter choice:';
  var response = ui.prompt('Clear Data', prompt, ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  var choice = response.getResponseText().toUpperCase().trim();
  
  if (choice === 'A') {
    clearAllProjectsDataOptimized();
  } else if (choice === 'S') {
    var p = showChoice('Select Project:', MENU_PROJECTS);
    if (p) clearProjectAllDataOptimized(MENU_PROJECTS[p-1].toUpperCase());
  } else {
    ui.alert('Invalid choice', 'Please enter A or S', ui.ButtonSet.OK);
  }
}

function showChoice(title, options) {
  var ui = SpreadsheetApp.getUi();
  var numbered = '';
  for (var i = 0; i < options.length; i++) numbered += (i + 1) + '. ' + options[i] + '\n';
  var result = ui.prompt(title, numbered + '\nEnter number:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  var choice = parseInt(result.getResponseText());
  return (choice >= 1 && choice <= options.length) ? choice : null;
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
      } catch (e) {
        console.error('Error clearing ' + proj + ':', e);
      }
    });
    
    ui.alert('Success', 'Cleared ' + successCount + '/' + projects.length + ' projects', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error clearing data: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function clearProjectAllDataOptimized(projectName) {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Clear ' + projectName + ' Data', 'Clear all ' + projectName + ' data? Comments preserved.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    if (projectName === 'TRICKY') clearTrickyCaches();
    clearProjectDataSilent(projectName);
    ui.alert('Success', projectName + ' data cleared', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error clearing ' + projectName + ': ' + e.toString(), ui.ButtonSet.OK);
  }
}

function debugSingleProject() {
  var p = showChoice('Select Project to Debug:', MENU_PROJECTS);
  if (p) {
    var projectName = MENU_PROJECTS[p-1].toUpperCase();
    if (projectName === 'TRICKY') clearTrickyCaches();
    setCurrentProject(projectName);
    debugReportGeneration();
  }
}

function refreshSettingsDialog() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var settings = refreshSettingsFromSheet();
    syncTriggersWithSettings();
    ui.alert('Settings Refreshed', '✅ Settings and triggers synchronized', ui.ButtonSet.OK);
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
  
  var triggers = ScriptApp.getProjectTriggers();
  var cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
  var trickyTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateTricky'; });
  var othersTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateOthers'; });
  
  var message = '📊 SYSTEM STATUS\n\n';
  message += '🔐 Bearer Token: ' + tokenStatus + '\n';
  message += '💾 Auto Cache: ' + cacheStatus + '\n';
  message += '🔄 Auto Update: ' + updateStatus + '\n\n';
  
  message += '⏰ TRIGGERS:\n';
  message += '• Cache: ' + (cacheTrigger ? '✅ 2:00 AM' : '❌ Not set') + '\n';
  message += '• TRICKY: ' + (trickyTrigger ? '✅ 5:00 AM' : '❌ Not set') + '\n';
  message += '• Others: ' + (othersTrigger ? '✅ 5:15 AM' : '❌ Not set') + '\n\n';
  
  message += '💡 Simplified schedule with TRICKY optimization';
  
  ui.alert('System Status', message, ui.ButtonSet.OK);
}

function recreateAllTriggers() {
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.alert('🔄 Recreate Triggers', 
    'Recreate all automation triggers?\n\n⏰ Schedule:\n• Cache: 2:00 AM\n• TRICKY: 5:00 AM (optimized)\n• Others: 5:15 AM (batch)', 
    ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  try {
    clearAllTriggers();
    
    var cacheEnabled = isAutoCacheEnabled();
    var updateEnabled = isAutoUpdateEnabled();
    
    if (cacheEnabled) {
      ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
    }
    
    if (updateEnabled) {
      ScriptApp.newTrigger('autoUpdateTricky').timeBased().atHour(5).nearMinute(0).everyDays(1).create();
      ScriptApp.newTrigger('autoUpdateOthers').timeBased().atHour(5).nearMinute(15).everyDays(1).create();
    }
    
    ui.alert('✅ Triggers Recreated', 'All triggers recreated with simplified schedule', ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert('❌ Error', 'Error recreating triggers: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function clearAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { 
    ScriptApp.deleteTrigger(t); 
  });
}

function syncTriggersWithSettings() {
  try {
    var settings = loadSettingsFromSheet();
    var triggers = ScriptApp.getProjectTriggers();
    
    var cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
    var trickyTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateTricky'; });
    var othersTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateOthers'; });
    
    if (settings.automation.autoCache && !cacheTrigger) {
      ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
    } else if (!settings.automation.autoCache && cacheTrigger) {
      ScriptApp.deleteTrigger(cacheTrigger);
    }
    
    if (settings.automation.autoUpdate) {
      if (!trickyTrigger) {
        ScriptApp.newTrigger('autoUpdateTricky').timeBased().atHour(5).nearMinute(0).everyDays(1).create();
      }
      if (!othersTrigger) {
        ScriptApp.newTrigger('autoUpdateOthers').timeBased().atHour(5).nearMinute(15).everyDays(1).create();
      }
    } else {
      if (trickyTrigger) ScriptApp.deleteTrigger(trickyTrigger);
      if (othersTrigger) ScriptApp.deleteTrigger(othersTrigger);
    }
    
  } catch (e) {
    console.error('Error syncing triggers:', e);
  }
}

function saveAllCommentsToCache() {
  var ui = SpreadsheetApp.getUi();
  try {
    var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    var successCount = 0;
    
    projects.forEach(function(proj) {
      try {
        saveProjectCommentsManual(proj);
        successCount++;
      } catch (e) {
        console.error('Error saving ' + proj + ' comments:', e);
      }
    });
    
    ui.alert('Success', 'Saved comments for ' + successCount + '/' + projects.length + ' projects', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error saving comments: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function saveProjectCommentsManual(projectName) {
  var config = getProjectConfig(projectName);
  var spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  var sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('No data found in ' + projectName + ' sheet');
  }
  
  var cache = new CommentCache(projectName);
  cache.syncCommentsFromSheet();
}

function quickAPICheckAll() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    ui.alert('🔐 Token Required', 'Bearer token not configured', ui.ButtonSet.OK);
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
        results += '❌ ' + proj + ': No data\n';
      } else {
        var count = raw.data.analytics.richStats.stats.length;
        results += '✅ ' + proj + ': ' + count + ' records\n';
      }
    } catch (e) {
      results += '❌ ' + proj + ': ' + e.toString().substring(0, 30) + '...\n';
    }
  });
  
  ui.alert('API Check Complete', results, ui.ButtonSet.OK);
}

function appsDbWizard() {
  var ui = SpreadsheetApp.getUi();
  
  if (CURRENT_PROJECT !== 'TRICKY') {
    if (ui.alert('Apps Database - TRICKY Only', 'Switch to TRICKY project?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    setCurrentProject('TRICKY');
  }
  
  var action = showChoice('📱 Apps Database Management', [
    'View Cache Status',
    'Refresh Apps Database', 
    'View Sample Data',
    'Clear Cache'
  ]);
  if (!action) return;
  
  switch(action) {
    case 1: showAppsDbStatus(); break;
    case 2: refreshAppsDatabase(); break;
    case 3: showAppsDbSample(); break;
    case 4: clearAppsDbCache(); break;
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
      message += '• Update Needed: ' + (appsDb.shouldUpdateCache() ? 'YES' : 'NO') + '\n\n';
      
      message += 'SAMPLE ENTRIES:\n';
      for (var i = 0; i < Math.min(3, bundleIds.length); i++) {
        var bundleId = bundleIds[i];
        var app = cache[bundleId];
        message += '• ' + bundleId + ' → ' + app.publisher + ' ' + app.appName + '\n';
      }
    }
    
    ui.alert('Apps Database Status', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error checking status: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function showAppsDbSample() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var appsDb = new AppsDatabase('TRICKY');
    var cache = appsDb.loadFromCache();
    var bundleIds = Object.keys(cache);
    
    if (bundleIds.length === 0) {
      ui.alert('No Data', 'Apps Database cache is empty', ui.ButtonSet.OK);
      return;
    }
    
    var message = '📱 APPS DATABASE SAMPLE\n\n';
    
    for (var i = 0; i < Math.min(5, bundleIds.length); i++) {
      var bundleId = bundleIds[i];
      var app = cache[bundleId];
      message += bundleId + '\n  → ' + app.publisher + ' ' + app.appName + '\n\n';
    }
    
    ui.alert('Apps Database Sample', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error showing sample: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function clearAppsDbCache() {
  var ui = SpreadsheetApp.getUi();
  
  if (ui.alert('Clear Apps Database Cache', 'Clear cached app data?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    var appsDb = new AppsDatabase('TRICKY');
    if (appsDb.cacheSheet && appsDb.cacheSheet.getLastRow() > 1) {
      appsDb.cacheSheet.deleteRows(2, appsDb.cacheSheet.getLastRow() - 1);
      clearTrickyCaches();
      ui.alert('Success', 'Apps Database cache cleared', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', 'Error clearing cache: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function openGitHubRepo() {
  var ui = SpreadsheetApp.getUi();
  var githubUrl = 'https://github.com/Poxagronka/pivot-table';
  
  var htmlOutput = HtmlService.createHtmlOutput(
    '<script>window.open("' + githubUrl + '", "_blank"); google.script.host.close();</script>'
  ).setWidth(400).setHeight(300);
  
  ui.showModalDialog(htmlOutput, 'Opening GitHub Repository...');
}