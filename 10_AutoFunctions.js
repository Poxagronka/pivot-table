function autoCacheAllProjects() {
  console.log('=== AUTO CACHE STARTED ===');
  
  if (!isAutoCacheEnabled()) {
    console.log('Auto cache is disabled in settings, skipping');
    return;
  }
  
  try {
    ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'].forEach(function(proj, index) {
      try {
        if (index > 0) Utilities.sleep(2000);
        console.log('Caching ' + proj + '...');
        cacheProjectComments(proj);
      } catch (e) {
        console.error('Error caching ' + proj + ':', e);
      }
    });
    console.log('=== AUTO CACHE COMPLETED ===');
  } catch (e) {
    console.error('AUTO CACHE FATAL ERROR:', e);
  }
}

function autoUpdateTricky() {
  console.log('=== AUTO UPDATE TRICKY STARTED ===');
  
  if (!isAutoUpdateEnabled()) {
    console.log('Auto update is disabled, skipping TRICKY');
    return;
  }
  
  try {
    clearTrickyCaches();
    updateProjectDataOptimized('TRICKY');
    console.log('=== AUTO UPDATE TRICKY COMPLETED ===');
  } catch (e) {
    console.error('AUTO UPDATE TRICKY ERROR:', e);
  }
}

function autoUpdateOthers() {
  console.log('=== AUTO UPDATE OTHERS STARTED ===');
  
  if (!isAutoUpdateEnabled()) {
    console.log('Auto update is disabled, skipping other projects');
    return;
  }
  
  try {
    var projects = ['MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    var successCount = 0;
    
    projects.forEach(function(proj, index) {
      try {
        if (index > 0) Utilities.sleep(5000);
        console.log('Updating ' + proj + '...');
        updateProjectDataOptimized(proj);
        successCount++;
        console.log(proj + ' updated successfully');
      } catch (e) {
        console.error('Error updating ' + proj + ':', e);
      }
    });
    
    try {
      console.log('Sorting project sheets...');
      Utilities.sleep(3000);
      sortProjectSheets();
    } catch (e) {
      console.error('Error sorting sheets:', e);
    }
    
    console.log('=== AUTO UPDATE OTHERS COMPLETED (' + successCount + '/' + projects.length + ') ===');
  } catch (e) {
    console.error('AUTO UPDATE OTHERS FATAL ERROR:', e);
  }
}

function cacheProjectComments(projectName) {
  var config = getProjectConfig(projectName);
  var spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  var sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log(projectName + ': No data to cache');
    return;
  }
  
  var cache = new CommentCache(projectName);
  cache.syncCommentsFromSheet();
  
  console.log(projectName + ': Comments cached');
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
    
    if (successCount === projects.length) {
      ui.alert('Success', 'All project comments have been saved to cache.', ui.ButtonSet.OK);
    } else {
      ui.alert('Partial Success', 'Saved comments for ' + successCount + ' of ' + projects.length + ' projects.', ui.ButtonSet.OK);
    }
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

function showAutomationStatus() {
  var ui = SpreadsheetApp.getUi();
  
  var cacheEnabled = isAutoCacheEnabled();
  var updateEnabled = isAutoUpdateEnabled();
  
  var triggers = ScriptApp.getProjectTriggers();
  var cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
  var trickyTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateTricky'; });
  var othersTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateOthers'; });
  
  var msg = '📊 AUTOMATION STATUS\n\n';
  
  msg += '💾 AUTO CACHE:\n';
  if (cacheEnabled && cacheTrigger) {
    msg += '✅ Enabled - Daily at 2:00 AM\n';
  } else if (cacheEnabled && !cacheTrigger) {
    msg += '⚠️ Enabled but trigger missing\n';
  } else {
    msg += '❌ Disabled\n';
  }
  
  msg += '\n🔄 AUTO UPDATE:\n';
  if (updateEnabled && trickyTrigger && othersTrigger) {
    msg += '✅ Enabled - Simplified schedule:\n';
    msg += '• TRICKY: 5:00 AM (optimized)\n';
    msg += '• Others: 5:15 AM (batch)\n';
  } else if (updateEnabled && (!trickyTrigger || !othersTrigger)) {
    msg += '⚠️ Enabled but triggers incomplete\n';
  } else {
    msg += '❌ Disabled\n';
  }
  
  msg += '\n⏱️ ACTIVE TRIGGERS: ' + triggers.length + ' total\n';
  msg += '• Cache: ' + (cacheTrigger ? '1' : '0') + '\n';
  msg += '• TRICKY: ' + (trickyTrigger ? '1' : '0') + '\n';
  msg += '• Others: ' + (othersTrigger ? '1' : '0') + '\n';
  
  ui.alert('Automation Status', msg, ui.ButtonSet.OK);
}

function enableAutoCache() {
  try {
    ScriptApp.getProjectTriggers()
      .filter(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; })
      .forEach(function(t) { ScriptApp.deleteTrigger(t); });
    
    ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
    saveSettingToSheet('automation.autoCache', true);
    
    console.log('Auto cache enabled');
  } catch (e) {
    console.error('Failed to enable auto cache:', e);
    throw e;
  }
}

function disableAutoCache() {
  try {
    ScriptApp.getProjectTriggers()
      .filter(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; })
      .forEach(function(t) { ScriptApp.deleteTrigger(t); });
    
    saveSettingToSheet('automation.autoCache', false);
    
    console.log('Auto cache disabled');
  } catch (e) {
    console.error('Failed to disable auto cache:', e);
    throw e;
  }
}

function enableAutoUpdate() {
  try {
    clearUpdateTriggers();
    
    ScriptApp.newTrigger('autoUpdateTricky').timeBased().atHour(5).nearMinute(0).everyDays(1).create();
    ScriptApp.newTrigger('autoUpdateOthers').timeBased().atHour(5).nearMinute(15).everyDays(1).create();
    
    saveSettingToSheet('automation.autoUpdate', true);
    
    console.log('Auto update enabled with simplified triggers');
  } catch (e) {
    console.error('Failed to enable auto update:', e);
    throw e;
  }
}

function disableAutoUpdate() {
  try {
    clearUpdateTriggers();
    saveSettingToSheet('automation.autoUpdate', false);
    
    console.log('Auto update disabled');
  } catch (e) {
    console.error('Failed to disable auto update:', e);
    throw e;
  }
}

function clearUpdateTriggers() {
  var updateFunctions = [
    'autoUpdateTricky', 'autoUpdateOthers',
    'autoUpdateMoloco', 'autoUpdateRegular', 'autoUpdateGoogleAds',
    'autoUpdateApplovin', 'autoUpdateMintegral', 'autoUpdateIncent', 'autoUpdateOverall',
    'autoUpdateAllProjects'
  ];
  
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return updateFunctions.includes(t.getHandlerFunction()); })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  
  console.log('Cleared all update triggers');
}

function recreateAllTriggers() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    console.log('Recreating all triggers...');
    
    clearUpdateTriggers();
    
    var cacheEnabled = isAutoCacheEnabled();
    var updateEnabled = isAutoUpdateEnabled();
    
    if (cacheEnabled) {
      ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
      console.log('Cache trigger recreated');
    }
    
    if (updateEnabled) {
      ScriptApp.newTrigger('autoUpdateTricky').timeBased().atHour(5).nearMinute(0).everyDays(1).create();
      ScriptApp.newTrigger('autoUpdateOthers').timeBased().atHour(5).nearMinute(15).everyDays(1).create();
      console.log('Update triggers recreated');
    }
    
    ui.alert('✅ Triggers Recreated', 'All triggers have been recreated with simplified schedule!', ui.ButtonSet.OK);
    
  } catch (e) {
    console.error('Error recreating triggers:', e);
    ui.alert('❌ Error', 'Error recreating triggers: ' + e.toString(), ui.ButtonSet.OK);
  }
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
      console.log('Created auto cache trigger');
    } else if (!settings.automation.autoCache && cacheTrigger) {
      ScriptApp.deleteTrigger(cacheTrigger);
      console.log('Deleted auto cache trigger');
    }
    
    if (settings.automation.autoUpdate) {
      if (!trickyTrigger) {
        ScriptApp.newTrigger('autoUpdateTricky').timeBased().atHour(5).nearMinute(0).everyDays(1).create();
        console.log('Created TRICKY update trigger');
      }
      if (!othersTrigger) {
        ScriptApp.newTrigger('autoUpdateOthers').timeBased().atHour(5).nearMinute(15).everyDays(1).create();
        console.log('Created others update trigger');
      }
    } else {
      if (trickyTrigger) {
        ScriptApp.deleteTrigger(trickyTrigger);
        console.log('Deleted TRICKY update trigger');
      }
      if (othersTrigger) {
        ScriptApp.deleteTrigger(othersTrigger);
        console.log('Deleted others update trigger');
      }
    }
    
    clearOldUpdateTriggers();
    
    console.log('Triggers synchronized with Settings sheet');
  } catch (e) {
    console.error('Error syncing triggers with settings:', e);
  }
}

function clearOldUpdateTriggers() {
  var oldFunctions = [
    'autoUpdateMoloco', 'autoUpdateRegular', 'autoUpdateGoogleAds',
    'autoUpdateApplovin', 'autoUpdateMintegral', 'autoUpdateIncent', 
    'autoUpdateOverall', 'autoUpdateAllProjects'
  ];
  
  var cleared = 0;
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return oldFunctions.includes(t.getHandlerFunction()); })
    .forEach(function(t) { 
      ScriptApp.deleteTrigger(t); 
      cleared++;
    });
  
  if (cleared > 0) {
    console.log('Cleared ' + cleared + ' old update triggers');
  }
}

function onSettingsChange() {
  clearSettingsCache();
  syncTriggersWithSettings();
}

function updateProjectDataOptimized(projectName) {
  if (projectName === 'TRICKY') {
    updateProjectDataOptimizedTricky();
    return;
  }
  
  updateProjectDataOptimizedStandard(projectName);
}

function updateProjectDataOptimizedTricky() {
  console.log('=== STARTING TRICKY OPTIMIZED UPDATE ===');
  
  var config = getProjectConfig('TRICKY');
  var spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  var sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log('TRICKY: No existing data to update');
    return;
  }
  
  console.log('TRICKY: Initializing optimized cache...');
  var trickyCache = initTrickyOptimizedCache();
  
  console.log('TRICKY: Caching comments...');
  var cache = new CommentCache('TRICKY');
  cache.syncCommentsFromSheet();
  
  console.log('TRICKY: Finding earliest week date...');
  var earliestDate = findEarliestWeekDate(sheet);
  if (!earliestDate) {
    console.log('TRICKY: No week data found');
    return;
  }
  
  var today = new Date();
  var dayOfWeek = today.getDay();
  var endDate = new Date(today);
  
  if (dayOfWeek === 0) {
    endDate.setDate(today.getDate() - 1);
  } else {
    endDate.setDate(today.getDate() - dayOfWeek);
  }
  
  var dateRange = {
    from: formatDateForAPI(earliestDate),
    to: formatDateForAPI(endDate)
  };
  
  console.log('TRICKY: Fetching optimized data ' + dateRange.from + ' to ' + dateRange.to);
  
  var raw = fetchProjectCampaignData('TRICKY', dateRange);
  
  if (!raw.data?.analytics?.richStats?.stats?.length) {
    console.log('TRICKY: No API data');
    return;
  }
  
  console.log('TRICKY: Processing API data with optimizations...');
  var originalProject = CURRENT_PROJECT;
  setCurrentProject('TRICKY');
  
  try {
    var processed = processApiData(raw);
    
    if (Object.keys(processed).length === 0) {
      console.log('TRICKY: No valid processed data');
      return;
    }
    
    console.log('TRICKY: Recreating sheet...');
    recreateSheetFast(spreadsheet, config.SHEET_NAME);
    
    console.log('TRICKY: Creating optimized pivot table...');
    createEnhancedPivotTable(processed);
    
    console.log('TRICKY: Applying cached comments...');
    cache.applyCommentsToSheet();
    
    console.log('=== TRICKY OPTIMIZED UPDATE COMPLETED ===');
    
  } finally {
    setCurrentProject(originalProject);
  }
}

function updateProjectDataOptimizedStandard(projectName) {
  console.log('Starting standard update for ' + projectName);
  
  var config = getProjectConfig(projectName);
  var spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  var sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log(projectName + ': No existing data to update');
    return;
  }
  
  var cache = new CommentCache(projectName);
  cache.syncCommentsFromSheet();
  console.log(projectName + ': Comments cached');
  
  var earliestDate = findEarliestWeekDate(sheet);
  if (!earliestDate) {
    console.log(projectName + ': No week data found');
    return;
  }
  
  var today = new Date();
  var dayOfWeek = today.getDay();
  var endDate = new Date(today);
  
  if (dayOfWeek === 0) {
    endDate.setDate(today.getDate() - 1);
  } else {
    endDate.setDate(today.getDate() - dayOfWeek);
  }
  
  var dateRange = {
    from: formatDateForAPI(earliestDate),
    to: formatDateForAPI(endDate)
  };
  
  console.log(projectName + ': Fetching ' + dateRange.from + ' to ' + dateRange.to);
  
  var raw = fetchProjectCampaignData(projectName, dateRange);
  
  if (!raw.data?.analytics?.richStats?.stats?.length) {
    console.log(projectName + ': No API data');
    return;
  }
  
  var processed = processProjectApiData(projectName, raw);
  
  if (Object.keys(processed).length === 0) {
    console.log(projectName + ': No valid processed data');
    return;
  }
  
  recreateSheetFast(spreadsheet, config.SHEET_NAME);
  
  var originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    if (projectName === 'OVERALL') {
      createOverallPivotTable(processed);
    } else {
      createEnhancedPivotTable(processed);
    }
    cache.applyCommentsToSheet();
  } finally {
    setCurrentProject(originalProject);
  }
  
  console.log(projectName + ': Update completed successfully');
}

function findEarliestWeekDate(sheet) {
  var data = sheet.getDataRange().getValues();
  var earliestDate = null;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'WEEK') {
      var weekRange = data[i][1];
      var startStr = weekRange.split(' - ')[0];
      var startDate = new Date(startStr);
      if (!earliestDate || startDate < earliestDate) {
        earliestDate = startDate;
      }
    }
  }
  
  return earliestDate;
}