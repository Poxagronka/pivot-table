function autoCacheAllProjects() {
  console.log('=== AUTO CACHE STARTED ===');
  
  if (!isAutoCacheEnabled()) {
    console.log('Auto cache is disabled in settings, skipping');
    return;
  }
  
  try {
    ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'].forEach(function(proj, index) {
      try {
        if (index > 0) {
          Utilities.sleep(2000);
        }
        console.log(`Caching ${proj}...`);
        cacheProjectComments(proj);
      } catch (e) {
        console.error(`Error caching ${proj}:`, e);
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
    updateProjectDataOptimized('TRICKY');
    console.log('=== AUTO UPDATE TRICKY COMPLETED ===');
  } catch (e) {
    console.error('AUTO UPDATE TRICKY ERROR:', e);
  }
}

function autoUpdateMoloco() {
  console.log('=== AUTO UPDATE MOLOCO STARTED ===');
  if (!isAutoUpdateEnabled()) {
    console.log('Auto update is disabled, skipping MOLOCO');
    return;
  }
  
  try {
    updateProjectDataOptimized('MOLOCO');
    console.log('=== AUTO UPDATE MOLOCO COMPLETED ===');
  } catch (e) {
    console.error('AUTO UPDATE MOLOCO ERROR:', e);
  }
}

function autoUpdateRegular() {
  console.log('=== AUTO UPDATE REGULAR STARTED ===');
  if (!isAutoUpdateEnabled()) {
    console.log('Auto update is disabled, skipping REGULAR');
    return;
  }
  
  try {
    updateProjectDataOptimized('REGULAR');
    console.log('=== AUTO UPDATE REGULAR COMPLETED ===');
  } catch (e) {
    console.error('AUTO UPDATE REGULAR ERROR:', e);
  }
}

function autoUpdateGoogleAds() {
  console.log('=== AUTO UPDATE GOOGLE_ADS STARTED ===');
  if (!isAutoUpdateEnabled()) {
    console.log('Auto update is disabled, skipping GOOGLE_ADS');
    return;
  }
  
  try {
    updateProjectDataOptimized('GOOGLE_ADS');
    console.log('=== AUTO UPDATE GOOGLE_ADS COMPLETED ===');
  } catch (e) {
    console.error('AUTO UPDATE GOOGLE_ADS ERROR:', e);
  }
}

function autoUpdateApplovin() {
  console.log('=== AUTO UPDATE APPLOVIN STARTED ===');
  if (!isAutoUpdateEnabled()) {
    console.log('Auto update is disabled, skipping APPLOVIN');
    return;
  }
  
  try {
    updateProjectDataOptimized('APPLOVIN');
    console.log('=== AUTO UPDATE APPLOVIN COMPLETED ===');
  } catch (e) {
    console.error('AUTO UPDATE APPLOVIN ERROR:', e);
  }
}

function autoUpdateMintegral() {
  console.log('=== AUTO UPDATE MINTEGRAL STARTED ===');
  if (!isAutoUpdateEnabled()) {
    console.log('Auto update is disabled, skipping MINTEGRAL');
    return;
  }
  
  try {
    updateProjectDataOptimized('MINTEGRAL');
    console.log('=== AUTO UPDATE MINTEGRAL COMPLETED ===');
  } catch (e) {
    console.error('AUTO UPDATE MINTEGRAL ERROR:', e);
  }
}

function autoUpdateIncent() {
  console.log('=== AUTO UPDATE INCENT STARTED ===');
  if (!isAutoUpdateEnabled()) {
    console.log('Auto update is disabled, skipping INCENT');
    return;
  }
  
  try {
    updateProjectDataOptimized('INCENT');
    console.log('=== AUTO UPDATE INCENT COMPLETED ===');
  } catch (e) {
    console.error('AUTO UPDATE INCENT ERROR:', e);
  }
}

function autoUpdateOverall() {
  console.log('=== AUTO UPDATE OVERALL STARTED ===');
  if (!isAutoUpdateEnabled()) {
    console.log('Auto update is disabled, skipping OVERALL');
    return;
  }
  
  try {
    updateProjectDataOptimized('OVERALL');
    console.log('=== AUTO UPDATE OVERALL COMPLETED ===');
  } catch (e) {
    console.error('AUTO UPDATE OVERALL ERROR:', e);
  }
}

function updateProjectDataOptimized(projectName) {
  console.log(`Starting optimized update for ${projectName}`);
  
  var config = getProjectConfig(projectName);
  var spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  var sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log(`${projectName}: No existing data to update`);
    return;
  }
  
  var cache = new CommentCache(projectName);
  cache.syncCommentsFromSheet();
  console.log(`${projectName}: Comments cached`);
  
  var earliestDate = findEarliestWeekDate(sheet);
  if (!earliestDate) {
    console.log(`${projectName}: No week data found`);
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
  
  console.log(`${projectName}: Fetching ${dateRange.from} to ${dateRange.to}`);
  
  var raw = fetchProjectCampaignData(projectName, dateRange);
  
  if (!raw.data?.analytics?.richStats?.stats?.length) {
    console.log(`${projectName}: No API data`);
    return;
  }
  
  var processed = processProjectApiData(projectName, raw);
  
  if (Object.keys(processed).length === 0) {
    console.log(`${projectName}: No valid processed data`);
    return;
  }
  
  clearProjectDataFast(projectName);
  
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
  
  console.log(`${projectName}: Update completed successfully`);
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

function clearProjectDataFast(projectName) {
  var config = getProjectConfig(projectName);
  var spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  var sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (sheet) {
    sheet.clear();
    console.log(`${projectName}: Sheet cleared`);
  }
}

function cacheProjectComments(projectName) {
  var config = getProjectConfig(projectName);
  var spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  var sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log(`${projectName}: No data to cache`);
    return;
  }
  
  var cache = new CommentCache(projectName);
  cache.syncCommentsFromSheet();
  
  console.log(`${projectName}: Comments cached`);
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
        console.error(`Error saving ${proj} comments:`, e);
      }
    });
    
    if (successCount === projects.length) {
      ui.alert('Success', 'All project comments have been saved to cache.', ui.ButtonSet.OK);
    } else {
      ui.alert('Partial Success', `Saved comments for ${successCount} of ${projects.length} projects.`, ui.ButtonSet.OK);
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
    throw new Error(`No data found in ${projectName} sheet`);
  }
  
  var cache = new CommentCache(projectName);
  cache.syncCommentsFromSheet();
}

function showAutomationStatus() {
  var ui = SpreadsheetApp.getUi();
  
  var cacheEnabled = isAutoCacheEnabled();
  var updateEnabled = isAutoUpdateEnabled();
  
  var triggers = ScriptApp.getProjectTriggers();
  var updateTriggers = getUpdateTriggers();
  
  var msg = '📊 AUTOMATION STATUS\n\n';
  
  msg += '💾 AUTO CACHE:\n';
  var cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
  if (cacheEnabled && cacheTrigger) {
    msg += '✅ Enabled - Daily at 2:00 AM\n';
  } else if (cacheEnabled && !cacheTrigger) {
    msg += '⚠️ Enabled but trigger missing\n';
  } else {
    msg += '❌ Disabled\n';
  }
  
  msg += '\n🔄 AUTO UPDATE (EXACT TIME TRIGGERS):\n';
  if (updateEnabled && updateTriggers.length === 8) {
    msg += '✅ Enabled - Exact times:\n';
    msg += '• TRICKY: 5:00 AM\n';
    msg += '• MOLOCO: 5:00 AM\n';
    msg += '• REGULAR: 5:00 AM\n';
    msg += '• GOOGLE_ADS: 5:00 AM\n';
    msg += '• APPLOVIN: 5:00 AM\n';
    msg += '• MINTEGRAL: 5:00 AM\n';
    msg += '• INCENT: 6:00 AM\n';
    msg += '• OVERALL: 6:00 AM\n';
  } else if (updateEnabled && updateTriggers.length > 0) {
    msg += `⚠️ Partially configured (${updateTriggers.length}/8 triggers)\n`;
  } else if (updateEnabled) {
    msg += '⚠️ Enabled but triggers missing\n';
  } else {
    msg += '❌ Disabled\n';
  }
  
  msg += `\n⏱️ ACTIVE TRIGGERS: ${triggers.length} total\n`;
  msg += `• Cache: ${triggers.filter(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; }).length}\n`;
  msg += `• Update: ${updateTriggers.length}\n`;
  
  ui.alert('Automation Status', msg, ui.ButtonSet.OK);
}

function getUpdateTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var updateFunctions = [
    'autoUpdateTricky', 'autoUpdateMoloco', 'autoUpdateRegular', 'autoUpdateGoogleAds',
    'autoUpdateApplovin', 'autoUpdateMintegral', 'autoUpdateIncent', 'autoUpdateOverall'
  ];
  
  return triggers.filter(function(t) {
    return updateFunctions.includes(t.getHandlerFunction());
  });
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
    clearAllUpdateTriggers();
    createUpdateTriggers();
    saveSettingToSheet('automation.autoUpdate', true);
    
    console.log('Auto update enabled with exact time triggers');
  } catch (e) {
    console.error('Failed to enable auto update:', e);
    throw e;
  }
}

function disableAutoUpdate() {
  try {
    clearAllUpdateTriggers();
    saveSettingToSheet('automation.autoUpdate', false);
    
    console.log('Auto update disabled');
  } catch (e) {
    console.error('Failed to disable auto update:', e);
    throw e;
  }
}

function clearAllUpdateTriggers() {
  var updateFunctions = [
    'autoUpdateTricky', 'autoUpdateMoloco', 'autoUpdateRegular', 'autoUpdateGoogleAds',
    'autoUpdateApplovin', 'autoUpdateMintegral', 'autoUpdateIncent', 'autoUpdateOverall',
    'autoUpdateAllProjects'
  ];
  
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return updateFunctions.includes(t.getHandlerFunction()); })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  
  console.log('Cleared all update triggers');
}

function createUpdateTriggers() {
  var schedule = [
    { func: 'autoUpdateTricky', hour: 5 },
    { func: 'autoUpdateMoloco', hour: 5 },
    { func: 'autoUpdateRegular', hour: 5 },
    { func: 'autoUpdateGoogleAds', hour: 5 },
    { func: 'autoUpdateApplovin', hour: 5 },
    { func: 'autoUpdateMintegral', hour: 5 },
    { func: 'autoUpdateIncent', hour: 6 },
    { func: 'autoUpdateOverall', hour: 6 }
  ];
  
  schedule.forEach(function(item, index) {
    ScriptApp.newTrigger(item.func)
      .timeBased()
      .everyDays(1)
      .atHour(item.hour)
      .create();
    
    console.log(`Created trigger for ${item.func} at ${item.hour}:00`);
  });
  
  console.log('All update triggers created');
}

function syncTriggersWithSettings() {
  try {
    var settings = loadSettingsFromSheet();
    var triggers = ScriptApp.getProjectTriggers();
    
    var cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
    var updateTriggers = getUpdateTriggers();
    
    if (settings.automation.autoCache && !cacheTrigger) {
      ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
      console.log('Created auto cache trigger');
    } else if (!settings.automation.autoCache && cacheTrigger) {
      ScriptApp.deleteTrigger(cacheTrigger);
      console.log('Deleted auto cache trigger');
    }
    
    if (settings.automation.autoUpdate && updateTriggers.length !== 8) {
      clearAllUpdateTriggers();
      createUpdateTriggers();
      console.log('Created update triggers with exact times');
    } else if (!settings.automation.autoUpdate && updateTriggers.length > 0) {
      clearAllUpdateTriggers();
      console.log('Deleted all update triggers');
    }
    
    console.log('Triggers synchronized with Settings sheet');
  } catch (e) {
    console.error('Error syncing triggers with settings:', e);
  }
}

function onSettingsChange() {
  clearSettingsCache();
  syncTriggersWithSettings();
}

function recreateAllTriggers() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    console.log('Recreating all triggers...');
    
    clearAllUpdateTriggers();
    
    var cacheEnabled = isAutoCacheEnabled();
    var updateEnabled = isAutoUpdateEnabled();
    
    if (cacheEnabled) {
      ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
      console.log('Cache trigger recreated');
    }
    
    if (updateEnabled) {
      createUpdateTriggers();
      console.log('Update triggers recreated');
    }
    
    ui.alert('✅ Triggers Recreated', 'All triggers have been recreated with exact timing!\n\n⏰ New schedule:\n• Cache: 2:00 AM\n• Updates: 5:00-6:10 AM (10min apart)', ui.ButtonSet.OK);
    
  } catch (e) {
    console.error('Error recreating triggers:', e);
    ui.alert('❌ Error', 'Error recreating triggers: ' + e.toString(), ui.ButtonSet.OK);
  }
}