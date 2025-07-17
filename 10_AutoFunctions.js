/**
 * Auto Functions - ОБНОВЛЕНО: всегда сортирует листы + учитывает день недели для предыдущей недели
 */

function autoCacheAllProjects() {
  console.log('=== AUTO CACHE STARTED ===');
  
  if (!isAutoCacheEnabled()) {
    console.log('Auto cache is disabled in settings, skipping');
    return;
  }
  
  try {
    ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'].forEach(function(proj) {
      try {
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

function autoUpdateAllProjects() {
  console.log('=== AUTO UPDATE STARTED ===');
  
  if (!isAutoUpdateEnabled()) {
    console.log('Auto update is disabled in settings, skipping');
    return;
  }
  
  try {
    var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    var successCount = 0;
    
    projects.forEach(function(proj) {
      try {
        console.log(`Updating ${proj}...`);
        updateProjectData(proj);
        successCount++;
      } catch (e) {
        console.error(`Error updating ${proj}:`, e);
      }
    });
    
    // ИЗМЕНЕНО: всегда сортируем листы после автообновления (убрали условие successCount > 1)
    if (successCount > 0) {
      try {
        sortProjectSheets();
        console.log('Project sheets sorted after auto-update');
      } catch (e) {
        console.error('Error sorting sheets after auto-update:', e);
      }
    }
    
    console.log(`=== AUTO UPDATE COMPLETED: ${successCount}/${projects.length} projects updated ===`);
  } catch (e) {
    console.error('AUTO UPDATE FATAL ERROR:', e);
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
  
  console.log(`${projectName}: Comments cached (groups unchanged)`);
}

function updateProjectData(projectName) {
  var config = getProjectConfig(projectName);
  var spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  var sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log(`${projectName}: No existing data to update`);
    return;
  }
  
  var earliestDate = null;
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'WEEK') {
      var weekRange = data[i][1];
      var startStr = weekRange.split(' - ')[0];
      var startDate = new Date(startStr);
      if (!earliestDate || startDate < earliestDate) earliestDate = startDate;
    }
  }
  
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
  
  console.log(`${projectName}: Fetching data from ${dateRange.from} to ${dateRange.to}`);
  
  var raw = fetchProjectCampaignData(projectName, dateRange);
  
  if (!raw.data?.analytics?.richStats?.stats?.length) {
    console.log(`${projectName}: No data returned from API`);
    return;
  }
  
  // ОБНОВЛЕНО: processProjectApiData теперь автоматически учитывает день недели
  var processed = processProjectApiData(projectName, raw);
  
  if (Object.keys(processed).length === 0) {
    console.log(`${projectName}: No valid data to process`);
    return;
  }
  
  clearProjectDataSilent(projectName);
  
  var originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    if (projectName === 'OVERALL') {
      createOverallPivotTable(processed);
    } else {
      createEnhancedPivotTable(processed);
    }
    var cache = new CommentCache(projectName);
    cache.applyCommentsToSheet();
  } finally {
    setCurrentProject(originalProject);
  }
  
  console.log(`${projectName}: Update completed`);
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
  console.log(`Saving comments for ${projectName}...`);
  
  try {
    var config = getProjectConfig(projectName);
    var spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
    var sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error(`No data found in ${projectName} sheet`);
    }
    
    var cache = new CommentCache(projectName);
    cache.syncCommentsFromSheet();
    
    console.log(`✅ ${projectName} comments saved successfully`);
    
  } catch (e) {
    console.error(`❌ Error saving ${projectName} comments:`, e);
    console.log('Error details:', e.toString());
    throw e;
  }
}

function showAutomationStatus() {
  var ui = SpreadsheetApp.getUi();
  
  var cacheEnabled = isAutoCacheEnabled();
  var updateEnabled = isAutoUpdateEnabled();
  
  var triggers = ScriptApp.getProjectTriggers();
  var cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
  var updateTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateAllProjects'; });
  
  var msg = '📊 AUTOMATION STATUS\n\n';
  
  msg += '💾 AUTO CACHE:\n';
  if (cacheEnabled && cacheTrigger) {
    msg += '✅ Enabled - Runs daily at 2:00 AM\n• Caches comments from all projects\n• Collapses all row groups after caching\n';
  } else if (cacheEnabled && !cacheTrigger) {
    msg += '⚠️ Enabled but trigger missing\n• Please use Settings sheet to fix\n';
  } else {
    msg += '❌ Disabled\n• Comments must be saved manually\n';
  }
  
  msg += '\n🔄 AUTO UPDATE:\n';
  if (updateEnabled && updateTrigger) {
    msg += '✅ Enabled - Runs daily at 5:00 AM\n• Updates all project data\n• Includes previous week data starting from Tuesday\n• Preserves all comments\n• Sorts project sheets after update\n';
  } else if (updateEnabled && !updateTrigger) {
    msg += '⚠️ Enabled but trigger missing\n• Please use Settings sheet to fix\n';
  } else {
    msg += '❌ Disabled\n• Data must be updated manually\n';
  }
  
  msg += `\n⏱️ ACTIVE TRIGGERS:\n• Total triggers: ${triggers.length}\n• Cache triggers: ${triggers.filter(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; }).length}\n• Update triggers: ${triggers.filter(function(t) { return t.getHandlerFunction() === 'autoUpdateAllProjects'; }).length}`;
  
  msg += '\n\n💡 TIP: Use Settings sheet to enable/disable automation';
  
  ui.alert('Automation Status', msg, ui.ButtonSet.OK);
}

function enableAutoCache() {
  try {
    ScriptApp.getProjectTriggers()
      .filter(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; })
      .forEach(function(t) { ScriptApp.deleteTrigger(t); });
    
    ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
    saveSettingToSheet('automation.autoCache', true);
    
    console.log('Auto cache enabled and saved to Settings sheet');
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
    
    console.log('Auto cache disabled and saved to Settings sheet');
  } catch (e) {
    console.error('Failed to disable auto cache:', e);
    throw e;
  }
}

function enableAutoUpdate() {
  try {
    ScriptApp.getProjectTriggers()
      .filter(function(t) { return t.getHandlerFunction() === 'autoUpdateAllProjects'; })
      .forEach(function(t) { ScriptApp.deleteTrigger(t); });
    
    // ИЗМЕНЕНО: автообновление каждый день в 5:00 AM
    ScriptApp.newTrigger('autoUpdateAllProjects').timeBased().atHour(5).everyDays(1).create();
    saveSettingToSheet('automation.autoUpdate', true);
    
    console.log('Auto update enabled and saved to Settings sheet');
  } catch (e) {
    console.error('Failed to enable auto update:', e);
    throw e;
  }
}

function disableAutoUpdate() {
  try {
    ScriptApp.getProjectTriggers()
      .filter(function(t) { return t.getHandlerFunction() === 'autoUpdateAllProjects'; })
      .forEach(function(t) { ScriptApp.deleteTrigger(t); });
    
    saveSettingToSheet('automation.autoUpdate', false);
    
    console.log('Auto update disabled and saved to Settings sheet');
  } catch (e) {
    console.error('Failed to disable auto update:', e);
    throw e;
  }
}

function syncTriggersWithSettings() {
  try {
    var settings = loadSettingsFromSheet();
    var triggers = ScriptApp.getProjectTriggers();
    
    var cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
    var updateTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateAllProjects'; });
    
    if (settings.automation.autoCache && !cacheTrigger) {
      ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
      console.log('Created auto cache trigger');
    } else if (!settings.automation.autoCache && cacheTrigger) {
      ScriptApp.deleteTrigger(cacheTrigger);
      console.log('Deleted auto cache trigger');
    }
    
    if (settings.automation.autoUpdate && !updateTrigger) {
      // ИЗМЕНЕНО: создаем триггер на каждый день в 5:00 AM
      ScriptApp.newTrigger('autoUpdateAllProjects').timeBased().atHour(5).everyDays(1).create();
      console.log('Created auto update trigger');
    } else if (!settings.automation.autoUpdate && updateTrigger) {
      ScriptApp.deleteTrigger(updateTrigger);
      console.log('Deleted auto update trigger');
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