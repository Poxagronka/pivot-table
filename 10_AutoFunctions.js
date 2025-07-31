function autoCacheAllProjects() {
  if (!isAutoCacheEnabled()) {
    return;
  }
  
  try {
    ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'INCENT_TRAFFIC', 'OVERALL'].forEach(function(proj) {
      try {
        cacheProjectComments(proj);
      } catch (e) {
        console.error(`Error caching ${proj}:`, e);
      }
    });
  } catch (e) {
    console.error('AUTO CACHE FATAL ERROR:', e);
  }
}

function autoUpdateAllProjects() {
  const overallStartTime = Date.now();
  
  if (!isAutoUpdateEnabled()) {
    return;
  }
  
  try {
    const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'INCENT_TRAFFIC', 'OVERALL'];
    let successCount = 0;
    let totalRecords = 0;
    let totalRows = 0;
    
    projects.forEach(function(proj) {
      const projectStartTime = Date.now();
      try {
        const result = updateProjectData(proj);
        if (result) {
          successCount++;
          totalRecords += result.recordCount || 0;
          totalRows += result.rowCount || 0;
        }
      } catch (e) {
        console.error(`Error updating ${proj}:`, e);
      }
    });
    
    if (successCount > 0) {
      try {
        sortProjectSheets();
      } catch (e) {
        console.error('Error sorting sheets after auto-update:', e);
      }
    }
    
    const totalTime = Date.now() - overallStartTime;
    logInfo(`Auto Update (${successCount}/${projects.length} projects)`, totalRecords, totalRows, totalTime);
    
  } catch (e) {
    console.error('AUTO UPDATE FATAL ERROR:', e);
  }
}

function cacheProjectComments(projectName) {
  projectName = projectName.toUpperCase();
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return;
  }
  
  const cache = new CommentCache(projectName);
  cache.syncCommentsFromSheet();
}

function updateProjectData(projectName) {
  projectName = projectName.toUpperCase();
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return null;
  }
  
  let earliestDate = null;
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'WEEK') {
      const weekRange = data[i][1];
      const startStr = weekRange.split(' - ')[0];
      const startDate = new Date(startStr);
      if (!earliestDate || startDate < earliestDate) earliestDate = startDate;
    }
  }
  
  if (!earliestDate) {
    return null;
  }
  
  const today = new Date();
  const dayOfWeek = today.getDay();
  let endDate = new Date(today);
  
  if (dayOfWeek === 0) {
    endDate.setDate(today.getDate() - 1);
  } else {
    endDate.setDate(today.getDate() - dayOfWeek);
  }
  
  const dateRange = {
    from: formatDateForAPI(earliestDate),
    to: formatDateForAPI(endDate)
  };
  
  const raw = fetchProjectCampaignData(projectName, dateRange);
  
  if (!raw.data?.analytics?.richStats?.stats?.length) {
    return null;
  }
  
  const processed = processProjectApiData(projectName, raw);
  
  if (Object.keys(processed).length === 0) {
    return null;
  }
  
  clearProjectDataSilent(projectName);
  
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  
  let rowCount = 0;
  const recordCount = raw.data.analytics.richStats.stats.length;
  
  try {
    if (projectName === 'OVERALL') {
      rowCount = createOverallPivotTable(processed);
    } else if (projectName === 'INCENT_TRAFFIC') {
      rowCount = createIncentTrafficPivotTable(processed);
    } else {
      rowCount = createEnhancedPivotTable(processed);
    }
    const cache = new CommentCache(projectName);
    cache.applyCommentsToSheet();
    
    return { recordCount, rowCount };
  } finally {
    setCurrentProject(originalProject);
  }
}

function saveAllCommentsManual() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'INCENT_TRAFFIC', 'OVERALL'];
    let successCount = 0;
    
    projects.forEach(function(proj) {
      try {
        saveProjectCommentsManual(proj);
        successCount++;
      } catch (e) {
        console.error(`Error saving ${proj} comments:`, e);
      }
    });
    
    if (successCount === projects.length) {
      ui.alert('Success', 'All project comments saved successfully!', ui.ButtonSet.OK);
    } else {
      ui.alert('Partial Success', `Saved comments for ${successCount} of ${projects.length} projects.`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', 'Error saving comments: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function saveProjectCommentsManual(projectName) {
  projectName = projectName.toUpperCase();
  
  try {
    const config = getProjectConfig(projectName);
    const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      throw new Error(`No data found in ${projectName} sheet`);
    }
    
    const cache = new CommentCache(projectName);
    cache.syncCommentsFromSheet();
    
  } catch (e) {
    console.error(`Error saving ${projectName} comments:`, e);
    throw e;
  }
}

function showAutomationStatus() {
  const ui = SpreadsheetApp.getUi();
  
  const cacheEnabled = isAutoCacheEnabled();
  const updateEnabled = isAutoUpdateEnabled();
  
  const triggers = ScriptApp.getProjectTriggers();
  const cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
  const updateTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateAllProjects'; });
  
  let msg = '📊 AUTOMATION STATUS\n\n';
  
  msg += '💾 AUTO CACHE:\n';
  if (cacheEnabled && cacheTrigger) {
    msg += '✅ Enabled - Runs every hour\n• Caches comments from all projects (including INCENT_TRAFFIC)\n• Collapses all row groups after caching\n';
  } else if (cacheEnabled && !cacheTrigger) {
    msg += '⚠️ Enabled but trigger missing\n• Please use Settings sheet to fix\n';
  } else {
    msg += '❌ Disabled\n• Comments must be saved manually\n';
  }
  
  msg += '\n🔄 AUTO UPDATE:\n';
  if (updateEnabled && updateTrigger) {
    msg += '✅ Enabled - Runs daily at 5:00 AM\n• Updates all project data (including INCENT_TRAFFIC)\n• Includes previous week data starting from Tuesday\n• Preserves all comments\n• Sorts project sheets after update\n';
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
    
    ScriptApp.newTrigger('autoCacheAllProjects').timeBased().everyHours(1).create();
    saveSettingToSheet('automation.autoCache', true);
    
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
    
    ScriptApp.newTrigger('autoUpdateAllProjects').timeBased().atHour(5).everyDays(1).create();
    saveSettingToSheet('automation.autoUpdate', true);
    
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
    
  } catch (e) {
    console.error('Failed to disable auto update:', e);
    throw e;
  }
}

function syncTriggersWithSettings() {
  try {
    const settings = loadSettingsFromSheet();
    const triggers = ScriptApp.getProjectTriggers();
    
    const cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
    const updateTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateAllProjects'; });
    
    if (settings.automation.autoCache && !cacheTrigger) {
      ScriptApp.newTrigger('autoCacheAllProjects').timeBased().everyHours(1).create();
    } else if (!settings.automation.autoCache && cacheTrigger) {
      ScriptApp.deleteTrigger(cacheTrigger);
    }
    
    if (settings.automation.autoUpdate && !updateTrigger) {
      ScriptApp.newTrigger('autoUpdateAllProjects').timeBased().atHour(5).everyDays(1).create();
    } else if (!settings.automation.autoUpdate && updateTrigger) {
      ScriptApp.deleteTrigger(updateTrigger);
    }
    
  } catch (e) {
    console.error('Error syncing triggers with settings:', e);
  }
}

function onSettingsChange() {
  clearSettingsCache();
  syncTriggersWithSettings();
}