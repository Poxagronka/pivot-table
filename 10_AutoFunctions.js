/**
 * Auto Functions - ОБНОВЛЕНО: добавлен Overall
 */

function autoCacheAllProjects() {
  console.log('=== AUTO CACHE STARTED ===');
  try {
    ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'].forEach(proj => {
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

function cacheProjectComments(projectName) {
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log(`${projectName}: No data to cache`);
    return;
  }
  
  const cache = new CommentCache(projectName);
  cache.syncCommentsFromSheet();
  
  console.log(`${projectName}: Comments cached (groups unchanged)`);
}

function autoUpdateAllProjects() {
  console.log('=== AUTO UPDATE STARTED ===');
  try {
    ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'].forEach(proj => {
      try {
        console.log(`Updating ${proj}...`);
        updateProjectData(proj);
      } catch (e) {
        console.error(`Error updating ${proj}:`, e);
      }
    });
    console.log('=== AUTO UPDATE COMPLETED ===');
  } catch (e) {
    console.error('AUTO UPDATE FATAL ERROR:', e);
  }
}

function updateProjectData(projectName) {
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log(`${projectName}: No existing data to update`);
    return;
  }
  
  let earliestDate = null;
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'WEEK') {
      const weekRange = data[i][1];
      const [startStr] = weekRange.split(' - ');
      const startDate = new Date(startStr);
      if (!earliestDate || startDate < earliestDate) earliestDate = startDate;
    }
  }
  
  if (!earliestDate) {
    console.log(`${projectName}: No week data found`);
    return;
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
  
  console.log(`${projectName}: Fetching data from ${dateRange.from} to ${dateRange.to}`);
  
  const raw = fetchProjectCampaignData(projectName, dateRange);
  
  if (!raw.data?.analytics?.richStats?.stats?.length) {
    console.log(`${projectName}: No data returned from API`);
    return;
  }
  
  const processed = processProjectApiData(projectName, raw);
  
  if (Object.keys(processed).length === 0) {
    console.log(`${projectName}: No valid data to process`);
    return;
  }
  
  clearProjectDataSilent(projectName);
  
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    if (projectName === 'OVERALL') {
      createOverallPivotTable(processed);
    } else {
      createEnhancedPivotTable(processed);
    }
    const cache = new CommentCache(projectName);
    cache.applyCommentsToSheet();
  } finally {
    setCurrentProject(originalProject);
  }
  
  console.log(`${projectName}: Update completed`);
}

function saveAllCommentsToCache() {
  const ui = SpreadsheetApp.getUi();
  try {
    const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    let successCount = 0;
    
    projects.forEach(proj => {
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
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error(`No data found in ${projectName} sheet`);
  }
  
  const cache = new CommentCache(projectName);
  cache.syncCommentsFromSheet();
}

function showAutomationStatus() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const cacheEnabled = props.getProperty(PROPERTY_KEYS.AUTO_CACHE_ENABLED) === 'true';
  const updateEnabled = props.getProperty(PROPERTY_KEYS.AUTO_UPDATE_ENABLED) === 'true';
  
  const triggers = ScriptApp.getProjectTriggers();
  const cacheTrigger = triggers.find(t => t.getHandlerFunction() === 'autoCacheAllProjects');
  const updateTrigger = triggers.find(t => t.getHandlerFunction() === 'autoUpdateAllProjects');
  
  let msg = '📊 AUTOMATION STATUS\n\n';
  
  msg += '💾 AUTO CACHE:\n';
  if (cacheEnabled && cacheTrigger) {
    msg += '✅ Enabled - Runs daily at 2:00 AM\n• Caches comments from all projects\n• Collapses all row groups after caching\n';
  } else if (cacheEnabled && !cacheTrigger) {
    msg += '⚠️ Enabled but trigger missing\n• Please disable and re-enable to fix\n';
  } else {
    msg += '❌ Disabled\n• Comments must be saved manually\n';
  }
  
  msg += '\n🔄 AUTO UPDATE:\n';
  if (updateEnabled && updateTrigger) {
    msg += '✅ Enabled - Runs every Monday at 5:00 AM\n• Updates all project data\n• Includes previous complete week\n• Preserves all comments\n';
  } else if (updateEnabled && !updateTrigger) {
    msg += '⚠️ Enabled but trigger missing\n• Please disable and re-enable to fix\n';
  } else {
    msg += '❌ Disabled\n• Data must be updated manually\n';
  }
  
  msg += `\n⏱️ ACTIVE TRIGGERS:\n• Total triggers: ${triggers.length}\n• Cache triggers: ${triggers.filter(t => t.getHandlerFunction() === 'autoCacheAllProjects').length}\n• Update triggers: ${triggers.filter(t => t.getHandlerFunction() === 'autoUpdateAllProjects').length}`;
  
  ui.alert('Automation Status', msg, ui.ButtonSet.OK);
}