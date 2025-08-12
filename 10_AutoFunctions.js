/**
 * Auto Functions - Unified automation system
 */

const ALL_PROJECTS = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'INCENT_TRAFFIC', 'OVERALL'];

const TRIGGER_CONFIG = {
  autoCache: {
    handler: 'autoCacheAllProjects',
    create: () => ScriptApp.newTrigger('autoCacheAllProjects').timeBased().everyHours(1).create(),
    settingKey: 'automation.autoCache',
    description: 'Every hour - saves comments automatically'
  },
  autoUpdate: {
    handler: 'autoUpdateAllProjects',
    create: () => ScriptApp.newTrigger('autoUpdateAllProjects').timeBased().atHour(5).everyDays(1).create(),
    settingKey: 'automation.autoUpdate',
    description: 'Daily at 5:00 AM - updates all projects data'
  }
};

// Main auto functions
function autoCacheAllProjects() {
  executeAutoOperation('CACHE', (proj) => {
    cacheProjectComments(proj);
  });
}

function autoUpdateAllProjects() {
  executeAutoOperation('UPDATE', (proj) => {
    updateProjectData(proj);
  }, true);
}

function executeAutoOperation(operationType, operation, sortAfter = false) {
  console.log(`=== AUTO ${operationType} STARTED ===`);
  
  const settingKey = operationType === 'CACHE' ? 'automation.autoCache' : 'automation.autoUpdate';
  if (!getSettingValue(settingKey)) {
    console.log(`Auto ${operationType.toLowerCase()} is disabled in settings, skipping`);
    return;
  }
  
  let successCount = 0;
  let failedProjects = [];
  
  ALL_PROJECTS.forEach(proj => {
    try {
      console.log(`${operationType}: Processing ${proj}...`);
      operation(proj);
      successCount++;
    } catch (e) {
      console.error(`${operationType} error for ${proj}:`, e);
      failedProjects.push(proj);
    }
  });
  
  if (sortAfter && successCount > 0) {
    try {
      sortProjectSheets();
      console.log('Project sheets sorted after auto-operation');
    } catch (e) {
      console.error('Error sorting sheets:', e);
    }
  }
  
  const status = failedProjects.length > 0 ? 
    `completed with errors (failed: ${failedProjects.join(', ')})` : 
    'completed successfully';
  
  console.log(`=== AUTO ${operationType} ${status.toUpperCase()}: ${successCount}/${ALL_PROJECTS.length} projects ===`);
}

// Project operations
function cacheProjectComments(projectName) {
  projectName = projectName.toUpperCase();
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log(`${projectName}: No data to cache`);
    return;
  }
  
  const cache = new CommentCache(projectName);
  cache.syncCommentsFromSheet();
  console.log(`${projectName}: Comments cached`);
}

// Manual operations
function saveAllCommentsToCache() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const results = ALL_PROJECTS.map(proj => {
      try {
        saveProjectCommentsManual(proj);
        return { project: proj, success: true };
      } catch (e) {
        console.error(`Error saving ${proj} comments:`, e);
        return { project: proj, success: false, error: e.toString() };
      }
    });
    
    const successCount = results.filter(r => r.success).length;
    const message = successCount === ALL_PROJECTS.length ? 
      'All project comments have been saved to cache.' :
      `Saved comments for ${successCount} of ${ALL_PROJECTS.length} projects.`;
    
    ui.alert('Save Comments', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error saving comments: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function saveProjectCommentsManual(projectName) {
  projectName = projectName.toUpperCase();
  console.log(`Saving comments for ${projectName}...`);
  
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error(`No data found in ${projectName} sheet`);
  }
  
  const cache = new CommentCache(projectName);
  cache.syncCommentsFromSheet();
  console.log(`✅ ${projectName} comments saved successfully`);
}

// Automation status
function showAutomationStatus() {
  const ui = SpreadsheetApp.getUi();
  
  const triggers = ScriptApp.getProjectTriggers();
  const status = {
    cache: {
      enabled: isAutoCacheEnabled(),
      trigger: triggers.find(t => t.getHandlerFunction() === 'autoCacheAllProjects')
    },
    update: {
      enabled: isAutoUpdateEnabled(),
      trigger: triggers.find(t => t.getHandlerFunction() === 'autoUpdateAllProjects')
    }
  };
  
  const formatStatus = (name, config) => {
    if (config.enabled && config.trigger) return `✅ Enabled - ${TRIGGER_CONFIG[name].description}`;
    if (config.enabled && !config.trigger) return '⚠️ Enabled but trigger missing';
    return '❌ Disabled';
  };
  
  const msg = `📊 AUTOMATION STATUS\n\n` +
    `💾 AUTO CACHE:\n${formatStatus('autoCache', status.cache)}\n\n` +
    `🔄 AUTO UPDATE:\n${formatStatus('autoUpdate', status.update)}\n\n` +
    `⏱️ ACTIVE TRIGGERS: ${triggers.length} total\n\n` +
    `💡 TIP: Use Settings sheet to enable/disable automation`;
  
  ui.alert('Automation Status', msg, ui.ButtonSet.OK);
}

// Trigger management
function enableAutoCache() { manageTrigger('autoCache', true); }
function disableAutoCache() { manageTrigger('autoCache', false); }
function enableAutoUpdate() { manageTrigger('autoUpdate', true); }
function disableAutoUpdate() { manageTrigger('autoUpdate', false); }

function manageTrigger(triggerType, enable) {
  try {
    const config = TRIGGER_CONFIG[triggerType];
    
    // Remove existing triggers
    ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === config.handler)
      .forEach(t => ScriptApp.deleteTrigger(t));
    
    // Create new trigger if enabling
    if (enable) {
      config.create();
      console.log(`${triggerType} enabled`);
    } else {
      console.log(`${triggerType} disabled`);
    }
    
    // Save to settings
    saveSettingToSheet(config.settingKey, enable);
    
  } catch (e) {
    console.error(`Failed to ${enable ? 'enable' : 'disable'} ${triggerType}:`, e);
    throw e;
  }
}

// Settings helpers
function getSettingValue(path) {
  try {
    const settings = loadSettingsFromSheet();
    const parts = path.split('.');
    let value = settings;
    for (const part of parts) {
      value = value[part];
    }
    return value;
  } catch (e) {
    return false;
  }
}

// Sync triggers with settings
function syncTriggersWithSettings() {
  try {
    const settings = loadSettingsFromSheet();
    const triggers = ScriptApp.getProjectTriggers();
    
    Object.entries(TRIGGER_CONFIG).forEach(([key, config]) => {
      const enabled = getSettingValue(config.settingKey);
      const trigger = triggers.find(t => t.getHandlerFunction() === config.handler);
      
      if (enabled && !trigger) {
        config.create();
        console.log(`Created ${key} trigger`);
      } else if (!enabled && trigger) {
        ScriptApp.deleteTrigger(trigger);
        console.log(`Deleted ${key} trigger`);
      }
    });
    
    console.log('Triggers synchronized with Settings sheet');
  } catch (e) {
    console.error('Error syncing triggers with settings:', e);
  }
}

// Event handler
function onSettingsChange() {
  clearSettingsCache();
  syncTriggersWithSettings();
}