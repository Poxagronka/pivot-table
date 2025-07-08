/**
 * Settings Functions - ОБНОВЛЕНО: добавлен Incent
 */

function showClearDataDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('🗑️ Clear All Data', 'YES = Select project\nNO = Clear ALL projects\nCANCEL = Exit\n\nComments will be preserved.', ui.ButtonSet.YES_NO_CANCEL);
  
  if (result === ui.Button.CANCEL) return;
  if (result === ui.Button.YES) showProjectSelectionForClear();
  else clearAllProjectsData();
}

function showProjectSelectionForClear() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Select Project to Clear', 'Enter project name:\n• TRICKY\n• MOLOCO\n• REGULAR\n• GOOGLE_ADS\n• APPLOVIN\n• MINTEGRAL\n• INCENT', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const projectName = response.getResponseText().toUpperCase().trim();
    if (['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT'].includes(projectName)) {
      clearProjectAllData(projectName);
    } else {
      ui.alert('Invalid Project', 'Please enter a valid project name.', ui.ButtonSet.OK);
    }
  }
}

function clearAllProjectsData() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Confirm Clear All', 'Clear data from ALL projects:\n• Tricky\n• Moloco\n• Regular\n• Google_Ads\n• Applovin\n• Mintegral\n• Incent\n\nComments preserved. Continue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT'];
    let successCount = 0;
    
    projects.forEach(proj => {
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
  const ui = SpreadsheetApp.getUi();
  if (ui.alert(`Clear ${projectName} Data`, `Clear all ${projectName} data?\n\nComments will be preserved.`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    clearProjectDataSilent(projectName);
    ui.alert('Success', `${projectName} data cleared. Comments preserved.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Error clearing ${projectName}: ${e.toString()}`, ui.ButtonSet.OK);
  }
}

function showQuickAPICheckDialog() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('🔍 Quick API Check', 'Enter:\n• TRICKY\n• MOLOCO\n• REGULAR\n• GOOGLE_ADS\n• APPLOVIN\n• MINTEGRAL\n• INCENT\n• ALL (check all)', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const input = response.getResponseText().toUpperCase().trim();
    if (input === 'ALL') checkAllProjectsAPI();
    else if (['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT'].includes(input)) checkProjectAPI(input);
    else ui.alert('Invalid Input', 'Please enter a valid project name or ALL.', ui.ButtonSet.OK);
  }
}

function checkProjectAPI(projectName) {
  const ui = SpreadsheetApp.getUi();
  try {
    setCurrentProject(projectName);
    const dateRange = getDateRange(7);
    const raw = fetchCampaignData(dateRange);
    
    if (!raw.data?.analytics?.richStats?.stats?.length) {
      ui.alert(`${projectName} API Check`, `❌ No data for last 7 days.\n\nPossible issues:\n• No active campaigns\n• Expired API token\n• Network config incorrect`, ui.ButtonSet.OK);
    } else {
      const count = raw.data.analytics.richStats.stats.length;
      ui.alert(`${projectName} API Check`, `✅ API working!\n\n• Records: ${count}\n• Period: Last 7 days\n• Status: Connected`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert(`${projectName} API Check`, `❌ API Error:\n\n${e.toString()}\n\nCheck:\n• Internet connection\n• API token\n• Project config`, ui.ButtonSet.OK);
  }
}

function checkAllProjectsAPI() {
  const ui = SpreadsheetApp.getUi();
  const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT'];
  let results = '🔍 API CHECK RESULTS\n\n';
  
  projects.forEach(proj => {
    try {
      setCurrentProject(proj);
      const dateRange = getDateRange(7);
      const raw = fetchCampaignData(dateRange);
      
      if (!raw.data?.analytics?.richStats?.stats?.length) {
        results += `❌ ${proj}: No data\n`;
      } else {
        const count = raw.data.analytics.richStats.stats.length;
        results += `✅ ${proj}: ${count} records\n`;
      }
    } catch (e) {
      results += `❌ ${proj}: ${e.toString().substring(0, 50)}...\n`;
    }
  });
  
  ui.alert('API Check Complete', results, ui.ButtonSet.OK);
}

function testAPIWithDateRange(startDate, endDate) {
  const ui = SpreadsheetApp.getUi();
  const choice = showChoice('Select Project for API Test:', ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT']);
  if (!choice) return;
  
  const projectName = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT'][choice-1];
  
  try {
    setCurrentProject(projectName);
    const dateRange = { from: startDate, to: endDate };
    const raw = fetchCampaignData(dateRange);
    
    if (!raw.data?.analytics?.richStats?.stats?.length) {
      ui.alert(`${projectName} API Test`, `❌ No data for ${startDate} to ${endDate}.\n\nAPI works but no campaigns match filters.`, ui.ButtonSet.OK);
    } else {
      const count = raw.data.analytics.richStats.stats.length;
      ui.alert(`${projectName} API Test`, `✅ API test successful!\n\n• Records: ${count}\n• Period: ${startDate} to ${endDate}\n• Project: ${projectName}`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert(`${projectName} API Test`, `❌ API Error:\n\n${e.toString()}`, ui.ButtonSet.OK);
  }
}

function showProjectSettings() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const cacheEnabled = props.getProperty(PROPERTY_KEYS.AUTO_CACHE_ENABLED) === 'true';
  const updateEnabled = props.getProperty(PROPERTY_KEYS.AUTO_UPDATE_ENABLED) === 'true';
  
  const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT'];
  let message = '🎯 PROJECT CONFIGURATION\n\n';
  
  projects.forEach(proj => {
    const target = getTargetEROAS(proj);
    const thresholds = getGrowthThresholds(proj);
    const config = getProjectConfig(proj);
    const apiConfig = getProjectApiConfig(proj);
    
    message += `📊 ${proj}:\n`;
    message += `• Target eROAS: ${target}%\n`;
    message += `• Growth: Healthy(${thresholds.healthyGrowth.minSpendChange}%/${thresholds.healthyGrowth.minProfitChange}%), Scaling(${thresholds.scalingDown.maxSpendChange}%)\n`;
    message += `• Network HID: ${apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.join(', ')}\n`;
    message += `• Campaign Filter: ${apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH || 'NO FILTER'}\n\n`;
  });
  
  message += '⚙️ AUTOMATION:\n';
  message += `• Auto Cache: ${cacheEnabled ? '✅ Enabled (2 AM daily)' : '❌ Disabled'}\n`;
  message += `• Auto Update: ${updateEnabled ? '✅ Enabled (Monday 5 AM)' : '❌ Disabled'}\n\n`;
  message += '📝 FEATURES: Week/Campaign comments, Auto-collapse, Project-specific thresholds';
  
  ui.alert('Project Settings', message, ui.ButtonSet.OK);
}

function showGrowthThresholdDetails(projectName) {
  const ui = SpreadsheetApp.getUi();
  const t = getGrowthThresholds(projectName);
  
  const message = `📊 ${projectName} GROWTH THRESHOLDS\n\n` +
    `🟢 ПОЗИТИВНЫЕ: Healthy Growth (Spend ≥${t.healthyGrowth.minSpendChange}% AND Profit ≥${t.healthyGrowth.minProfitChange}%), Efficiency Improvement (спенд падает, профит растет), переход из убытка в прибыль\n` +
    `🔴 КРИТИЧЕСКИЕ: Inefficient Growth (Profit ≤${t.inefficientGrowth.maxProfitChange}%), переход из прибыли в убыток\n` +
    `🟠 ПРЕДУПРЕЖДАЮЩИЕ: Declining Efficiency (спенд растет/стабилен, профит падает умеренно)\n` +
    `🔵 СОКРАЩЕНИЕ: Scaling Down (Spend ≤${t.scalingDown.maxSpendChange}%) - Efficient/Moderate/Problematic\n` +
    `🟡 УМЕРЕННЫЕ: Moderate Growth/Decline, Minimal Growth, различные паттерны\n` +
    `⚪ СТАБИЛЬНЫЕ: Минимальные изменения в любую сторону`;
  
  ui.alert(`${projectName} Growth Thresholds`, message, ui.ButtonSet.OK);
}

function quickAdjustThresholds(projectName) {
  const ui = SpreadsheetApp.getUi();
  const current = getGrowthThresholds(projectName);
  
  const result = ui.alert(`Quick Adjust - ${projectName}`, 
    `Current:\n• Healthy: Spend >${current.healthyGrowth.minSpendChange}%, Profit >${current.healthyGrowth.minProfitChange}%\n• Scaling: Spend <${current.scalingDown.maxSpendChange}%\n\nYES = Adjust Healthy\nNO = Adjust Scaling`, 
    ui.ButtonSet.YES_NO_CANCEL);
  
  if (result === ui.Button.CANCEL) return;
  
  const newThresholds = { ...current };
  
  if (result === ui.Button.YES) {
    const spendResp = ui.prompt('Healthy Growth - Spend', `Min spend change % (current: ${current.healthyGrowth.minSpendChange}%):`, ui.ButtonSet.OK_CANCEL);
    if (spendResp.getSelectedButton() === ui.Button.OK) {
      const val = parseInt(spendResp.getResponseText());
      if (!isNaN(val) && val >= 0 && val <= 100) newThresholds.healthyGrowth.minSpendChange = val;
    }
    
    const profitResp = ui.prompt('Healthy Growth - Profit', `Min profit change % (current: ${current.healthyGrowth.minProfitChange}%):`, ui.ButtonSet.OK_CANCEL);
    if (profitResp.getSelectedButton() === ui.Button.OK) {
      const val = parseInt(profitResp.getResponseText());
      if (!isNaN(val) && val >= -50 && val <= 100) {
        newThresholds.healthyGrowth.minProfitChange = val;
        newThresholds.inefficientGrowth.minSpendChange = newThresholds.healthyGrowth.minSpendChange;
      }
    }
  } else {
    const scalingResp = ui.prompt('Scaling Down', `Max spend change % (current: ${current.scalingDown.maxSpendChange}%):`, ui.ButtonSet.OK_CANCEL);
    if (scalingResp.getSelectedButton() === ui.Button.OK) {
      const val = parseInt(scalingResp.getResponseText());
      if (!isNaN(val) && val >= -100 && val <= 0) newThresholds.scalingDown.maxSpendChange = val;
    }
  }
  
  setGrowthThresholds(projectName, newThresholds);
  ui.alert('✅ Updated', `${projectName} thresholds updated!`, ui.ButtonSet.OK);
}

function showAutoCacheSettings() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const isEnabled = props.getProperty(PROPERTY_KEYS.AUTO_CACHE_ENABLED) === 'true';
  
  const result = ui.alert('💾 Auto Cache Settings', 
    `Auto-cache: ${isEnabled ? '✅ ENABLED' : '❌ DISABLED'}\n\nCaches comments at 2:00 AM daily and collapses groups.\n\n${isEnabled ? 'DISABLE' : 'ENABLE'} auto-cache?`, 
    ui.ButtonSet.YES_NO_CANCEL);
  
  if (result === ui.Button.YES) {
    isEnabled ? disableAutoCache() : enableAutoCache();
  }
}

function enableAutoCache() {
  const ui = SpreadsheetApp.getUi();
  try {
    ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'autoCacheAllProjects')
      .forEach(t => ScriptApp.deleteTrigger(t));
    
    ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
    PropertiesService.getScriptProperties().setProperty(PROPERTY_KEYS.AUTO_CACHE_ENABLED, 'true');
    
    ui.alert('Auto Cache Enabled', '✅ Auto-cache enabled!\n\n• Daily at 2:00 AM\n• All projects cached\n• Groups collapsed after caching', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to enable auto-cache: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function disableAutoCache() {
  const ui = SpreadsheetApp.getUi();
  try {
    ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'autoCacheAllProjects')
      .forEach(t => ScriptApp.deleteTrigger(t));
    
    PropertiesService.getScriptProperties().setProperty(PROPERTY_KEYS.AUTO_CACHE_ENABLED, 'false');
    ui.alert('Auto Cache Disabled', '❌ Auto-cache disabled.\n\nComments no longer cached automatically.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to disable auto-cache: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function showAutoUpdateSettings() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const isEnabled = props.getProperty(PROPERTY_KEYS.AUTO_UPDATE_ENABLED) === 'true';
  
  const result = ui.alert('🔄 Auto Update Settings', 
    `Auto-update: ${isEnabled ? '✅ ENABLED' : '❌ DISABLED'}\n\nUpdates all projects every Monday at 5:00 AM.\n\n${isEnabled ? 'DISABLE' : 'ENABLE'} auto-update?`, 
    ui.ButtonSet.YES_NO_CANCEL);
  
  if (result === ui.Button.YES) {
    isEnabled ? disableAutoUpdate() : enableAutoUpdate();
  }
}

function enableAutoUpdate() {
  const ui = SpreadsheetApp.getUi();
  try {
    ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'autoUpdateAllProjects')
      .forEach(t => ScriptApp.deleteTrigger(t));
    
    ScriptApp.newTrigger('autoUpdateAllProjects').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(5).create();
    PropertiesService.getScriptProperties().setProperty(PROPERTY_KEYS.AUTO_UPDATE_ENABLED, 'true');
    
    ui.alert('Auto Update Enabled', '✅ Auto-update enabled!\n\n• Every Monday at 5:00 AM\n• All projects updated\n• Comments preserved', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to enable auto-update: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function disableAutoUpdate() {
  const ui = SpreadsheetApp.getUi();
  try {
    ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'autoUpdateAllProjects')
      .forEach(t => ScriptApp.deleteTrigger(t));
    
    PropertiesService.getScriptProperties().setProperty(PROPERTY_KEYS.AUTO_UPDATE_ENABLED, 'false');
    ui.alert('Auto Update Disabled', '❌ Auto-update disabled.\n\nData no longer updated automatically.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to disable auto-update: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function getCompleteAutomationStatus() {
  const props = PropertiesService.getScriptProperties();
  const cacheEnabled = props.getProperty(PROPERTY_KEYS.AUTO_CACHE_ENABLED) === 'true';
  const updateEnabled = props.getProperty(PROPERTY_KEYS.AUTO_UPDATE_ENABLED) === 'true';
  
  const triggers = ScriptApp.getProjectTriggers();
  const cacheTrigger = triggers.find(t => t.getHandlerFunction() === 'autoCacheAllProjects');
  const updateTrigger = triggers.find(t => t.getHandlerFunction() === 'autoUpdateAllProjects');
  
  let msg = '📊 AUTOMATION STATUS\n\n💾 AUTO CACHE:\n';
  
  if (cacheEnabled && cacheTrigger) {
    msg += '✅ Enabled - Daily at 2:00 AM\n• Caches all project comments\n• Collapses groups after caching\n';
  } else if (cacheEnabled && !cacheTrigger) {
    msg += '⚠️ Enabled but trigger missing\n• Disable and re-enable to fix\n';
  } else {
    msg += '❌ Disabled\n• Manual comment saving required\n';
  }
  
  msg += '\n🔄 AUTO UPDATE:\n';
  
  if (updateEnabled && updateTrigger) {
    msg += '✅ Enabled - Every Monday at 5:00 AM\n• Updates all projects\n• Includes previous week\n• Preserves comments\n';
  } else if (updateEnabled && !updateTrigger) {
    msg += '⚠️ Enabled but trigger missing\n• Disable and re-enable to fix\n';
  } else {
    msg += '❌ Disabled\n• Manual updates required\n';
  }
  
  msg += `\n⏱️ TRIGGERS:\n• Total: ${triggers.length}\n• Cache: ${triggers.filter(t => t.getHandlerFunction() === 'autoCacheAllProjects').length}\n• Update: ${triggers.filter(t => t.getHandlerFunction() === 'autoUpdateAllProjects').length}`;
  
  return msg;
}

function showChoice(title, options) {
  const ui = SpreadsheetApp.getUi();
  const numbered = options.map((opt, i) => `${i + 1} - ${opt}`).join('\n');
  const result = ui.prompt(title, `${numbered}\n\nEnter number:`, ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  const choice = parseInt(result.getResponseText());
  return (choice >= 1 && choice <= options.length) ? choice : null;
}