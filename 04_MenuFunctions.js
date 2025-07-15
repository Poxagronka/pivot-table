var MENU_PROJECTS = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Overall'];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📊 Campaign Report');
  
  menu.addItem('📈 Generate Report (Any Period)...', 'smartReportWizard')
      .addItem('🔄 Update All to Current', 'updateAllProjectsInBatches')
      .addItem('🎯 Update Selected Projects', 'updateSelectedProjectsToCurrent')
      .addSeparator()
      .addItem('⚙️ Open Settings Sheet', 'openSettingsSheet')
      .addItem('🔄 Refresh Settings', 'refreshSettingsDialog')
      .addItem('🔧 Force Update Settings', 'forceUpdateSettingsSheet')
      .addItem('📊 System Status', 'showQuickStatus')
      .addItem('🔄 Recreate Triggers', 'recreateAllTriggers')
      .addSeparator()
      .addItem('💾 Save All Comments', 'saveAllCommentsToCache')
      .addItem('🔍 Quick API Check', 'quickAPICheckAll')
      .addItem('🗑️ Clear Data...', 'clearDataWizard')
      .addItem('📱 Apps Database (TRICKY)', 'appsDbWizard')
      .addSeparator()
      .addItem('🐛 Debug Single Project', 'debugSingleProject')
      .addItem('🐙 GitHub Repository', 'openGitHubRepo')
      .addToUi();
}

function updateAllProjectsInBatches() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    ui.alert('🔐 Token Required', 'Bearer token is not configured. Please set it in Settings sheet first.', ui.ButtonSet.OK);
    return;
  }
  
  var result = ui.alert('🔄 Update All Projects', 
    'Update all projects in optimized batches?\n\n• TRICKY: Optimized processing\n• Other projects: Standard processing\n\nThis will be slower but more reliable.', 
    ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  try {
    var batch1 = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS'];
    var batch2 = ['APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    
    console.log('Starting batch 1 with TRICKY optimizations...');
    var batch1Results = updateProjectBatchOptimized(batch1, 1);
    
    if (batch1Results.successCount > 0) {
      console.log('Waiting 30 seconds before batch 2...');
      Utilities.sleep(30000);
    }
    
    console.log('Starting batch 2...');
    var batch2Results = updateProjectBatchOptimized(batch2, 2);
    
    var totalSuccess = batch1Results.successCount + batch2Results.successCount;
    var totalErrors = batch1Results.errors.concat(batch2Results.errors);
    
    try {
      console.log('Sorting project sheets...');
      Utilities.sleep(5000);
      sortProjectSheetsWithRetry();
    } catch (e) {
      console.error('Error sorting sheets:', e);
      totalErrors.push(`Sorting: ${e.toString().substring(0, 50)}...`);
    }
    
    var message = `✅ Batch update completed!\n\n• Successfully updated: ${totalSuccess}/8 projects\n• TRICKY: Optimized processing used`;
    if (totalErrors.length > 0) {
      message += `\n• Errors:\n${totalErrors.join('\n')}`;
      message += '\n\n💡 TIP: Try updating failed projects individually.';
    }
    
    ui.alert('Update Complete', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error during batch update: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function updateProjectBatchOptimized(projects, batchNumber) {
  var successCount = 0;
  var errors = [];
  
  console.log(`=== BATCH ${batchNumber} START (OPTIMIZED) ===`);
  
  projects.forEach(function(proj, index) {
    try {
      console.log(`Updating ${proj} (${index + 1}/${projects.length})...`);
      
      if (proj === 'TRICKY') {
        console.log('Using TRICKY optimized update...');
        clearTrickyCaches();
      }
      
      if (index > 0) {
        const waitTime = proj === 'TRICKY' ? 12000 : 8000;
        console.log(`Waiting ${waitTime/1000} seconds before next project...`);
        Utilities.sleep(waitTime);
      }
      
      updateProjectDataOptimized(proj);
      successCount++;
      console.log(`${proj} updated successfully`);
      
    } catch (e) {
      console.error(`Error updating ${proj}:`, e);
      var errorMsg = e.toString();
      if (errorMsg.includes('timeout') || errorMsg.includes('timed out')) {
        errors.push(`${proj}: Timeout - try individually`);
      } else {
        errors.push(`${proj}: ${errorMsg.substring(0, 50)}...`);
      }
      
      console.log('Waiting 5 seconds after error...');
      Utilities.sleep(5000);
    }
  });
  
  console.log(`=== BATCH ${batchNumber} END ===`);
  
  return { successCount: successCount, errors: errors };
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
  
  var hasTricky = selected.some(function(proj) { return proj.toLowerCase() === 'tricky'; });
  var message = `Update ${selected.length} selected projects?\n\n${selected.join(', ')}`;
  if (hasTricky) {
    message += '\n\n🚀 TRICKY will use optimized processing';
  }
  message += '\n\nThis may take several minutes.';
  
  var result = ui.alert('🔄 Update Selected Projects', message, ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  try {
    var successCount = 0;
    var errors = [];
    
    selected.forEach(function(proj, index) {
      try {
        var projectName = proj.toUpperCase();
        console.log(`Updating ${projectName} (${index + 1}/${selected.length})...`);
        
        if (projectName === 'TRICKY') {
          console.log('Clearing TRICKY caches for optimization...');
          clearTrickyCaches();
        }
        
        if (index > 0) {
          const waitTime = projectName === 'TRICKY' ? 12000 : 8000;
          console.log(`Waiting ${waitTime/1000} seconds before next project...`);
          Utilities.sleep(waitTime);
        }
        
        updateProjectDataOptimized(projectName);
        successCount++;
        console.log(`${projectName} updated successfully`);
        
      } catch (e) {
        console.error(`Error updating ${proj}:`, e);
        var errorMsg = e.toString();
        if (errorMsg.includes('timeout') || errorMsg.includes('timed out')) {
          errors.push(`${proj}: Timeout - try individually`);
        } else {
          errors.push(`${proj}: ${errorMsg.substring(0, 50)}...`);
        }
        
        console.log('Waiting 5 seconds after error...');
        Utilities.sleep(5000);
      }
    });
    
    if (successCount > 0) {
      try {
        console.log('Sorting project sheets...');
        Utilities.sleep(3000);
        sortProjectSheetsWithRetry();
      } catch (e) {
        console.error('Error sorting sheets:', e);
        errors.push(`Sorting: ${e.toString().substring(0, 30)}...`);
      }
    }
    
    var message = `✅ Update completed!\n\n• Successfully updated: ${successCount}/${selected.length} projects`;
    if (hasTricky) {
      message += '\n• TRICKY: Optimized processing used';
    }
    if (errors.length > 0) {
      message += `\n• Errors:\n${errors.join('\n')}`;
      message += '\n\n💡 TIP: Try updating problematic projects individually.';
    }
    
    ui.alert('Update Complete', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error during update: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function updateSingleProject() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    ui.alert('🔐 Token Required', 'Bearer token is not configured. Please set it in Settings sheet first.', ui.ButtonSet.OK);
    return;
  }
  
  var project = showChoice('Select Project to Update:', MENU_PROJECTS);
  if (!project) return;
  
  var projectName = MENU_PROJECTS[project-1].toUpperCase();
  
  try {
    console.log(`Updating single project: ${projectName}`);
    
    if (projectName === 'TRICKY') {
      console.log('Using TRICKY optimized processing...');
      clearTrickyCaches();
    }
    
    updateProjectDataOptimized(projectName);
    
    var successMessage = `${projectName} updated successfully!`;
    if (projectName === 'TRICKY') {
      successMessage += '\n\n🚀 Used optimized processing for faster performance';
    }
    
    ui.alert('✅ Success', successMessage, ui.ButtonSet.OK);
  } catch (e) {
    console.error(`Error updating ${projectName}:`, e);
    var errorMsg = e.toString();
    if (errorMsg.includes('timeout') || errorMsg.includes('timed out')) {
      ui.alert('⏱️ Timeout', `${projectName} update timed out. Try reducing date range or check API status.`, ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Error', `Error updating ${projectName}:\n\n${errorMsg}`, ui.ButtonSet.OK);
    }
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
    message += `🚀 TRICKY Optimizations: Available\n`;
    
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
  message += `🎯 Metrics: Unified (eROAS D730)\n`;
  message += `🚀 TRICKY Optimizations: ✅ Active\n\n`;
  
  var triggers = ScriptApp.getProjectTriggers();
  var cacheTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoCacheAllProjects'; });
  var updateTriggers = getUpdateTriggers();
  
  var cacheEnabled = isAutoCacheEnabled();
  var updateEnabled = isAutoUpdateEnabled();
  
  var syncIssues = [];
  if (cacheEnabled && !cacheTrigger) {
    syncIssues.push('• Cache trigger missing (will auto-create)');
  }
  if (!cacheEnabled && cacheTrigger) {
    syncIssues.push('• Cache trigger exists but disabled (will remove)');
  }
  if (updateEnabled && updateTriggers.length !== 8) {
    syncIssues.push('• Update triggers incomplete (will auto-create)');
  }
  if (!updateEnabled && updateTriggers.length > 0) {
    syncIssues.push('• Update triggers exist but disabled (will remove)');
  }
  
  if (syncIssues.length > 0) {
    message += '⚠️ SYNC ISSUES:\n' + syncIssues.join('\n') + '\n\n';
    message += 'Use "🔄 Refresh Settings" to fix.\n\n';
  } else {
    message += '✅ All triggers synchronized\n\n';
  }
  
  message += '📅 AUTOMATION SCHEDULE:\n';
  message += '• Auto Cache: Daily at 2:00 AM\n';
  message += '• Auto Update: Exact times:\n';
  message += '  - TRICKY: 5:00 AM (optimized)\n';
  message += '  - MOLOCO: 5:00 AM\n';
  message += '  - REGULAR: 5:00 AM\n';
  message += '  - GOOGLE_ADS: 5:00 AM\n';
  message += '  - APPLOVIN: 5:00 AM\n';
  message += '  - MINTEGRAL: 5:00 AM\n';
  message += '  - INCENT: 6:00 AM\n';
  message += '  - OVERALL: 6:00 AM\n\n';
  
  message += '💡 TIP: Use "📈 Generate Report" for flexible periods (any number of days)\n';
  message += '🔧 Use "🔄 Update All" for batch processing or update projects individually\n';
  message += '🚀 TRICKY uses optimized processing for better performance\n';
  message += '⚠️ Large periods (>180 days) may cause timeouts - use date ranges instead';
  
  ui.alert('System Status', message, ui.ButtonSet.OK);
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
      
      if (proj === 'TRICKY') {
        console.log('Clearing TRICKY caches for API check...');
        clearTrickyCaches();
      }
      
      var dateRange = getDateRange(7);
      var raw = fetchCampaignData(dateRange);
      
      if (!raw.data?.analytics?.richStats?.stats?.length) {
        results += `❌ ${proj}: No data\n`;
      } else {
        var count = raw.data.analytics.richStats.stats.length;
        var optimizedNote = proj === 'TRICKY' ? ' (optimized)' : '';
        results += `✅ ${proj}: ${count} records${optimizedNote}\n`;
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
  
  var period = showChoice('📅 Select Period - Step 2/3', ['Last 30 days', 'Last 60 days', 'Last 90 days', 'Custom days (any number)', 'Date range (specific dates)']);
  if (!period) return;
  
  var days = [30, 60, 90];
  
  if (scope === 1) {
    if (period <= 3) {
      quickGenerateAllForDaysOptimized(days[period-1]);
    } else if (period === 4) {
      var customDays = promptCustomDays();
      if (customDays) quickGenerateAllForDaysOptimized(customDays);
    } else {
      var dates = promptDateRange();
      if (dates) runAllProjectsDateRangeOptimized(dates.start, dates.end);
    }
  } else if (scope === 2) {
    var project = showChoice('Select Project - Step 3/3', MENU_PROJECTS);
    if (!project) return;
    var projectName = MENU_PROJECTS[project-1].toUpperCase();
    
    if (period <= 3) {
      generateProjectReportOptimized(projectName, days[period-1]);
    } else if (period === 4) {
      var customDays = promptCustomDays();
      if (customDays) generateProjectReportOptimized(projectName, customDays);
    } else {
      var dates = promptDateRange();
      if (dates) generateProjectReportForDateRangeOptimized(projectName, dates.start, dates.end);
    }
  } else {
    var selected = showMultiChoice('Select Projects:', MENU_PROJECTS);
    if (!selected || selected.length === 0) return;
    
    if (period <= 3) {
      runSelectedProjectsOptimized(selected, days[period-1]);
    } else if (period === 4) {
      var customDays = promptCustomDays();
      if (customDays) runSelectedProjectsOptimized(selected, customDays);
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
  if (ui.alert('Confirm Clear All', 'Clear data from ALL projects? Comments preserved.\n\n🚀 TRICKY will use optimized clearing.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    var successCount = 0;
    
    projects.forEach(function(proj) {
      try {
        if (proj === 'TRICKY') {
          console.log('Clearing TRICKY with optimizations...');
          clearTrickyCaches();
        }
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

function clearProjectAllDataOptimized(projectName) {
  var ui = SpreadsheetApp.getUi();
  
  var message = `Clear all ${projectName} data? Comments preserved.`;
  if (projectName === 'TRICKY') {
    message += '\n\n🚀 Will use optimized clearing process.';
  }
  
  if (ui.alert(`Clear ${projectName} Data`, message, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    if (projectName === 'TRICKY') {
      console.log('Clearing TRICKY with optimizations...');
      clearTrickyCaches();
    }
    clearProjectDataSilent(projectName);
    ui.alert('Success', `${projectName} data cleared. Comments preserved.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Error clearing ${projectName}: ${e.toString()}`, ui.ButtonSet.OK);
  }
}

function debugSingleProject() {
  var p = showChoice('Select Project to Debug:', MENU_PROJECTS);
  if (p) debugProjectReportGenerationOptimized(MENU_PROJECTS[p-1].toUpperCase());
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
      console.log('Created update triggers with TRICKY optimizations');
    } else if (!settings.automation.autoUpdate && updateTriggers.length > 0) {
      clearAllUpdateTriggers();
      console.log('Deleted all update triggers');
    }
    
    console.log('Triggers synchronized with Settings sheet');
  } catch (e) {
    console.error('Error syncing triggers with settings:', e);
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
    ScriptApp.newTrigger(item.func)
      .timeBased()
      .everyDays(1)
      .atHour(item.hour)
      .nearMinute(item.minute)
      .create();
  });
}

function updateProjectDataWithRetry(projectName, maxRetries = 2) {
  var baseDelay = 3000;
  
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      if (projectName === 'TRICKY') {
        console.log('Using TRICKY optimized update with retry...');
        clearTrickyCaches();
      }
      updateProjectDataOptimized(projectName);
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

function promptCustomDays() {
  var ui = SpreadsheetApp.getUi();
  var message = 'Enter number of days for report:\n\n';
  message += '📅 Common periods:\n';
  message += '• 7 days - Last week\n';
  message += '• 14 days - Last 2 weeks\n';
  message += '• 30 days - Last month\n';
  message += '• 90 days - Last quarter\n';
  message += '• 180 days - Last 6 months\n';
  message += '• 365 days - Last year\n\n';
  message += '⚠️ Note: Large periods (>180 days) may cause timeouts\n';
  message += '💡 Tip: Use "Date range" for specific periods or large datasets\n';
  message += '🚀 TRICKY uses optimized processing for better performance';
  
  var result = ui.prompt('Custom Days', message, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  
  var input = result.getResponseText().trim();
  var days = parseInt(input);
  
  if (isNaN(days) || days < 1) {
    ui.alert('❌ Invalid Input', 'Please enter a valid number of days (minimum 1).', ui.ButtonSet.OK);
    return null;
  }
  
  if (days > 730) {
    ui.alert('❌ Period Too Large', 
      `${days} days (over 2 years) is too large and will likely cause timeouts.\n\nMaximum recommended: 365 days\n\nFor large historical data, use "Date range" option with smaller chunks.`, 
      ui.ButtonSet.OK);
    return null;
  }
  
  if (days > 365) {
    var confirm = ui.alert('⚠️ Large Period Warning', 
      `You entered ${days} days (over 1 year).\n\nThis may cause timeouts or performance issues.\n\nRecommendation: Use "Date range" for better control.\n\nContinue anyway?`, 
      ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return null;
  } else if (days > 180) {
    var confirm = ui.alert('⚠️ Performance Warning', 
      `You entered ${days} days (${Math.round(days/30)} months).\n\nThis may take longer to process and could timeout.\n\nContinue?`, 
      ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return null;
  }
  
  return days;
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

function quickGenerateAllForDaysOptimized(days) {
  var ui = SpreadsheetApp.getUi();
  var success = 0;
  
  try {
    for (var i = 0; i < MENU_PROJECTS.length; i++) {
      var p = MENU_PROJECTS[i];
      try { 
        generateProjectReportOptimized(p.toUpperCase(), days); 
        success++; 
      } catch(e) { 
        console.error(e); 
      }
    }
    sortProjectSheets();
    ui.alert('✅ Complete', 'Generated ' + success + '/' + MENU_PROJECTS.length + ' reports (TRICKY optimized)', ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Error', e.toString(), ui.ButtonSet.OK);
  }
}

function runSelectedProjectsOptimized(projects, days) {
  for (var i = 0; i < projects.length; i++) {
    generateProjectReportOptimized(projects[i].toUpperCase(), days);
  }
  sortProjectSheets();
  SpreadsheetApp.getUi().alert('✅ Complete', 'Generated ' + projects.length + ' reports (TRICKY optimized)', SpreadsheetApp.getUi().ButtonSet.OK);
}

function runAllProjectsDateRangeOptimized(start, end) {
  for (var i = 0; i < MENU_PROJECTS.length; i++) {
    generateProjectReportForDateRangeOptimized(MENU_PROJECTS[i].toUpperCase(), start, end);
  }
  sortProjectSheets();
  SpreadsheetApp.getUi().alert('✅ Complete', 'All reports generated (TRICKY optimized)', SpreadsheetApp.getUi().ButtonSet.OK);
}

function runSelectedProjectsDateRangeOptimized(projects, start, end) {
  for (var i = 0; i < projects.length; i++) {
    generateProjectReportForDateRangeOptimized(projects[i].toUpperCase(), start, end);
  }
  sortProjectSheets();
  SpreadsheetApp.getUi().alert('✅ Complete', 'Generated ' + projects.length + ' reports (TRICKY optimized)', SpreadsheetApp.getUi().ButtonSet.OK);
}

function generateProjectReportOptimized(projectName, days) { 
  console.log(`Generating ${projectName} report for ${days} days`);
  if (days > 180) {
    console.log(`Warning: Large period (${days} days) - monitoring for timeouts`);
  }
  
  if (projectName === 'TRICKY') {
    console.log('Using TRICKY optimized processing...');
    clearTrickyCaches();
  }
  
  setCurrentProject(projectName); 
  generateReport(days); 
}

function generateProjectReportForDateRangeOptimized(projectName, startDate, endDate) { 
  if (projectName === 'TRICKY') {
    console.log('Using TRICKY optimized processing for date range...');
    clearTrickyCaches();
  }
  setCurrentProject(projectName); 
  generateReportForDateRange(startDate, endDate); 
}

function debugProjectReportGenerationOptimized(projectName) { 
  if (projectName === 'TRICKY') {
    console.log('Using TRICKY optimized debug...');
    clearTrickyCaches();
  }
  setCurrentProject(projectName); 
  debugReportGeneration(); 
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
  
  var action = showChoice('📱 Apps Database Management (Optimized)', [
    'View Cache Status',
    'Refresh Apps Database', 
    'View Sample Data',
    'Clear Cache',
    'Debug Update Process',
    'Clear Optimization Caches'
  ]);
  if (!action) return;
  
  switch(action) {
    case 1: showAppsDbStatusOptimized(); break;
    case 2: refreshAppsDatabase(); break;
    case 3: showAppsDbSample(); break;
    case 4: clearAppsDbCache(); break;
    case 5: debugAppsDatabase(); break;
    case 6: clearTrickyOptimizationCaches(); break;
  }
}

function showAppsDbStatusOptimized() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var appsDb = new AppsDatabase('TRICKY');
    var cache = appsDb.loadFromCache();
    var appCount = Object.keys(cache).length;
    
    var message = '📱 APPS DATABASE STATUS (OPTIMIZED)\n\n';
    message += '• Total Apps: ' + appCount + '\n';
    
    if (appCount > 0) {
      var bundleIds = Object.keys(cache);
      var sampleApp = cache[bundleIds[0]];
      message += '• Last Updated: ' + (sampleApp.lastUpdated || 'Unknown') + '\n';
      message += '• Cache Sheet: ' + (appsDb.cacheSheet ? 'Found' : 'Missing') + '\n';
      
      var shouldUpdate = appsDb.shouldUpdateCache();
      message += '• Update Needed: ' + (shouldUpdate ? 'YES (>24h old)' : 'NO') + '\n';
      message += '• Optimization Cache: ' + (appsDb.optimizedCache ? 'Active' : 'Not Loaded') + '\n\n';
      
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
    
    var message = '📱 APPS DATABASE SAMPLE (OPTIMIZED)\n\n';
    var sampleCount = Math.min(5, bundleIds.length);
    
    for (var i = 0; i < sampleCount; i++) {
      var bundleId = bundleIds[i];
      var app = cache[bundleId];
      message += bundleId + '\n  → ' + app.publisher + ' ' + app.appName + '\n\n';
    }
    
    if (bundleIds.length > sampleCount) {
      message += '... and ' + (bundleIds.length - sampleCount) + ' more apps\n\n';
    }
    
    message += '🚀 Optimization cache: ' + (appsDb.optimizedCache ? 'Active' : 'Will load on demand');
    
    ui.alert('Apps Database Sample', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error showing sample data: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function clearAppsDbCache() {
  var ui = SpreadsheetApp.getUi();
  
  if (ui.alert('Clear Apps Database Cache', 'Clear cached app data? Will rebuild on next refresh.\n\n🚀 This will also clear optimization caches.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    var appsDb = new AppsDatabase('TRICKY');
    if (appsDb.cacheSheet && appsDb.cacheSheet.getLastRow() > 1) {
      appsDb.cacheSheet.deleteRows(2, appsDb.cacheSheet.getLastRow() - 1);
      clearTrickyCaches();
      ui.alert('Success', 'Apps Database cache and optimization caches cleared.', ui.ButtonSet.OK);
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
  
  var htmlOutput = HtmlService.createHtmlOutput(
    '<script>window.open("' + githubUrl + '", "_blank"); google.script.host.close();</script>'
  ).setWidth(400).setHeight(300);
  
  ui.showModalDialog(htmlOutput, 'Opening GitHub Repository...');
}

function recreateAllTriggers() {
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.alert('🔄 Recreate Triggers', 
    'Recreate all automation triggers?\n\n⏰ New schedule:\n• Cache: 2:00 AM\n• Updates: 5:00-6:10 AM (10min apart)\n🚀 TRICKY: Optimized processing', 
    ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  try {
    console.log('Recreating all triggers with TRICKY optimizations...');
    
    clearAllUpdateTriggers();
    
    var cacheEnabled = isAutoCacheEnabled();
    var updateEnabled = isAutoUpdateEnabled();
    
    if (cacheEnabled) {
      ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
      console.log('Cache trigger recreated');
    }
    
    if (updateEnabled) {
      createUpdateTriggers();
      console.log('Update triggers recreated with TRICKY optimizations');
    }
    
    ui.alert('✅ Triggers Recreated', 'All triggers recreated successfully!\n\n🚀 TRICKY optimizations: Active', ui.ButtonSet.OK);
    
  } catch (e) {
    console.error('Error recreating triggers:', e);
    ui.alert('❌ Error', 'Error recreating triggers: ' + e.toString(), ui.ButtonSet.OK);
  }
}