/**
 * Menu Functions - ОБНОВЛЕНО: медленный режим по умолчанию для множественных обновлений
 */

var MENU_PROJECTS = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Incent_traffic', 'Overall'];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📊 Campaign Report');
  
  menu.addItem('📈 Generate Report...', 'smartReportWizard')
      .addItem('🔄 Update All Projects (Safe Mode)', 'updateAllProjectsToCurrent')
      .addItem('🎯 Update Selected Projects (Safe Mode)', 'updateSelectedProjectsToCurrent')
      .addItem('🚀 Quick Update Single Project', 'updateSingleProjectQuick')
      .addSeparator()
      .addSubMenu(ui.createMenu('⚙️ Settings')
        .addItem('📄 Open Settings Sheet', 'openSettingsSheet')
        .addItem('🔄 Refresh Settings', 'refreshSettingsDialog')
        .addItem('🔧 Force Update Settings', 'forceUpdateSettingsSheet')
        .addItem('📊 System Status', 'showQuickStatus')
        .addSeparator()
        .addItem('🧹 Clear Column Cache', 'clearColumnCacheDialog')
        .addItem('💾 Save All Comments', 'saveAllCommentsToCache')
        .addItem('🗑️ Clear Data...', 'clearDataWizard')
        .addSeparator()
        .addItem('🔍 Quick API Check', 'quickAPICheckAll')
        .addItem('📱 Apps Database (TRICKY)', 'appsDbWizard')
        .addItem('🐛 Debug Single Project', 'debugSingleProject'))
      .addSeparator()
      .addItem('🐙 GitHub Repository', 'openGitHubRepo')
      .addToUi();
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
  
  var estimatedMinutes = Math.ceil(selected.length * 1.5);
  var result = ui.alert('🔄 Update Selected Projects', 
    `Update ${selected.length} selected projects?\n\n${selected.join(', ')}\n\n` +
    `⏱️ Estimated time: ${estimatedMinutes} minutes\n` +
    `(Using safe mode with extended pauses to prevent timeouts)`, 
    ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  // Предварительная загрузка настроек
  try {
    ui.alert('Preparing...', 'Loading settings and preparing for update...', ui.ButtonSet.OK);
    preloadSettings();
    Utilities.sleep(2000);
  } catch (e) {
    console.error('Error preloading settings:', e);
  }
  
  var successfulProjects = [];
  var failedProjects = [];
  
  selected.forEach(function(proj, index) {
    try {
      var projectName = proj.toUpperCase();
      
      // Показываем прогресс для каждого проекта
      console.log(`\n=== UPDATING ${projectName} (${index + 1}/${selected.length}) ===`);
      console.log(`Completed so far: ${successfulProjects.join(', ') || 'None'}`);
      
      // Очистка кешей перед каждым проектом (кроме первого)
      if (index > 0) {
        console.log('Clearing caches and waiting before project update...');
        clearSettingsCache();
        clearAllCommentColumnCaches();
        SpreadsheetApp.flush();
        
        // ОСНОВНАЯ ПАУЗА: 20 секунд между проектами
        console.log('Waiting 20 seconds before next project...');
        Utilities.sleep(20000);
      }
      
      // Дополнительная пауза каждые 3 проекта
      if (index > 0 && index % 3 === 0) {
        console.log('Extended cooldown after 3 projects (30 seconds)...');
        Utilities.sleep(30000);
      }
      
      // Обновление проекта
      updateProjectDataWithRetry(projectName);
      
      successfulProjects.push(projectName);
      console.log(`✅ ${projectName} updated successfully`);
      
      // Пауза после успешного обновления
      Utilities.sleep(3000);
      
    } catch (e) {
      console.error(`❌ Failed to update ${proj}:`, e);
      failedProjects.push({
        project: proj,
        error: e.toString().substring(0, 80)
      });
      
      // Увеличенная пауза после ошибки (30 секунд)
      console.log('Error occurred. Waiting 30 seconds before continuing...');
      Utilities.sleep(30000);
    }
  });
  
  // Сортировка листов
  if (successfulProjects.length > 0) {
    try {
      console.log('Waiting before sorting sheets...');
      Utilities.sleep(5000);
      sortProjectSheetsWithRetry();
    } catch (e) {
      console.error('Error sorting sheets:', e);
    }
  }
  
  // Финальный отчет
  var message = `✅ Update completed!\n\n`;
  
  if (successfulProjects.length > 0) {
    message += `• Successfully updated: ${successfulProjects.length}/${selected.length} projects\n`;
    message += `  ${successfulProjects.join(', ')}\n\n`;
  }
  
  if (failedProjects.length > 0) {
    message += `• Failed projects:\n`;
    failedProjects.forEach(function(fail) {
      message += `  ${fail.project}: ${fail.error}...\n`;
    });
    message += '\n💡 TIP: Try updating failed projects individually.';
  }
  
  ui.alert('Update Complete', message, ui.ButtonSet.OK);
}

function updateAllProjectsToCurrent() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    ui.alert('🔐 Token Required', 'Bearer token is not configured. Please set it in Settings sheet first.', ui.ButtonSet.OK);
    return;
  }
  
  var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
  var estimatedMinutes = Math.ceil(projects.length * 1.5);
  
  var result = ui.alert('🔄 Update All Projects', 
    `This will update all ${projects.length} projects with the latest data.\n\n` +
    `⏱️ Estimated time: ${estimatedMinutes} minutes\n` +
    `(Using safe mode with extended pauses to prevent timeouts)\n\n` +
    `Continue?`, 
    ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  // Предварительная загрузка настроек
  try {
    ui.alert('Preparing...', 'Loading settings and preparing for batch update...', ui.ButtonSet.OK);
    preloadSettings();
    Utilities.sleep(2000);
  } catch (e) {
    console.error('Error preloading settings:', e);
  }
  
  var successfulProjects = [];
  var failedProjects = [];
  
  projects.forEach(function(proj, index) {
    try {
      console.log(`\n=== UPDATING ${proj} (${index + 1}/${projects.length}) ===`);
      console.log(`Completed so far: ${successfulProjects.join(', ') || 'None'}`);
      
      // Очистка кешей перед каждым проектом (кроме первого)
      if (index > 0) {
        console.log('Clearing caches and waiting before project update...');
        clearSettingsCache();
        clearAllCommentColumnCaches();
        SpreadsheetApp.flush();
        
        // ОСНОВНАЯ ПАУЗА: 20 секунд между проектами
        console.log('Waiting 20 seconds before next project...');
        Utilities.sleep(20000);
      }
      
      // Дополнительная пауза каждые 3 проекта
      if (index > 0 && index % 3 === 0) {
        console.log('Extended cooldown after 3 projects (30 seconds)...');
        Utilities.sleep(30000);
        
        // Показываем промежуточный прогресс
        if (index < projects.length - 1) {
          ui.alert('Progress Update', 
            `Completed: ${index} of ${projects.length} projects\n\n` +
            `Next: ${projects.slice(index, Math.min(index + 3, projects.length)).join(', ')}\n\n` +
            `Please wait...`, 
            ui.ButtonSet.OK);
        }
      }
      
      // Обновление проекта
      updateProjectDataWithRetry(proj);
      
      successfulProjects.push(proj);
      console.log(`✅ ${proj} updated successfully`);
      
      // Пауза после успешного обновления
      Utilities.sleep(3000);
      
    } catch (e) {
      console.error(`❌ Failed to update ${proj}:`, e);
      failedProjects.push({
        project: proj,
        error: e.toString().substring(0, 80)
      });
      
      // Увеличенная пауза после ошибки (30 секунд)
      console.log('Error occurred. Waiting 30 seconds before continuing...');
      Utilities.sleep(30000);
    }
  });
  
  // Сортировка листов
  if (successfulProjects.length > 0) {
    try {
      console.log('Waiting before sorting sheets...');
      Utilities.sleep(5000);
      sortProjectSheetsWithRetry();
      console.log('Sheets sorted successfully');
    } catch (e) {
      console.error('Error sorting sheets:', e);
    }
  }
  
  // Финальный отчет
  var message = `✅ Update completed!\n\n`;
  
  if (successfulProjects.length > 0) {
    message += `• Successfully updated: ${successfulProjects.length}/${projects.length} projects\n`;
  }
  
  if (failedProjects.length > 0) {
    message += `\n• Failed projects:\n`;
    failedProjects.forEach(function(fail) {
      message += `  ${fail.project}: ${fail.error}...\n`;
    });
    message += '\n💡 TIP: Try updating failed projects individually.';
  }
  
  ui.alert('Update Complete', message, ui.ButtonSet.OK);
}

// Функция быстрого обновления одного проекта (без больших пауз)
function updateSingleProjectQuick() {
  var ui = SpreadsheetApp.getUi();
  
  if (!isBearerTokenConfigured()) {
    ui.alert('🔐 Token Required', 'Bearer token is not configured. Please set it in Settings sheet first.', ui.ButtonSet.OK);
    return;
  }
  
  var projects = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Overall'];
  var choice = showChoice('Select Project to Update (Quick Mode):', projects);
  
  if (!choice) return;
  
  var projectName = projects[choice - 1].toUpperCase();
  
  var result = ui.alert('🚀 Quick Update Single Project', 
    `Update ${projectName} project?\n\n` +
    `This will use minimal pauses (faster but may timeout if system is busy).`, 
    ui.ButtonSet.YES_NO);
  
  if (result !== ui.Button.YES) return;
  
  try {
    ui.alert('Processing...', `Updating ${projectName}...`, ui.ButtonSet.OK);
    
    // Минимальная пауза для одного проекта
    updateProjectDataWithRetry(projectName);
    
    // Сортировка листов
    Utilities.sleep(2000);
    sortProjectSheetsWithRetry();
    
    ui.alert('Success', `✅ ${projectName} updated successfully!`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `❌ Failed to update ${projectName}:\n\n${e.toString()}`, ui.ButtonSet.OK);
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
  var updateTrigger = triggers.find(function(t) { return t.getHandlerFunction() === 'autoUpdateAllProjects'; });
  
  var cacheEnabled = isAutoCacheEnabled();
  var updateEnabled = isAutoUpdateEnabled();
  
  var syncIssues = [];
  if (cacheEnabled && !cacheTrigger) {
    syncIssues.push('• Cache trigger missing (will auto-create)');
  }
  if (!cacheEnabled && cacheTrigger) {
    syncIssues.push('• Cache trigger exists but disabled (will remove)');
  }
  if (updateEnabled && !updateTrigger) {
    syncIssues.push('• Update trigger missing (will auto-create)');
  }
  if (!updateEnabled && updateTrigger) {
    syncIssues.push('• Update trigger exists but disabled (will remove)');
  }
  
  if (syncIssues.length > 0) {
    message += '⚠️ SYNC ISSUES:\n' + syncIssues.join('\n') + '\n\n';
    message += 'Use "🔄 Refresh Settings" to fix.\n\n';
  } else {
    message += '✅ All triggers synchronized\n\n';
  }
  
  message += '📅 AUTOMATION SCHEDULE:\n';
  message += '• Auto Cache: Every hour\n';
  message += '• Auto Update: Daily at 5:00 AM\n';
  message += '• Previous week data: Included starting from Tuesday\n\n';
  
  message += '💡 TIP: Use Settings sheet to configure all options\n';
  message += '🔧 Use "Force Update Settings" if you have old targets';
  
  ui.alert('System Status', message, ui.ButtonSet.OK);
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
  
  var period = showChoice('📅 Select Period - Step 2/3', ['Last 30 days', 'Last 60 days', 'Last 90 days', 'Custom days', 'Date range']);
  if (!period) return;
  
  var days = [30, 60, 90];
  
  if (scope === 1) {
    if (period <= 3) {
      quickGenerateAllForDays(days[period-1]);
    } else if (period === 4) {
      var customDays = promptNumber('Enter days:', [120, 360]);
      if (customDays) quickGenerateAllForDays(customDays);
    } else {
      var dates = promptDateRange();
      if (dates) runAllProjectsDateRange(dates.start, dates.end);
    }
  } else if (scope === 2) {
    var project = showChoice('Select Project - Step 3/3', MENU_PROJECTS);
    if (!project) return;
    var projectName = MENU_PROJECTS[project-1].toUpperCase();
    
    if (period <= 3) {
      generateProjectReport(projectName, days[period-1]);
    } else if (period === 4) {
      var customDays = promptNumber('Enter days:', [120, 360]);
      if (customDays) generateProjectReport(projectName, customDays);
    } else {
      var dates = promptDateRange();
      if (dates) generateProjectReportForDateRange(projectName, dates.start, dates.end);
    }
  } else {
    var selected = showMultiChoice('Select Projects:', MENU_PROJECTS);
    if (!selected || selected.length === 0) return;
    
    if (period <= 3) {
      runSelectedProjects(selected, days[period-1]);
    } else if (period === 4) {
      var customDays = promptNumber('Enter days:', [120, 360]);
      if (customDays) runSelectedProjects(selected, customDays);
    } else {
      var dates = promptDateRange();
      if (dates) runSelectedProjectsDateRange(selected, dates.start, dates.end);
    }
  }
}

function clearDataWizard() {
  var choice = showChoice('🗑️ Clear Data', ['Clear All Projects', 'Clear Single Project']);
  if (!choice) return;
  
  if (choice === 1) {
    clearAllProjectsData();
  } else {
    var p = showChoice('Select Project:', MENU_PROJECTS);
    if (p) clearProjectAllData(MENU_PROJECTS[p-1].toUpperCase());
  }
}

function clearAllProjectsData() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Confirm Clear All', 'Clear data from ALL projects? Comments preserved.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    var projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
    var successCount = 0;
    
    projects.forEach(function(proj) {
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
  var ui = SpreadsheetApp.getUi();
  if (ui.alert(`Clear ${projectName} Data`, `Clear all ${projectName} data? Comments preserved.`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    clearProjectDataSilent(projectName);
    ui.alert('Success', `${projectName} data cleared. Comments preserved.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Error clearing ${projectName}: ${e.toString()}`, ui.ButtonSet.OK);
  }
}

function debugSingleProject() {
  var p = showChoice('Select Project to Debug:', MENU_PROJECTS);
  if (p) debugProjectReportGeneration(MENU_PROJECTS[p-1].toUpperCase());
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
      ScriptApp.newTrigger('autoUpdateAllProjects').timeBased().atHour(5).everyDays(1).create();
      console.log('Created auto update trigger');
    } else if (!settings.automation.autoUpdate && updateTrigger) {
      ScriptApp.deleteTrigger(updateTrigger);
      console.log('Deleted auto update trigger');
    }
    
    console.log('Triggers synchronized with Settings sheet');
  } catch (e) {
    console.error('Error syncing triggers with settings:', e);
    throw e;
  }
}

function updateProjectDataWithRetry(projectName, maxRetries = 3) {
  var baseDelay = 5000; // Увеличена базовая задержка до 5 секунд
  
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      // Очистка кешей перед попыткой
      clearSettingsCache();
      SpreadsheetApp.flush();
      Utilities.sleep(1000);
      
      updateProjectData(projectName);
      return;
    } catch (e) {
      console.error(`${projectName} update attempt ${attempt} failed:`, e);
      
      // Специальная обработка таймаутов
      if (e.toString().includes('timed out') || e.toString().includes('Service Spreadsheets')) {
        console.log('Timeout detected - waiting longer before retry...');
        
        // Очистка всех кешей
        clearSettingsCache();
        clearAllCommentColumnCaches();
        
        // Экспоненциальная задержка для таймаутов
        var timeoutDelay = baseDelay * Math.pow(2, attempt);
        console.log(`Waiting ${timeoutDelay}ms before retry...`);
        Utilities.sleep(timeoutDelay);
        
        // Дополнительная пауза для восстановления соединения
        SpreadsheetApp.flush();
        Utilities.sleep(2000);
      } else {
        // Обычная экспоненциальная задержка для других ошибок
        var delay = baseDelay * Math.pow(1.5, attempt - 1);
        console.log(`Waiting ${delay}ms before retry...`);
        Utilities.sleep(delay);
      }
      
      if (attempt === maxRetries) {
        throw e;
      }
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

function quickGenerateAllForDays(days) {
  var ui = SpreadsheetApp.getUi();
  var success = 0;
  
  try {
    for (var i = 0; i < MENU_PROJECTS.length; i++) {
      var p = MENU_PROJECTS[i];
      try { 
        generateProjectReport(p.toUpperCase(), days); 
        success++; 
      } catch(e) { 
        console.error(e); 
      }
    }
    sortProjectSheets();
    ui.alert('✅ Complete', 'Generated ' + success + '/' + MENU_PROJECTS.length + ' reports', ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Error', e.toString(), ui.ButtonSet.OK);
  }
}

function runSelectedProjects(projects, days) {
  for (var i = 0; i < projects.length; i++) {
    generateProjectReport(projects[i].toUpperCase(), days);
  }
  sortProjectSheets();
  SpreadsheetApp.getUi().alert('✅ Complete', 'Generated ' + projects.length + ' reports', SpreadsheetApp.getUi().ButtonSet.OK);
}

function runAllProjectsDateRange(start, end) {
  for (var i = 0; i < MENU_PROJECTS.length; i++) {
    generateProjectReportForDateRange(MENU_PROJECTS[i].toUpperCase(), start, end);
  }
  sortProjectSheets();
  SpreadsheetApp.getUi().alert('✅ Complete', 'All reports generated', SpreadsheetApp.getUi().ButtonSet.OK);
}

function runSelectedProjectsDateRange(projects, start, end) {
  for (var i = 0; i < projects.length; i++) {
    generateProjectReportForDateRange(projects[i].toUpperCase(), start, end);
  }
  sortProjectSheets();
  SpreadsheetApp.getUi().alert('✅ Complete', 'Generated ' + projects.length + ' reports', SpreadsheetApp.getUi().ButtonSet.OK);
}

function generateProjectReport(projectName, days) { setCurrentProject(projectName); generateReport(days); }
function generateProjectReportForDateRange(projectName, startDate, endDate) { setCurrentProject(projectName); generateReportForDateRange(startDate, endDate); }
function debugProjectReportGeneration(projectName) { setCurrentProject(projectName); debugReportGeneration(); }

function appsDbWizard() {
  var ui = SpreadsheetApp.getUi();
  
  if (CURRENT_PROJECT !== 'TRICKY') {
    var switchResult = ui.alert('Apps Database - TRICKY Only', 
      'Apps Database is only used for TRICKY project.\n\nSwitch to TRICKY project now?', 
      ui.ButtonSet.YES_NO);
    
    if (switchResult !== ui.Button.YES) return;
    setCurrentProject('TRICKY');
  }
  
  var action = showChoice('📱 Apps Database Management', [
    'View Cache Status',
    'Refresh Apps Database', 
    'View Sample Data',
    'Clear Cache',
    'Debug Update Process'
  ]);
  if (!action) return;
  
  switch(action) {
    case 1: showAppsDbStatus(); break;
    case 2: refreshAppsDatabase(); break;
    case 3: showAppsDbSample(); break;
    case 4: clearAppsDbCache(); break;
    case 5: debugAppsDatabase(); break;
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
      ui.alert('Success', 'Apps Database cache cleared.', ui.ButtonSet.OK);
    } else {
      ui.alert('No Cache', 'Apps Database cache sheet not found.', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', 'Error clearing cache: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function clearColumnCacheDialog() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('🧹 Clear Column Cache', 
    'Clear cached column positions for all projects?\n\nThis will force re-detection of column positions on next use.', 
    ui.ButtonSet.YES_NO);
  
  if (result === ui.Button.YES) {
    clearAllCommentColumnCaches();
    ui.alert('✅ Success', 'Column cache cleared for all projects.', ui.ButtonSet.OK);
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