var MENU_PROJECTS = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Incent_traffic', 'Overall'];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📊 Campaign Report');
  
  menu.addItem('📈 Generate Report...', 'smartReportWizard')
      .addItem('🔄 Update All Projects', 'updateAllProjects')
      .addItem('🎯 Update Selected Projects', 'updateSelectedProjects')
      .addSeparator()
      .addSubMenu(ui.createMenu('⚙️ Settings')
        .addItem('📄 Open Settings Sheet', 'openSettingsSheet')
        .addItem('🔄 Refresh Settings', 'refreshSettings')
        .addItem('📊 System Status', 'showQuickStatus')
        .addSeparator()
        .addItem('🧹 Clear Column Cache', 'clearColumnCache')
        .addItem('💾 Save All Comments', 'saveAllCommentsToCache')
        .addSeparator()
        .addItem('🔍 Quick API Check', 'quickAPICheckAll')
        .addItem('📱 Apps Database (TRICKY)', 'appsDbWizard')
        .addItem('🐛 Debug Single Project', 'debugSingleProject'))
      .addSeparator()
      .addItem('🐙 GitHub Repository', 'openGitHubRepo')
      .addToUi();
}

function updateSelectedProjects() {
  if (!isBearerTokenConfigured()) {
    openSettingsSheet();
    return;
  }
  
  var projects = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Incent_Traffic', 'Overall'];
  var selected = showMultiChoice('Select Projects to Update:', projects);
  
  if (!selected || selected.length === 0) {
    return;
  }
  
  try {
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
      
      console.log(`\n=== UPDATING ${projectName} (${index + 1}/${selected.length}) ===`);
      console.log(`Completed so far: ${successfulProjects.join(', ') || 'None'}`);
      
      if (index > 0) {
        console.log('Clearing caches and waiting before project update...');
        clearSettingsCache();
        clearAllCommentColumnCaches();
        SpreadsheetApp.flush();
        
        console.log('Waiting 3 seconds before next project...');
        Utilities.sleep(3000);
      }
      
      updateProjectDataWithRetry(projectName);
      
      successfulProjects.push(projectName);
      console.log(`✅ ${projectName} updated successfully`);
      
      Utilities.sleep(2000);
      
    } catch (e) {
      console.error(`❌ Failed to update ${proj}:`, e);
      failedProjects.push({
        project: proj,
        error: e.toString().substring(0, 80)
      });
      
      console.log('Error occurred. Waiting 10 seconds before continuing...');
      Utilities.sleep(10000);
    }
  });
  
  if (successfulProjects.length > 0) {
    try {
      console.log('Waiting before sorting sheets...');
      Utilities.sleep(3000);
      sortProjectSheetsWithRetry();
    } catch (e) {
      console.error('Error sorting sheets:', e);
    }
  }
  
  console.log(`Update completed: ${successfulProjects.length}/${selected.length} projects updated`);
}

function updateAllProjects() {
  if (!isBearerTokenConfigured()) {
    openSettingsSheet();
    return;
  }
  
  var projects = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Incent_Traffic', 'Overall'];
  
  try {
    var successfulProjects = [];
    var failedProjects = [];
    
    projects.forEach(function(proj, index) {
      try {
        console.log(`Updating ${proj} (${index + 1}/${projects.length})...`);
        
        var projectName = proj.toUpperCase();
        updateProjectDataWithRetry(projectName);
        
        successfulProjects.push(proj);
        console.log(`✅ ${proj} updated successfully`);
        
        if (index < projects.length - 1) {
          Utilities.sleep(2000);
        }
        
      } catch (e) {
        console.error(`❌ Failed to update ${proj}:`, e);
        failedProjects.push({
          project: proj,
          error: e.toString().substring(0, 50)
        });
        
        if (index < projects.length - 1) {
          Utilities.sleep(3000);
        }
      }
    });
    
    if (successfulProjects.length > 0) {
      try {
        Utilities.sleep(1000);
        sortProjectSheetsWithRetry();
      } catch (e) {
        console.error('Error sorting sheets:', e);
      }
    }
    
    console.log(`Update completed: ${successfulProjects.length}/${projects.length} projects updated`);
    
  } catch (e) {
    console.error(`Update failed: ${e.toString()}`);
  }
}

function refreshSettings() {
  try {
    var settings = refreshSettingsFromSheet();
    
    try {
      syncTriggersWithSettings();
      console.log('Settings refreshed and triggers synchronized');
    } catch (e) {
      console.error('Error syncing triggers:', e);
    }
    
  } catch (e) {
    console.error('Error refreshing settings:', e);
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
  
  message += '💡 TIP: Use Settings sheet to configure all options';
  
  ui.alert('System Status', message, ui.ButtonSet.OK);
}

function quickAPICheckAll() {
  if (!isBearerTokenConfigured()) {
    openSettingsSheet();
    return;
  }
  
  var projects = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Incent_Traffic', 'Overall'];
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
  
  console.log(results);
}

function smartReportWizard() {
  if (!isBearerTokenConfigured()) {
    openSettingsSheet();
    return;
  }
  
  var scope = showChoice('📈 Generate Report - Step 1/2: Select Projects', ['All Projects', 'Single Project', 'Multiple Projects']);
  if (!scope) return;
  
  var weeks = promptWeeks('📅 Generate Report - Step 2/2: Select Weeks', 'Enter number of weeks (1-52):');
  if (!weeks) return;
  
  if (scope === 1) {
    generateAllProjects(weeks);
  } else if (scope === 2) {
    var project = showChoice('Select Project:', MENU_PROJECTS);
    if (!project) return;
    var projectName = MENU_PROJECTS[project-1].toUpperCase();
    generateSingleProject(projectName, weeks);
  } else {
    var selected = showMultiChoice('Select Projects:', MENU_PROJECTS);
    if (!selected || selected.length === 0) return;
    generateMultipleProjects(selected, weeks);
  }
}

function generateAllProjects(weeks) {
  var success = 0;
  var total = MENU_PROJECTS.length;
  
  try {
    for (var i = 0; i < MENU_PROJECTS.length; i++) {
      var proj = MENU_PROJECTS[i];
      try { 
        console.log(`Generating report for ${proj} (${weeks} weeks)...`);
        generateProjectReportByWeeks(proj.toUpperCase(), weeks); 
        success++; 
      } catch(e) { 
        console.error(`Error generating ${proj}:`, e); 
      }
    }
    
    sortProjectSheets();
    console.log(`Generated ${success}/${total} reports successfully for last ${weeks} weeks`);
  } catch(e) {
    console.error('Error generating all projects:', e);
  }
}

function generateSingleProject(projectName, weeks) {
  try {
    generateProjectReportByWeeks(projectName, weeks);
    sortProjectSheets();
    console.log(`${projectName} report generated successfully for last ${weeks} weeks`);
  } catch(e) {
    console.error(`Error generating ${projectName} report: ${e.toString()}`);
  }
}

function generateMultipleProjects(projects, weeks) {
  var success = 0;
  
  try {
    for (var i = 0; i < projects.length; i++) {
      var proj = projects[i];
      try {
        console.log(`Generating report for ${proj} (${weeks} weeks)...`);
        generateProjectReportByWeeks(proj.toUpperCase(), weeks);
        success++;
      } catch(e) {
        console.error(`Error generating ${proj}:`, e);
      }
    }
    
    sortProjectSheets();
    console.log(`Generated ${success}/${projects.length} reports successfully for last ${weeks} weeks`);
  } catch(e) {
    console.error('Error generating multiple projects:', e);
  }
}

function generateProjectReportByWeeks(projectName, weeks) {
  var days = weeks * 7;
  setCurrentProject(projectName);
  generateReport(days);
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
  }
}

function updateProjectDataWithRetry(projectName, maxRetries = 1) {
  var baseDelay = 3000;
  
  for (var attempt = 1; attempt <= maxRetries + 1; attempt++) {
    try {
      clearSettingsCache();
      SpreadsheetApp.flush();
      Utilities.sleep(1000);
      
      updateProjectData(projectName);
      return;
    } catch (e) {
      console.error(`${projectName} update attempt ${attempt} failed:`, e);
      
      if (attempt > maxRetries) {
        throw e;
      }
      
      if (e.toString().includes('timed out') || e.toString().includes('Service Spreadsheets')) {
        console.log('Timeout detected - waiting before retry...');
        
        clearSettingsCache();
        clearAllCommentColumnCaches();
        
        var timeoutDelay = baseDelay * 2;
        console.log(`Waiting ${timeoutDelay}ms before retry...`);
        Utilities.sleep(timeoutDelay);
        
        SpreadsheetApp.flush();
        Utilities.sleep(2000);
      } else {
        var delay = baseDelay;
        console.log(`Waiting ${delay}ms before retry...`);
        Utilities.sleep(delay);
      }
    }
  }
}

function sortProjectSheetsWithRetry(maxRetries = 1) {
  var baseDelay = 2000;
  
  for (var attempt = 1; attempt <= maxRetries + 1; attempt++) {
    try {
      sortProjectSheets();
      return;
    } catch (e) {
      console.error(`Sheet sorting attempt ${attempt} failed:`, e);
      
      if (attempt > maxRetries) {
        throw e;
      }
      
      var delay = baseDelay;
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

function promptWeeks(title, prompt) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(title, prompt + '\n\nCommon options: 4, 8, 12, 16, 20, 24', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  var weeks = parseInt(result.getResponseText());
  if (isNaN(weeks) || weeks < 1 || weeks > 52) {
    return null;
  }
  return weeks;
}

function generateProjectReport(projectName, days) { setCurrentProject(projectName); generateReport(days); }
function generateProjectReportForDateRange(projectName, startDate, endDate) { setCurrentProject(projectName); generateReportForDateRange(startDate, endDate); }
function debugProjectReportGeneration(projectName) { setCurrentProject(projectName); debugReportGeneration(); }

function appsDbWizard() {
  if (CURRENT_PROJECT !== 'TRICKY') {
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
  try {
    var appsDb = new AppsDatabase('TRICKY');
    var cache = appsDb.loadFromCache();
    var appCount = Object.keys(cache).length;
    
    console.log(`Apps Database Status: ${appCount} apps cached`);
    
    if (appCount > 0) {
      var bundleIds = Object.keys(cache);
      var sampleApp = cache[bundleIds[0]];
      console.log(`Last Updated: ${sampleApp.lastUpdated || 'Unknown'}`);
      console.log(`Cache Sheet: ${appsDb.cacheSheet ? 'Found' : 'Missing'}`);
      
      var shouldUpdate = appsDb.shouldUpdateCache();
      console.log(`Update Needed: ${shouldUpdate ? 'YES (>24h old)' : 'NO'}`);
      
      console.log('Sample entries:');
      var sampleCount = Math.min(3, bundleIds.length);
      for (var i = 0; i < sampleCount; i++) {
        var bundleId = bundleIds[i];
        var app = cache[bundleId];
        console.log(`${bundleId} → ${app.publisher} ${app.appName}`);
      }
    } else {
      console.log('Status: Empty cache - refresh needed');
    }
  } catch (e) {
    console.error('Error checking Apps Database status:', e);
  }
}

function showAppsDbSample() {
  try {
    var appsDb = new AppsDatabase('TRICKY');
    var cache = appsDb.loadFromCache();
    var bundleIds = Object.keys(cache);
    
    if (bundleIds.length === 0) {
      console.log('Apps Database cache is empty. Please refresh first.');
      return;
    }
    
    console.log('Apps Database Sample:');
    var sampleCount = Math.min(5, bundleIds.length);
    
    for (var i = 0; i < sampleCount; i++) {
      var bundleId = bundleIds[i];
      var app = cache[bundleId];
      console.log(`${bundleId} → ${app.publisher} ${app.appName}`);
    }
    
    if (bundleIds.length > sampleCount) {
      console.log(`... and ${bundleIds.length - sampleCount} more apps`);
    }
  } catch (e) {
    console.error('Error showing sample data:', e);
  }
}

function clearAppsDbCache() {
  var ui = SpreadsheetApp.getUi();
  
  if (ui.alert('Clear Apps Database Cache', 'Clear cached app data? Will rebuild on next refresh.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  
  try {
    var appsDb = new AppsDatabase('TRICKY');
    if (appsDb.cacheSheet && appsDb.cacheSheet.getLastRow() > 1) {
      appsDb.cacheSheet.deleteRows(2, appsDb.cacheSheet.getLastRow() - 1);
      console.log('Apps Database cache cleared');
    } else {
      console.log('Apps Database cache sheet not found');
    }
  } catch (e) {
    console.error('Error clearing cache:', e);
  }
}

function clearColumnCache() {
  clearAllCommentColumnCaches();
  console.log('Column cache cleared for all projects');
}

function openGitHubRepo() {
  var ui = SpreadsheetApp.getUi();
  var githubUrl = 'https://github.com/Poxagronka/pivot-table';
  
  var htmlOutput = HtmlService.createHtmlOutput(
    '<script>window.open("' + githubUrl + '", "_blank"); google.script.host.close();</script>'
  ).setWidth(400).setHeight(300);
  
  ui.showModalDialog(htmlOutput, 'Opening GitHub Repository...');
}