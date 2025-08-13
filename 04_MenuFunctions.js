var MENU_PROJECTS = getAllProjectNames();

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Campaign Report')
    .addItem('📈 Generate Report...', 'smartReportWizard')
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
  var selected = showMultiChoice('Select Projects to Update:', MENU_PROJECTS);
  if (selected?.length) updateProjectsBatch(selected, true);
}

function updateAllProjects() { updateProjectsBatch(MENU_PROJECTS, false); }

function updateProjectsBatch(projects, isSelective = false) {
  if (!isBearerTokenConfigured()) return openSettingsSheet();
  
  if (isSelective) try { preloadSettings(); Utilities.sleep(200); } catch(e) {}
  
  var results = { success: [], failed: [] };
  
  projects.forEach((proj, i) => {
    try {
      if (isSelective && i > 0) {
        clearSettingsCache();
        clearAllCommentColumnCaches();
        SpreadsheetApp.flush();
        Utilities.sleep(300);
      }
      updateProjectDataWithRetry(proj.toUpperCase());
      results.success.push(proj);
      if (i < projects.length - 1) Utilities.sleep(200);
    } catch (e) {
      console.error(`❌ Failed ${proj}:`, e);
      results.failed.push({project: proj, error: e.toString().substring(0, 80)});
      if (i < projects.length - 1) Utilities.sleep(isSelective ? 1000 : 300);
    }
  });
  
  if (results.success.length) try { Utilities.sleep(isSelective ? 300 : 100); sortProjectSheetsWithRetry(); } catch(e) {}
  console.log(`Updated: ${results.success.length}/${projects.length}`);
}

function refreshSettings() {
  try {
    refreshSettingsFromSheet();
    try { syncTriggersWithSettings(); } catch(e) { console.error('Trigger sync error:', e); }
  } catch(e) { console.error('Settings refresh error:', e); }
}

function showQuickStatus() {
  refreshSettingsFromSheet();
  var token = isBearerTokenConfigured() ? '✅' : '❌';
  var cache = isAutoCacheEnabled() ? '✅' : '❌';
  var update = isAutoUpdateEnabled() ? '✅' : '❌';
  
  var triggers = ScriptApp.getProjectTriggers();
  var ct = triggers.find(t => t.getHandlerFunction() === 'autoCacheAllProjects');
  var ut = triggers.find(t => t.getHandlerFunction() === 'autoUpdateAllProjects');
  
  var issues = [];
  if (isAutoCacheEnabled() && !ct) issues.push('• Cache trigger missing');
  if (!isAutoCacheEnabled() && ct) issues.push('• Cache trigger exists but disabled');
  if (isAutoUpdateEnabled() && !ut) issues.push('• Update trigger missing');
  if (!isAutoUpdateEnabled() && ut) issues.push('• Update trigger exists but disabled');
  
  var msg = `📊 SYSTEM STATUS\n\n🔐 Token: ${token}\n💾 Cache: ${cache}\n🔄 Update: ${update}\n🎯 Metrics: eROAS D730\n\n`;
  if (issues.length) msg += '⚠️ ISSUES:\n' + issues.join('\n') + '\n\nUse "🔄 Refresh Settings" to fix\n\n';
  else msg += '✅ All triggers synchronized\n\n';
  msg += '📅 SCHEDULE:\n• Cache: Every hour\n• Update: Daily 5AM\n• Includes previous week from Tuesday\n\n💡 TIP: Use Settings sheet';
  
  SpreadsheetApp.getUi().alert('System Status', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function quickAPICheckAll() {
  if (!isBearerTokenConfigured()) return openSettingsSheet();
  
  var results = '🔍 API CHECK\n\n';
  MENU_PROJECTS.forEach(proj => {
    try {
      setCurrentProject(proj);
      var raw = fetchCampaignData(getDateRange(7));
      var count = raw.data?.analytics?.richStats?.stats?.length || 0;
      results += `✅ ${proj}: ${count} records\n`;
    } catch(e) { results += `❌ ${proj}: ${e.toString().substring(0, 30)}...\n`; }
  });
  console.log(results);
}

function smartReportWizard() {
  if (!isBearerTokenConfigured()) return openSettingsSheet();
  
  var scope = showChoice('📈 Generate Report - Step 1/2', ['All Projects', 'Single Project', 'Multiple Projects']);
  if (!scope) return;
  
  var weeks = promptWeeks('📅 Step 2/2: Weeks', 'Enter weeks (1-52):');
  if (!weeks) return;
  
  if (scope === 1) generateProjectsBatch(MENU_PROJECTS, weeks);
  else if (scope === 2) {
    var p = showChoice('Select Project:', MENU_PROJECTS);
    if (p) generateProjectsBatch([MENU_PROJECTS[p-1]], weeks);
  } else {
    var sel = showMultiChoice('Select Projects:', MENU_PROJECTS);
    if (sel?.length) generateProjectsBatch(sel, weeks);
  }
}

function generateProjectsBatch(projects, weeks) {
  var success = 0;
  projects.forEach(proj => {
    try {
      setCurrentProject(proj.toUpperCase());
      generateReport(weeks * 7);
      success++;
    } catch(e) { console.error(`Error ${proj}:`, e); }
  });
  if (success) sortProjectSheets();
  console.log(`Generated ${success}/${projects.length} for ${weeks} weeks`);
}

function updateProjectDataWithRetry(projectName, maxRetries = 1) {
  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      clearSettingsCache();
      SpreadsheetApp.flush();
      Utilities.sleep(100);
      updateProjectData(projectName);
      return;
    } catch(e) {
      console.error(`${projectName} attempt ${attempt+1} failed:`, e);
      if (attempt >= maxRetries) throw e;
      if (e.toString().includes('timed out') || e.toString().includes('Service Spreadsheets')) {
        clearSettingsCache();
        clearAllCommentColumnCaches();
        Utilities.sleep(600 * (attempt + 1));
        SpreadsheetApp.flush();
        Utilities.sleep(200);
      } else Utilities.sleep(300);
    }
  }
}

function sortProjectSheetsWithRetry(maxRetries = 1) {
  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    try { sortProjectSheets(); return; }
    catch(e) {
      console.error(`Sort attempt ${attempt+1} failed:`, e);
      if (attempt >= maxRetries) throw e;
      Utilities.sleep(200);
    }
  }
}

function showChoice(title, options) {
  var ui = SpreadsheetApp.getUi();
  var list = options.map((o, i) => `${i+1} - ${o}`).join('\n');
  var result = ui.prompt(title, list + '\n\nEnter number:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  var n = parseInt(result.getResponseText());
  return (n >= 1 && n <= options.length) ? n : null;
}

function showMultiChoice(title, options) {
  var ui = SpreadsheetApp.getUi();
  var list = options.map((o, i) => `${i+1} - ${o}`).join('\n');
  var result = ui.prompt(title, list + '\n\nEnter numbers (e.g., 1,3):', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  return result.getResponseText().split(',')
    .map(s => parseInt(s.trim()))
    .filter(n => n >= 1 && n <= options.length)
    .map(n => options[n-1]);
}

function promptWeeks(title, prompt) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(title, prompt + '\n\nCommon: 4, 8, 12, 16, 20, 24', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  var w = parseInt(result.getResponseText());
  return (w >= 1 && w <= 52) ? w : null;
}

function appsDbWizard() {
  if (CURRENT_PROJECT !== 'TRICKY') setCurrentProject('TRICKY');
  var actions = ['View Status', 'Refresh', 'Sample Data', 'Clear Cache', 'Debug'];
  var choice = showChoice('📱 Apps Database', actions);
  if (!choice) return;
  [showAppsDbStatus, refreshAppsDatabase, showAppsDbSample, clearAppsDbCache, debugAppsDatabase][choice-1]();
}

function showAppsDbStatus() {
  try {
    var db = new AppsDatabase('TRICKY');
    var cache = db.loadFromCache();
    var count = Object.keys(cache).length;
    console.log(`Apps DB: ${count} apps${count ? (db.shouldUpdateCache() ? ' (update needed)' : '') : ' (empty)'}`);
  } catch(e) { console.error('Apps DB error:', e); }
}

function showAppsDbSample() {
  try {
    var db = new AppsDatabase('TRICKY');
    var cache = db.loadFromCache();
    var keys = Object.keys(cache);
    if (!keys.length) return console.log('Apps DB empty');
    console.log(`Apps DB: ${keys.length} total`);
    keys.slice(0, 3).forEach(k => console.log(`${k} → ${cache[k].publisher} ${cache[k].appName}`));
  } catch(e) { console.error('Sample error:', e); }
}

function clearAppsDbCache() {
  if (SpreadsheetApp.getUi().alert('Clear Apps Database?', 'Clear cached app data?', SpreadsheetApp.getUi().ButtonSet.YES_NO) !== SpreadsheetApp.getUi().Button.YES) return;
  try {
    var db = new AppsDatabase('TRICKY');
    if (db.cacheSheet?.getLastRow() > 1) {
      db.cacheSheet.deleteRows(2, db.cacheSheet.getLastRow() - 1);
      console.log('Apps DB cleared');
    }
  } catch(e) { console.error('Clear error:', e); }
}

function syncTriggersWithSettings() {
  try {
    var s = loadSettingsFromSheet();
    var triggers = ScriptApp.getProjectTriggers();
    var ct = triggers.find(t => t.getHandlerFunction() === 'autoCacheAllProjects');
    var ut = triggers.find(t => t.getHandlerFunction() === 'autoUpdateAllProjects');
    
    if (s.automation.autoCache && !ct) ScriptApp.newTrigger('autoCacheAllProjects').timeBased().atHour(2).everyDays(1).create();
    else if (!s.automation.autoCache && ct) ScriptApp.deleteTrigger(ct);
    
    if (s.automation.autoUpdate && !ut) ScriptApp.newTrigger('autoUpdateAllProjects').timeBased().atHour(5).everyDays(1).create();
    else if (!s.automation.autoUpdate && ut) ScriptApp.deleteTrigger(ut);
  } catch(e) { console.error('Trigger sync error:', e); }
}

function openGitHubRepo() {
  var html = HtmlService.createHtmlOutput('<script>window.open("https://github.com/Poxagronka/pivot-table","_blank");google.script.host.close();</script>')
    .setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening GitHub...');
}

function clearColumnCache() { clearAllCommentColumnCaches(); console.log('Column cache cleared'); }
function debugSingleProject() { var p = showChoice('Debug Project:', MENU_PROJECTS); if (p) debugProjectReportGeneration(MENU_PROJECTS[p-1].toUpperCase()); }
function generateAllProjects(weeks) { generateProjectsBatch(MENU_PROJECTS, weeks); }
function generateSingleProject(projectName, weeks) { generateProjectsBatch([projectName], weeks); }
function generateMultipleProjects(projects, weeks) { generateProjectsBatch(projects, weeks); }
function generateProjectReport(projectName, days) { setCurrentProject(projectName); generateReport(days); }
function generateProjectReportByWeeks(projectName, weeks) { setCurrentProject(projectName); generateReport(weeks * 7); }
function generateProjectReportForDateRange(projectName, startDate, endDate) { setCurrentProject(projectName); generateReportForDateRange(startDate, endDate); }
function debugProjectReportGeneration(projectName) { setCurrentProject(projectName); debugReportGeneration(); }