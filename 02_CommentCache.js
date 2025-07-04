/**
 * Comment Cache Management - Multi Project Support
 * Handles saving and loading comments from a hidden sheet
 * Includes automatic caching functionality at 2AM daily
 * NOW SUPPORTS CAMPAIGN-LEVEL COMMENTS
 */
class CommentCache {
  constructor(projectName = null) {
    // If no project specified, use current project
    this.projectName = projectName || CURRENT_PROJECT;
    this.config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
    this.cacheSheet = this.getOrCreateCacheSheet();
  }

  /**
   * Get or create the comments cache sheet for current project
   */
  getOrCreateCacheSheet() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    let sheet = spreadsheet.getSheetByName(this.config.COMMENTS_CACHE_SHEET);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(this.config.COMMENTS_CACHE_SHEET);
      sheet.hideSheet();
      // Headers: AppName, WeekRange, CampaignId, SourceApp, Comment, LastUpdated
      sheet.getRange(1, 1, 1, 6).setValues([['AppName', 'WeekRange', 'CampaignId', 'SourceApp', 'Comment', 'LastUpdated']]);
    }
    return sheet;
  }

  /**
   * Generate a unique key for comment identification
   * Supports both WEEK-level and CAMPAIGN-level comments
   */
  getCommentKey(appName, weekRange, campaignId = null, sourceApp = null) {
    if (campaignId && sourceApp) {
      // Campaign-level comment
      return `${appName}|||${weekRange}|||${campaignId}|||${sourceApp}`;
    } else {
      // Week-level comment
      return `${appName}|||${weekRange}|||WEEK|||WEEK`;
    }
  }

  /**
   * Load all comments from cache sheet
   */
  loadAllComments() {
    const comments = {};
    const data = this.cacheSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const [appName, weekRange, campaignId, sourceApp, comment, lastUpdated] = data[i];
      if (comment) {
        const key = this.getCommentKey(appName, weekRange, campaignId, sourceApp);
        comments[key] = comment;
      }
    }
    return comments;
  }

  /**
   * Save a comment to cache
   */
  saveComment(appName, weekRange, comment, campaignId = null, sourceApp = null) {
    if (!comment || !comment.trim()) return;
    
    const data = this.cacheSheet.getDataRange().getValues();
    let found = false;
    
    // Update existing or add new
    for (let i = 1; i < data.length; i++) {
      const rowAppName = data[i][0];
      const rowWeekRange = data[i][1];
      const rowCampaignId = data[i][2];
      const rowSourceApp = data[i][3];
      
      if (rowAppName === appName && 
          rowWeekRange === weekRange && 
          rowCampaignId === (campaignId || 'WEEK') && 
          rowSourceApp === (sourceApp || 'WEEK')) {
        // Only update if new comment is longer (appending text)
        const existingComment = data[i][4] || '';
        if (comment.length > existingComment.length) {
          this.cacheSheet.getRange(i + 1, 5, 1, 2).setValues([[comment, new Date()]]);
        }
        found = true;
        break;
      }
    }
    
    if (!found) {
      const lastRow = this.cacheSheet.getLastRow();
      this.cacheSheet.getRange(lastRow + 1, 1, 1, 6).setValues([[
        appName, 
        weekRange, 
        campaignId || 'WEEK', 
        sourceApp || 'WEEK', 
        comment, 
        new Date()
      ]]);
    }
  }

  /**
   * Sync comments from the main sheet to cache
   * ОПТИМИЗИРОВАНО: Не раскрывает группы - getDataRange читает все данные
   */
  syncCommentsFromSheet() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(this.config.SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return;
    
    // УБРАНО: expandAllGroups(sheet) - не нужно!
    
    const data = sheet.getDataRange().getValues();
    let currentApp = '';
    let currentWeek = '';
    
    for (let i = 1; i < data.length; i++) {
      const level = data[i][0];
      const nameOrRange = data[i][1];
      const campaignId = data[i][2]; // ID column
      const comment = data[i][15]; // Comments column
      
      if (level === 'APP') {
        currentApp = nameOrRange;
        currentWeek = '';
      } else if (level === 'WEEK' && currentApp) {
        currentWeek = nameOrRange;
        if (comment) {
          // Save week-level comment
          this.saveComment(currentApp, currentWeek, comment);
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek && comment) {
        // Save campaign-level comment
        const sourceApp = nameOrRange; // Source App name
        const campaignIdValue = campaignId && typeof campaignId === 'string' && campaignId.includes('HYPERLINK') 
          ? this.extractCampaignIdFromHyperlink(campaignId) 
          : campaignId;
        
        this.saveComment(currentApp, currentWeek, comment, campaignIdValue, sourceApp);
      }
    }
  }

  /**
   * Extract campaign ID from HYPERLINK formula
   */
  extractCampaignIdFromHyperlink(hyperlinkFormula) {
    try {
      // Extract from =HYPERLINK("https://app.appgrowth.com/campaigns/12345", "12345")
      const match = hyperlinkFormula.match(/campaigns\/([^"]+)/);
      return match ? match[1] : 'Unknown';
    } catch (e) {
      return 'Unknown';
    }
  }

  /**
   * Apply cached comments back to the main sheet
   */
  applyCommentsToSheet() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(this.config.SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return;
    
    const comments = this.loadAllComments();
    const data = sheet.getDataRange().getValues();
    let currentApp = '';
    let currentWeek = '';
    
    for (let i = 1; i < data.length; i++) {
      const level = data[i][0];
      const nameOrRange = data[i][1];
      const campaignId = data[i][2]; // ID column
      
      if (level === 'APP') {
        currentApp = nameOrRange;
        currentWeek = '';
      } else if (level === 'WEEK' && currentApp) {
        currentWeek = nameOrRange;
        // Apply week-level comment
        const weekKey = this.getCommentKey(currentApp, currentWeek);
        const weekComment = comments[weekKey];
        if (weekComment) {
          sheet.getRange(i + 1, 16).setValue(weekComment); // Comments column
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek) {
        // Apply campaign-level comment
        const sourceApp = nameOrRange; // Source App name
        const campaignIdValue = campaignId && typeof campaignId === 'string' && campaignId.includes('HYPERLINK') 
          ? this.extractCampaignIdFromHyperlink(campaignId) 
          : campaignId;
        
        const campaignKey = this.getCommentKey(currentApp, currentWeek, campaignIdValue, sourceApp);
        const campaignComment = comments[campaignKey];
        if (campaignComment) {
          sheet.getRange(i + 1, 16).setValue(campaignComment); // Comments column
        }
      }
    }
  }

  /**
   * Sync comments from the main sheet to cache (quiet version)
   * ПЕРЕИМЕНОВАНО: теперь идентична обычной версии, так как раскрытие групп убрано
   */
  syncCommentsFromSheetQuiet() {
    // Теперь идентична syncCommentsFromSheet(), так как мы убрали expandAllGroups
    this.syncCommentsFromSheet();
  }
}

/**
 * GROUP MANAGEMENT
 * Enhanced group management for reliable collapsing
 */

/**
 * Collapse all groups recursively - most reliable method
 * Collapses groups one by one, starting from deepest level
 * Structure: APP groups contain WEEK groups, which contain CAMPAIGN rows (not groups)
 */
function collapseAllGroupsRecursively(sheet) {
  console.log('Starting recursive collapse (one-by-one method)...');
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    console.log('No data to process');
    return;
  }
  
  // First pass: Identify all group boundaries
  const groups = identifyGroups(data);
  
  console.log(`Found ${groups.apps.length} app groups, ${groups.weeks.length} week groups`);
  
  // Collapse in order: weeks first (deeper level), then apps
  let totalCollapsed = 0;
  
  // 1. Collapse week groups under each app
  console.log('Collapsing week groups...');
  let weekCollapsed = 0;
  groups.weeks.forEach((group, index) => {
    try {
      sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
      weekCollapsed++;
      SpreadsheetApp.flush();
      Utilities.sleep(50); // Small delay between each collapse
      
      if (index % 5 === 0 && index > 0) {
        console.log(`  Collapsed ${index} week groups...`);
      }
    } catch (e) {
      // Group might already be collapsed or not exist
      console.log(`  Could not collapse week group at row ${group.start}: ${e.toString()}`);
    }
  });
  console.log(`  Collapsed ${weekCollapsed} of ${groups.weeks.length} week groups`);
  totalCollapsed += weekCollapsed;
  
  // 2. Collapse app groups
  console.log('Collapsing app groups...');
  let appCollapsed = 0;
  groups.apps.forEach((group, index) => {
    try {
      sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
      appCollapsed++;
      SpreadsheetApp.flush();
      Utilities.sleep(50); // Small delay between each collapse
    } catch (e) {
      // Group might already be collapsed or not exist
      console.log(`  Could not collapse app group at row ${group.start}: ${e.toString()}`);
    }
  });
  console.log(`  Collapsed ${appCollapsed} of ${groups.apps.length} app groups`);
  totalCollapsed += appCollapsed;
  
  console.log(`Recursive collapse completed: ${totalCollapsed} total groups collapsed`);
}

/**
 * Identify all groups in the sheet data
 * Returns object with arrays of app and week groups
 * Campaigns are not groups, just regular rows
 */
function identifyGroups(data) {
  const groups = {
    apps: [],
    weeks: []
  };
  
  let currentApp = null;
  let appStartRow = null;
  let weekStartRow = null;
  
  for (let i = 1; i < data.length; i++) {
    const level = data[i][0];
    const nextLevel = i + 1 < data.length ? data[i + 1][0] : null;
    
    if (level === 'APP') {
      // Close previous week if exists
      if (weekStartRow !== null && i > weekStartRow + 1) {
        groups.weeks.push({
          start: weekStartRow + 1,
          count: i - weekStartRow - 1,
          app: currentApp
        });
      }
      
      // Close previous app if exists
      if (appStartRow !== null && i > appStartRow + 1) {
        groups.apps.push({
          start: appStartRow + 1,
          count: i - appStartRow - 1,
          name: currentApp
        });
      }
      
      // Start new app
      currentApp = data[i][1];
      appStartRow = i;
      weekStartRow = null;
      
    } else if (level === 'WEEK') {
      // Close previous week if exists
      if (weekStartRow !== null && i > weekStartRow + 1) {
        groups.weeks.push({
          start: weekStartRow + 1,
          count: i - weekStartRow - 1,
          app: currentApp
        });
      }
      
      // Start new week
      weekStartRow = i;
      
    } else if (level === 'CAMPAIGN') {
      // Campaigns are just rows, not groups - nothing to do
      continue;
    }
  }
  
  // Close any remaining groups
  if (weekStartRow !== null && data.length > weekStartRow + 1) {
    groups.weeks.push({
      start: weekStartRow + 1,
      count: data.length - weekStartRow - 1,
      app: currentApp
    });
  }
  
  if (appStartRow !== null && data.length > appStartRow + 1) {
    groups.apps.push({
      start: appStartRow + 1,
      count: data.length - appStartRow - 1,
      name: currentApp
    });
  }
  
  return groups;
}

/**
 * Expand all groups in the sheet - used ONLY when user explicitly needs to see data
 */
function expandAllGroups(sheet) {
  try {
    const maxRows = sheet.getMaxRows();
    let expanded = true;
    let attempts = 0;
    const maxAttempts = 10;
    
    while (expanded && attempts < maxAttempts) {
      attempts++;
      try {
        sheet.getRange(1, 1, maxRows, 1).expandGroups();
        SpreadsheetApp.flush();
        Utilities.sleep(50);
      } catch (e) {
        // No more groups to expand
        expanded = false;
      }
    }
    
    console.log(`Groups expanded after ${attempts} attempts`);
  } catch (e) {
    console.log('Error expanding groups:', e);
  }
}