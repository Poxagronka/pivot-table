/**
 * Comment Cache Management - Multi Project Support
 * Handles saving and loading comments from a hidden sheet
 * Includes automatic caching functionality at 3AM daily
 * NOW SUPPORTS CAMPAIGN-LEVEL AND SOURCE_APP-LEVEL COMMENTS
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
      // Headers: AppName, WeekRange, Level, Identifier, SourceApp, Comment, LastUpdated
      sheet.getRange(1, 1, 1, 7).setValues([['AppName', 'WeekRange', 'Level', 'Identifier', 'SourceApp', 'Comment', 'LastUpdated']]);
    }
    return sheet;
  }

  /**
   * Generate a unique key for comment identification
   * Supports WEEK, SOURCE_APP, and CAMPAIGN level comments
   */
  getCommentKey(appName, weekRange, level, identifier = null, sourceApp = null) {
    return `${appName}|||${weekRange}|||${level}|||${identifier || 'N/A'}|||${sourceApp || 'N/A'}`;
  }

  /**
   * Load all comments from cache sheet
   */
  loadAllComments() {
    const comments = {};
    const data = this.cacheSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const [appName, weekRange, level, identifier, sourceApp, comment, lastUpdated] = data[i];
      if (comment) {
        const key = this.getCommentKey(appName, weekRange, level, identifier, sourceApp);
        comments[key] = comment;
      }
    }
    return comments;
  }

  /**
   * Save a comment to cache
   */
  saveComment(appName, weekRange, level, comment, identifier = null, sourceApp = null) {
    if (!comment || !comment.trim()) return;
    
    const data = this.cacheSheet.getDataRange().getValues();
    let found = false;
    
    // Update existing or add new
    for (let i = 1; i < data.length; i++) {
      const rowAppName = data[i][0];
      const rowWeekRange = data[i][1];
      const rowLevel = data[i][2];
      const rowIdentifier = data[i][3];
      const rowSourceApp = data[i][4];
      
      if (rowAppName === appName && 
          rowWeekRange === weekRange && 
          rowLevel === level &&
          rowIdentifier === (identifier || 'N/A') &&
          rowSourceApp === (sourceApp || 'N/A')) {
        // Only update if new comment is longer (appending text)
        const existingComment = data[i][5] || '';
        if (comment.length > existingComment.length) {
          this.cacheSheet.getRange(i + 1, 6, 1, 2).setValues([[comment, new Date()]]);
        }
        found = true;
        break;
      }
    }
    
    if (!found) {
      const lastRow = this.cacheSheet.getLastRow();
      this.cacheSheet.getRange(lastRow + 1, 1, 1, 7).setValues([[
        appName, 
        weekRange, 
        level,
        identifier || 'N/A', 
        sourceApp || 'N/A', 
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
    
    const data = sheet.getDataRange().getValues();
    let currentApp = '';
    let currentWeek = '';
    
    for (let i = 1; i < data.length; i++) {
      const level = data[i][0];
      const nameOrRange = data[i][1];
      const idOrEmpty = data[i][2]; // ID column
      const comment = data[i][15]; // Comments column (last column)
      
      if (level === 'APP') {
        currentApp = nameOrRange;
        currentWeek = '';
      } else if (level === 'WEEK' && currentApp) {
        currentWeek = nameOrRange;
        if (comment) {
          // Save week-level comment
          this.saveComment(currentApp, currentWeek, 'WEEK', comment);
        }
      } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
        // Save source app-level comment
        if (comment) {
          const sourceAppDisplayName = nameOrRange; // Publisher + App Name or bundle ID
          this.saveComment(currentApp, currentWeek, 'SOURCE_APP', comment, sourceAppDisplayName);
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek && comment) {
        // Save campaign-level comment
        const sourceAppName = nameOrRange; // Source App name
        const campaignIdValue = idOrEmpty && typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK') 
          ? this.extractCampaignIdFromHyperlink(idOrEmpty) 
          : idOrEmpty;
        
        this.saveComment(currentApp, currentWeek, 'CAMPAIGN', comment, campaignIdValue, sourceAppName);
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
      const idOrEmpty = data[i][2]; // ID column
      
      if (level === 'APP') {
        currentApp = nameOrRange;
        currentWeek = '';
      } else if (level === 'WEEK' && currentApp) {
        currentWeek = nameOrRange;
        // Apply week-level comment
        const weekKey = this.getCommentKey(currentApp, currentWeek, 'WEEK');
        const weekComment = comments[weekKey];
        if (weekComment) {
          sheet.getRange(i + 1, 16).setValue(weekComment); // Comments column
        }
      } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
        // Apply source app-level comment
        const sourceAppDisplayName = nameOrRange; // Publisher + App Name or bundle ID
        const sourceAppKey = this.getCommentKey(currentApp, currentWeek, 'SOURCE_APP', sourceAppDisplayName);
        const sourceAppComment = comments[sourceAppKey];
        if (sourceAppComment) {
          sheet.getRange(i + 1, 16).setValue(sourceAppComment); // Comments column
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek) {
        // Apply campaign-level comment
        const sourceAppName = nameOrRange; // Source App name
        const campaignIdValue = idOrEmpty && typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK') 
          ? this.extractCampaignIdFromHyperlink(idOrEmpty) 
          : idOrEmpty;
        
        const campaignKey = this.getCommentKey(currentApp, currentWeek, 'CAMPAIGN', campaignIdValue, sourceAppName);
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
 * Structure: APP groups contain WEEK groups, which contain SOURCE_APP groups (TRICKY only), which contain CAMPAIGN rows
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
  
  console.log(`Found ${groups.apps.length} app groups, ${groups.weeks.length} week groups, ${groups.sourceApps.length} source app groups`);
  
  // Collapse in order: deepest level first
  let totalCollapsed = 0;
  
  // 1. Collapse source app groups (deepest level for TRICKY)
  if (groups.sourceApps.length > 0) {
    console.log('Collapsing source app groups...');
    let sourceAppCollapsed = 0;
    groups.sourceApps.forEach((group, index) => {
      try {
        sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
        sourceAppCollapsed++;
        SpreadsheetApp.flush();
        Utilities.sleep(50);
        
        if (index % 5 === 0 && index > 0) {
          console.log(`  Collapsed ${index} source app groups...`);
        }
      } catch (e) {
        console.log(`  Could not collapse source app group at row ${group.start}: ${e.toString()}`);
      }
    });
    console.log(`  Collapsed ${sourceAppCollapsed} of ${groups.sourceApps.length} source app groups`);
    totalCollapsed += sourceAppCollapsed;
  }
  
  // 2. Collapse week groups
  console.log('Collapsing week groups...');
  let weekCollapsed = 0;
  groups.weeks.forEach((group, index) => {
    try {
      sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
      weekCollapsed++;
      SpreadsheetApp.flush();
      Utilities.sleep(50);
      
      if (index % 5 === 0 && index > 0) {
        console.log(`  Collapsed ${index} week groups...`);
      }
    } catch (e) {
      console.log(`  Could not collapse week group at row ${group.start}: ${e.toString()}`);
    }
  });
  console.log(`  Collapsed ${weekCollapsed} of ${groups.weeks.length} week groups`);
  totalCollapsed += weekCollapsed;
  
  // 3. Collapse app groups (top level)
  console.log('Collapsing app groups...');
  let appCollapsed = 0;
  groups.apps.forEach((group, index) => {
    try {
      sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
      appCollapsed++;
      SpreadsheetApp.flush();
      Utilities.sleep(50);
    } catch (e) {
      console.log(`  Could not collapse app group at row ${group.start}: ${e.toString()}`);
    }
  });
  console.log(`  Collapsed ${appCollapsed} of ${groups.apps.length} app groups`);
  totalCollapsed += appCollapsed;
  
  console.log(`Recursive collapse completed: ${totalCollapsed} total groups collapsed`);
}

/**
 * Identify all groups in the sheet data
 * Returns object with arrays of app, week, and source app groups
 * Campaigns are not groups, just regular rows
 */
function identifyGroups(data) {
  const groups = {
    apps: [],
    weeks: [],
    sourceApps: []
  };
  
  let currentApp = null;
  let appStartRow = null;
  let weekStartRow = null;
  let sourceAppStartRow = null;
  
  for (let i = 1; i < data.length; i++) {
    const level = data[i][0];
    
    if (level === 'APP') {
      // Close previous source app if exists
      if (sourceAppStartRow !== null && i > sourceAppStartRow + 1) {
        groups.sourceApps.push({
          start: sourceAppStartRow + 1,
          count: i - sourceAppStartRow - 1,
          app: currentApp
        });
      }
      
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
      sourceAppStartRow = null;
      
    } else if (level === 'WEEK') {
      // Close previous source app if exists
      if (sourceAppStartRow !== null && i > sourceAppStartRow + 1) {
        groups.sourceApps.push({
          start: sourceAppStartRow + 1,
          count: i - sourceAppStartRow - 1,
          app: currentApp
        });
      }
      
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
      sourceAppStartRow = null;
      
    } else if (level === 'SOURCE_APP') {
      // Close previous source app if exists
      if (sourceAppStartRow !== null && i > sourceAppStartRow + 1) {
        groups.sourceApps.push({
          start: sourceAppStartRow + 1,
          count: i - sourceAppStartRow - 1,
          app: currentApp
        });
      }
      
      // Start new source app
      sourceAppStartRow = i;
      
    } else if (level === 'CAMPAIGN') {
      // Campaigns are just rows, not groups - nothing to do
      continue;
    }
  }
  
  // Close any remaining groups
  if (sourceAppStartRow !== null && data.length > sourceAppStartRow + 1) {
    groups.sourceApps.push({
      start: sourceAppStartRow + 1,
      count: data.length - sourceAppStartRow - 1,
      app: currentApp
    });
  }
  
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