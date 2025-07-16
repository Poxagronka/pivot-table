class CommentCache {
  constructor(projectName = null) {
    this.projectName = projectName || CURRENT_PROJECT;
    this.config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
    this.cacheSheet = this.getOrCreateCacheSheet();
    this.COMMENT_COLUMN = 18;
  }

  getOrCreateCacheSheet() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    let sheet = spreadsheet.getSheetByName(this.config.COMMENTS_CACHE_SHEET);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(this.config.COMMENTS_CACHE_SHEET);
      sheet.hideSheet();
      sheet.getRange(1, 1, 1, 6).setValues([['AppName', 'WeekRange', 'CampaignId', 'SourceApp', 'Comment', 'LastUpdated']]);
    }
    return sheet;
  }

  getCommentKey(appName, weekRange, campaignId, sourceApp) {
    return `${appName}|||${weekRange}|||${campaignId || 'WEEK'}|||${sourceApp || 'N/A'}`;
  }

  loadAllComments() {
    const data = this.cacheSheet.getDataRange().getValues();
    const comments = {};
    
    for (let i = 1; i < data.length; i++) {
      const [appName, weekRange, campaignId, sourceApp, comment, lastUpdated] = data[i];
      if (comment && comment.trim()) {
        const key = this.getCommentKey(appName, weekRange, campaignId, sourceApp);
        comments[key] = {
          comment: comment.trim(),
          lastUpdated: lastUpdated
        };
      }
    }
    
    console.log(`${this.projectName}: Loaded ${Object.keys(comments).length} comments from cache`);
    return comments;
  }

  syncCommentsFromSheet() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(this.config.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      console.log(`${this.projectName}: No data to sync`);
      return;
    }

    const data = sheet.getDataRange().getValues();
    const commentsToSave = [];
    let currentApp = '';
    let currentWeek = '';
    let currentSourceApp = '';
    
    console.log(`${this.projectName}: Syncing ${data.length} rows...`);

    for (let i = 1; i < data.length; i++) {
      const level = data[i][0];
      const nameOrRange = data[i][1];
      const idOrEmpty = data[i][2];
      const comment = data[i][this.COMMENT_COLUMN];
      
      if (level === 'APP') {
        currentApp = nameOrRange;
        currentWeek = '';
        currentSourceApp = '';
      } else if (level === 'WEEK' && currentApp) {
        currentWeek = nameOrRange;
        currentSourceApp = '';
        if (comment && comment.trim()) {
          commentsToSave.push({
            appName: currentApp,
            weekRange: currentWeek,
            campaignId: 'WEEK',
            sourceApp: 'N/A',
            comment: comment.trim()
          });
        }
      } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
        currentSourceApp = nameOrRange;
        if (comment && comment.trim()) {
          commentsToSave.push({
            appName: currentApp,
            weekRange: currentWeek,
            campaignId: 'SOURCE_APP',
            sourceApp: currentSourceApp,
            comment: comment.trim()
          });
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek && comment && comment.trim()) {
        let campaignId = idOrEmpty;
        if (typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK')) {
          const match = idOrEmpty.match(/campaigns\/([^"]+)/);
          campaignId = match ? match[1] : 'Unknown';
        }
        
        commentsToSave.push({
          appName: currentApp,
          weekRange: currentWeek,
          campaignId: campaignId,
          sourceApp: currentSourceApp || nameOrRange,
          comment: comment.trim()
        });
      }
    }

    console.log(`${this.projectName}: Found ${commentsToSave.length} comments to save`);
    
    if (commentsToSave.length > 0) {
      this.batchSaveComments(commentsToSave);
    }
  }

  batchSaveComments(commentsToSave) {
    const existingData = this.cacheSheet.getDataRange().getValues();
    const existingComments = {};
    
    for (let i = 1; i < existingData.length; i++) {
      const [appName, weekRange, campaignId, sourceApp] = existingData[i];
      const key = this.getCommentKey(appName, weekRange, campaignId, sourceApp);
      existingComments[key] = i + 1;
    }

    const newRows = [];
    const updateRows = [];
    
    commentsToSave.forEach(commentData => {
      const key = this.getCommentKey(
        commentData.appName,
        commentData.weekRange,
        commentData.campaignId,
        commentData.sourceApp
      );
      
      const row = [
        commentData.appName,
        commentData.weekRange,
        commentData.campaignId,
        commentData.sourceApp,
        commentData.comment,
        new Date()
      ];
      
      if (existingComments[key]) {
        updateRows.push({ row: existingComments[key], data: row });
      } else {
        newRows.push(row);
      }
    });

    updateRows.forEach(update => {
      this.cacheSheet.getRange(update.row, 1, 1, 6).setValues([update.data]);
    });

    if (newRows.length > 0) {
      const startRow = this.cacheSheet.getLastRow() + 1;
      this.cacheSheet.getRange(startRow, 1, newRows.length, 6).setValues(newRows);
    }

    console.log(`${this.projectName}: Updated ${updateRows.length} comments, added ${newRows.length} new`);
  }

  applyCommentsToSheet() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(this.config.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      console.log(`${this.projectName}: No sheet data to apply comments to`);
      return;
    }

    const comments = this.loadAllComments();
    if (Object.keys(comments).length === 0) {
      console.log(`${this.projectName}: No comments to apply`);
      return;
    }

    const data = sheet.getDataRange().getValues();
    const commentUpdates = [];
    let currentApp = '';
    let currentWeek = '';
    let currentSourceApp = '';

    for (let i = 1; i < data.length; i++) {
      const level = data[i][0];
      const nameOrRange = data[i][1];
      const idOrEmpty = data[i][2];
      
      if (level === 'APP') {
        currentApp = nameOrRange;
        currentWeek = '';
        currentSourceApp = '';
      } else if (level === 'WEEK' && currentApp) {
        currentWeek = nameOrRange;
        currentSourceApp = '';
        
        const weekKey = this.getCommentKey(currentApp, currentWeek, 'WEEK', 'N/A');
        const weekComment = comments[weekKey];
        if (weekComment) {
          commentUpdates.push({ row: i + 1, comment: weekComment.comment });
        }
      } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
        currentSourceApp = nameOrRange;
        
        const sourceAppKey = this.getCommentKey(currentApp, currentWeek, 'SOURCE_APP', currentSourceApp);
        const sourceAppComment = comments[sourceAppKey];
        if (sourceAppComment) {
          commentUpdates.push({ row: i + 1, comment: sourceAppComment.comment });
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek) {
        let campaignId = idOrEmpty;
        if (typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK')) {
          const match = idOrEmpty.match(/campaigns\/([^"]+)/);
          campaignId = match ? match[1] : 'Unknown';
        }
        
        const campaignKey = this.getCommentKey(currentApp, currentWeek, campaignId, currentSourceApp || nameOrRange);
        const campaignComment = comments[campaignKey];
        if (campaignComment) {
          commentUpdates.push({ row: i + 1, comment: campaignComment.comment });
        }
      }
    }

    console.log(`${this.projectName}: Applying ${commentUpdates.length} comments`);
    
    if (commentUpdates.length > 0) {
      const batchSize = 100;
      for (let i = 0; i < commentUpdates.length; i += batchSize) {
        const batch = commentUpdates.slice(i, i + batchSize);
        batch.forEach(update => {
          sheet.getRange(update.row, 19).setValue(update.comment);
        });
        
        if (i + batchSize < commentUpdates.length) {
          SpreadsheetApp.flush();
          Utilities.sleep(200);
        }
      }
    }
  }

  clearCache() {
    if (this.cacheSheet.getLastRow() > 1) {
      this.cacheSheet.deleteRows(2, this.cacheSheet.getLastRow() - 1);
    }
    console.log(`${this.projectName}: Cache cleared`);
  }

  syncCommentsFromSheetQuiet() {
    this.syncCommentsFromSheet();
  }
}

function collapseAllGroupsRecursively(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  
  const groups = identifyGroups(data);
  let totalCollapsed = 0;
  
  if (groups.sourceApps.length > 0) {
    groups.sourceApps.forEach(group => {
      try {
        sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
        totalCollapsed++;
      } catch (e) {}
    });
  }
  
  groups.weeks.forEach(group => {
    try {
      sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
      totalCollapsed++;
    } catch (e) {}
  });
  
  groups.apps.forEach(group => {
    try {
      sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
      totalCollapsed++;
    } catch (e) {}
  });
  
  console.log(`Collapsed ${totalCollapsed} groups`);
}

function identifyGroups(data) {
  const groups = { apps: [], weeks: [], sourceApps: [] };
  let currentApp = null;
  let appStartRow = null;
  let weekStartRow = null;
  let sourceAppStartRow = null;
  
  for (let i = 1; i < data.length; i++) {
    const level = data[i][0];
    
    if (level === 'APP') {
      if (sourceAppStartRow !== null && i > sourceAppStartRow + 1) {
        groups.sourceApps.push({
          start: sourceAppStartRow + 1,
          count: i - sourceAppStartRow - 1,
          app: currentApp
        });
      }
      
      if (weekStartRow !== null && i > weekStartRow + 1) {
        groups.weeks.push({
          start: weekStartRow + 1,
          count: i - weekStartRow - 1,
          app: currentApp
        });
      }
      
      if (appStartRow !== null && i > appStartRow + 1) {
        groups.apps.push({
          start: appStartRow + 1,
          count: i - appStartRow - 1,
          name: currentApp
        });
      }
      
      currentApp = data[i][1];
      appStartRow = i;
      weekStartRow = null;
      sourceAppStartRow = null;
      
    } else if (level === 'WEEK') {
      if (sourceAppStartRow !== null && i > sourceAppStartRow + 1) {
        groups.sourceApps.push({
          start: sourceAppStartRow + 1,
          count: i - sourceAppStartRow - 1,
          app: currentApp
        });
      }
      
      if (weekStartRow !== null && i > weekStartRow + 1) {
        groups.weeks.push({
          start: weekStartRow + 1,
          count: i - weekStartRow - 1,
          app: currentApp
        });
      }
      
      weekStartRow = i;
      sourceAppStartRow = null;
      
    } else if (level === 'SOURCE_APP') {
      if (sourceAppStartRow !== null && i > sourceAppStartRow + 1) {
        groups.sourceApps.push({
          start: sourceAppStartRow + 1,
          count: i - sourceAppStartRow - 1,
          app: currentApp
        });
      }
      
      sourceAppStartRow = i;
    }
  }
  
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

function expandAllGroups(sheet) {
  try {
    const maxRows = sheet.getMaxRows();
    for (let i = 0; i < 5; i++) {
      try {
        sheet.getRange(1, 1, maxRows, 1).expandGroups();
      } catch (e) {
        break;
      }
    }
  } catch (e) {
    console.log('Error expanding groups:', e);
  }
}

function testCommentCacheFixed() {
  console.log('=== TESTING FIXED COMMENT CACHE ===');
  
  const projects = ['TRICKY', 'MOLOCO', 'REGULAR'];
  
  projects.forEach(projectName => {
    console.log(`Testing ${projectName}...`);
    const cache = new CommentCache(projectName);
    
    cache.syncCommentsFromSheet();
    const comments = cache.loadAllComments();
    
    console.log(`${projectName}: ${Object.keys(comments).length} comments loaded`);
    
    if (Object.keys(comments).length > 0) {
      cache.applyCommentsToSheet();
      console.log(`${projectName}: Comments applied successfully`);
    }
  });
}