class CommentCache {
  constructor(projectName = null) {
    this.projectName = projectName || CURRENT_PROJECT;
    this.config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
    this.cacheSheet = this.getOrCreateCacheSheet();
    this.isTricky = this.projectName === 'TRICKY';
  }

  getOrCreateCacheSheet() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    let sheet = spreadsheet.getSheetByName(this.config.COMMENTS_CACHE_SHEET);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(this.config.COMMENTS_CACHE_SHEET);
      sheet.hideSheet();
      sheet.getRange(1, 1, 1, 7).setValues([['AppName', 'WeekRange', 'Level', 'Identifier', 'SourceApp', 'Comment', 'LastUpdated']]);
    }
    return sheet;
  }

  getCommentKey(appName, weekRange, level, identifier = null, sourceApp = null) {
    return `${appName}|||${weekRange}|||${level}|||${identifier || 'N/A'}|||${sourceApp || 'N/A'}`;
  }

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

  syncCommentsFromSheet() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(this.config.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    const commentsToSave = [];
    let currentApp = '';
    let currentWeek = '';
    let currentSourceApp = '';
    
    for (let i = 1; i < data.length; i++) {
      const level = data[i][0];
      const nameOrRange = data[i][1];
      const idOrEmpty = data[i][2];
      const comment = data[i][18];
      
      if (level === 'APP') {
        currentApp = nameOrRange;
        currentWeek = '';
        currentSourceApp = '';
      } else if (level === 'WEEK' && currentApp) {
        currentWeek = nameOrRange;
        currentSourceApp = '';
        if (comment) {
          commentsToSave.push({
            appName: currentApp,
            weekRange: currentWeek,
            level: 'WEEK',
            comment: comment,
            identifier: null,
            sourceApp: null
          });
        }
      } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
        currentSourceApp = nameOrRange;
        if (comment) {
          commentsToSave.push({
            appName: currentApp,
            weekRange: currentWeek,
            level: 'SOURCE_APP',
            comment: comment,
            identifier: currentSourceApp,
            sourceApp: null
          });
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek && currentSourceApp) {
        if (comment) {
          let campaignIdValue = idOrEmpty;
          if (typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK')) {
            campaignIdValue = this.extractCampaignIdFromHyperlink(idOrEmpty);
          }
          
          commentsToSave.push({
            appName: currentApp,
            weekRange: currentWeek,
            level: 'CAMPAIGN',
            comment: comment,
            identifier: campaignIdValue,
            sourceApp: currentSourceApp
          });
        }
      }
    }
    
    if (commentsToSave.length > 0) {
      this.batchSaveComments(commentsToSave);
    }
  }

  batchSaveComments(commentsToSave) {
    const existingData = this.cacheSheet.getDataRange().getValues();
    const existingMap = new Map();
    
    for (let i = 1; i < existingData.length; i++) {
      const [appName, weekRange, level, identifier, sourceApp, comment] = existingData[i];
      const key = this.getCommentKey(appName, weekRange, level, identifier, sourceApp);
      existingMap.set(key, { 
        row: i + 1, 
        comment: comment || '',
        data: existingData[i]
      });
    }
    
    const updates = [];
    const newRows = [];
    const now = new Date();
    
    commentsToSave.forEach(commentData => {
      const key = this.getCommentKey(
        commentData.appName,
        commentData.weekRange,
        commentData.level,
        commentData.identifier,
        commentData.sourceApp
      );
      
      const existing = existingMap.get(key);
      
      if (existing) {
        if (commentData.comment.length > existing.comment.length) {
          updates.push({
            row: existing.row,
            comment: commentData.comment,
            lastUpdated: now
          });
        }
      } else {
        newRows.push([
          commentData.appName,
          commentData.weekRange,
          commentData.level,
          commentData.identifier || 'N/A',
          commentData.sourceApp || 'N/A',
          commentData.comment,
          now
        ]);
      }
    });
    
    if (updates.length > 0) {
      const batchSize = 100;
      for (let i = 0; i < updates.length; i += batchSize) {
        const batch = updates.slice(i, i + batchSize);
        batch.forEach(update => {
          this.cacheSheet.getRange(update.row, 6, 1, 2).setValues([[update.comment, update.lastUpdated]]);
        });
        if (i + batchSize < updates.length) {
          Utilities.sleep(50);
        }
      }
    }
    
    if (newRows.length > 0) {
      const batchSize = 100;
      for (let i = 0; i < newRows.length; i += batchSize) {
        const batch = newRows.slice(i, i + batchSize);
        const lastRow = this.cacheSheet.getLastRow();
        this.cacheSheet.getRange(lastRow + 1, 1, batch.length, 7).setValues(batch);
        if (i + batchSize < newRows.length) {
          Utilities.sleep(50);
        }
      }
    }
  }

  extractCampaignIdFromHyperlink(hyperlinkFormula) {
    try {
      const match = hyperlinkFormula.match(/campaigns\/([^"]+)/);
      return match ? match[1] : 'Unknown';
    } catch (e) {
      return 'Unknown';
    }
  }

  applyCommentsToSheet() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(this.config.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return;
    }
    
    const comments = this.loadAllComments();
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
        const weekKey = this.getCommentKey(currentApp, currentWeek, 'WEEK');
        const weekComment = comments[weekKey];
        if (weekComment) {
          commentUpdates.push({ row: i + 1, comment: weekComment });
        }
      } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
        currentSourceApp = nameOrRange;
        const sourceAppKey = this.getCommentKey(currentApp, currentWeek, 'SOURCE_APP', currentSourceApp);
        const sourceAppComment = comments[sourceAppKey];
        if (sourceAppComment) {
          commentUpdates.push({ row: i + 1, comment: sourceAppComment });
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek && currentSourceApp) {
        let campaignIdValue = idOrEmpty;
        if (typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK')) {
          campaignIdValue = this.extractCampaignIdFromHyperlink(idOrEmpty);
        }
        
        const campaignKey = this.getCommentKey(currentApp, currentWeek, 'CAMPAIGN', campaignIdValue, currentSourceApp);
        const campaignComment = comments[campaignKey];
        if (campaignComment) {
          commentUpdates.push({ row: i + 1, comment: campaignComment });
        }
      }
    }
    
    if (commentUpdates.length > 0) {
      const batchSize = 100;
      for (let i = 0; i < commentUpdates.length; i += batchSize) {
        const batch = commentUpdates.slice(i, i + batchSize);
        batch.forEach(update => {
          sheet.getRange(update.row, 19).setValue(update.comment);
        });
        
        if (i + batchSize < commentUpdates.length) {
          Utilities.sleep(50);
        }
      }
    }
  }

  syncCommentsFromSheetQuiet() {
    this.syncCommentsFromSheet();
  }
}

function collapseAllGroupsRecursively(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return;
  }
  
  const groups = identifyGroups(data);
  
  let totalCollapsed = 0;
  
  if (groups.sourceApps.length > 0) {
    let sourceAppCollapsed = 0;
    groups.sourceApps.forEach((group, index) => {
      try {
        sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
        sourceAppCollapsed++;
        SpreadsheetApp.flush();
        Utilities.sleep(50);
      } catch (e) {
      }
    });
    totalCollapsed += sourceAppCollapsed;
  }
  
  let weekCollapsed = 0;
  groups.weeks.forEach((group, index) => {
    try {
      sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
      weekCollapsed++;
      SpreadsheetApp.flush();
      Utilities.sleep(50);
    } catch (e) {
    }
  });
  totalCollapsed += weekCollapsed;
  
  let appCollapsed = 0;
  groups.apps.forEach((group, index) => {
    try {
      sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
      appCollapsed++;
      SpreadsheetApp.flush();
      Utilities.sleep(50);
    } catch (e) {
    }
  });
  totalCollapsed += appCollapsed;
}

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
      
    } else if (level === 'CAMPAIGN') {
      continue;
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
        expanded = false;
      }
    }
  } catch (e) {
  }
}