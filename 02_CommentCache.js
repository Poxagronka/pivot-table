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
    if (this.isTricky) {
      return this.loadAllCommentsTrickyOptimized();
    }
    return this.loadAllCommentsStandard();
  }

  loadAllCommentsTrickyOptimized() {
    console.log('Loading TRICKY comments optimized...');
    const comments = {};
    const data = this.cacheSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const [appName, weekRange, level, identifier, sourceApp, comment, lastUpdated] = data[i];
      if (comment) {
        const key = this.getCommentKey(appName, weekRange, level, identifier, sourceApp);
        comments[key] = comment;
      }
    }
    
    console.log(`TRICKY comments loaded: ${Object.keys(comments).length} entries`);
    return comments;
  }

  loadAllCommentsStandard() {
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

  saveComment(appName, weekRange, level, comment, identifier = null, sourceApp = null) {
    if (!comment || !comment.trim()) return;
    
    const data = this.cacheSheet.getDataRange().getValues();
    let found = false;
    
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

  syncCommentsFromSheet() {
    if (this.isTricky) {
      this.syncCommentsFromSheetTrickyOptimized();
    } else {
      this.syncCommentsFromSheetStandard();
    }
  }

  syncCommentsFromSheetTrickyOptimized() {
    console.log('Syncing TRICKY comments optimized...');
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(this.config.SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return;
    
    const data = sheet.getDataRange().getValues();
    const commentsToSave = [];
    let currentApp = '';
    let currentWeek = '';
    
    for (let i = 1; i < data.length; i++) {
      const level = data[i][0];
      const nameOrRange = data[i][1];
      const idOrEmpty = data[i][2];
      const comment = data[i][19];
      
      if (level === 'APP') {
        currentApp = nameOrRange;
        currentWeek = '';
      } else if (level === 'WEEK' && currentApp) {
        currentWeek = nameOrRange;
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
        if (comment) {
          const sourceAppDisplayName = nameOrRange;
          commentsToSave.push({
            appName: currentApp,
            weekRange: currentWeek,
            level: 'SOURCE_APP',
            comment: comment,
            identifier: sourceAppDisplayName,
            sourceApp: null
          });
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek && comment) {
        const sourceAppName = nameOrRange;
        const campaignIdValue = idOrEmpty && typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK') 
          ? this.extractCampaignIdFromHyperlink(idOrEmpty) 
          : idOrEmpty;
        
        commentsToSave.push({
          appName: currentApp,
          weekRange: currentWeek,
          level: 'CAMPAIGN',
          comment: comment,
          identifier: campaignIdValue,
          sourceApp: sourceAppName
        });
      }
    }
    
    console.log(`TRICKY optimized: Found ${commentsToSave.length} comments to save`);
    
    commentsToSave.forEach(commentData => {
      this.saveComment(
        commentData.appName,
        commentData.weekRange,
        commentData.level,
        commentData.comment,
        commentData.identifier,
        commentData.sourceApp
      );
    });
    
    console.log('TRICKY comments sync completed');
  }

  syncCommentsFromSheetStandard() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(this.config.SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return;
    
    const data = sheet.getDataRange().getValues();
    let currentApp = '';
    let currentWeek = '';
    
    for (let i = 1; i < data.length; i++) {
      const level = data[i][0];
      const nameOrRange = data[i][1];
      const idOrEmpty = data[i][2];
      const comment = data[i][19];
      
      if (level === 'APP') {
        currentApp = nameOrRange;
        currentWeek = '';
      } else if (level === 'WEEK' && currentApp) {
        currentWeek = nameOrRange;
        if (comment) {
          this.saveComment(currentApp, currentWeek, 'WEEK', comment);
        }
      } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
        if (comment) {
          const sourceAppDisplayName = nameOrRange;
          this.saveComment(currentApp, currentWeek, 'SOURCE_APP', comment, sourceAppDisplayName);
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek && comment) {
        const sourceAppName = nameOrRange;
        const campaignIdValue = idOrEmpty && typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK') 
          ? this.extractCampaignIdFromHyperlink(idOrEmpty) 
          : idOrEmpty;
        
        this.saveComment(currentApp, currentWeek, 'CAMPAIGN', comment, campaignIdValue, sourceAppName);
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
    try {
      if (this.isTricky) {
        this.applyCommentsToSheetTrickyOptimized();
      } else {
        this.applyCommentsToSheetStandard();
      }
    } catch (e) {
      console.log('Error applying comments to sheet:', e);
      console.log('Sheet may have been recreated, skipping comment application');
    }
  }

  applyCommentsToSheetTrickyOptimized() {
    console.log('Applying TRICKY comments optimized...');
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(this.config.SHEET_NAME);
    if (!sheet) {
      console.log('Sheet not found, cannot apply comments');
      return;
    }
    if (sheet.getLastRow() < 2) {
      console.log('Sheet has no data, cannot apply comments');
      return;
    }
    
    const comments = this.loadAllComments();
    const data = sheet.getDataRange().getValues();
    const commentUpdates = [];
    let currentApp = '';
    let currentWeek = '';
    
    for (let i = 1; i < data.length; i++) {
      const level = data[i][0];
      const nameOrRange = data[i][1];
      const idOrEmpty = data[i][2];
      
      if (level === 'APP') {
        currentApp = nameOrRange;
        currentWeek = '';
      } else if (level === 'WEEK' && currentApp) {
        currentWeek = nameOrRange;
        const weekKey = this.getCommentKey(currentApp, currentWeek, 'WEEK');
        const weekComment = comments[weekKey];
        if (weekComment) {
          commentUpdates.push({ row: i + 1, comment: weekComment });
        }
      } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
        const sourceAppDisplayName = nameOrRange;
        const sourceAppKey = this.getCommentKey(currentApp, currentWeek, 'SOURCE_APP', sourceAppDisplayName);
        const sourceAppComment = comments[sourceAppKey];
        if (sourceAppComment) {
          commentUpdates.push({ row: i + 1, comment: sourceAppComment });
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek) {
        const sourceAppName = nameOrRange;
        const campaignIdValue = idOrEmpty && typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK') 
          ? this.extractCampaignIdFromHyperlink(idOrEmpty) 
          : idOrEmpty;
        
        const campaignKey = this.getCommentKey(currentApp, currentWeek, 'CAMPAIGN', campaignIdValue, sourceAppName);
        const campaignComment = comments[campaignKey];
        if (campaignComment) {
          commentUpdates.push({ row: i + 1, comment: campaignComment });
        }
      }
    }
    
    console.log(`TRICKY optimized: Applying ${commentUpdates.length} comments`);
    
    if (commentUpdates.length > 0) {
      const batchSize = 100;
      for (let i = 0; i < commentUpdates.length; i += batchSize) {
        const batch = commentUpdates.slice(i, i + batchSize);
        batch.forEach(update => {
          sheet.getRange(update.row, 19).setValue(update.comment);
        });
        
        if (i + batchSize < commentUpdates.length) {
          Utilities.sleep(100);
        }
      }
    }
    
    console.log('TRICKY comments application completed');
  }

  applyCommentsToSheetStandard() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(this.config.SHEET_NAME);
    if (!sheet) {
      console.log('Sheet not found, cannot apply comments');
      return;
    }
    if (sheet.getLastRow() < 2) {
      console.log('Sheet has no data, cannot apply comments');
      return;
    }
    
    const comments = this.loadAllComments();
    const data = sheet.getDataRange().getValues();
    let currentApp = '';
    let currentWeek = '';
    
    for (let i = 1; i < data.length; i++) {
      const level = data[i][0];
      const nameOrRange = data[i][1];
      const idOrEmpty = data[i][2];
      
      if (level === 'APP') {
        currentApp = nameOrRange;
        currentWeek = '';
      } else if (level === 'WEEK' && currentApp) {
        currentWeek = nameOrRange;
        const weekKey = this.getCommentKey(currentApp, currentWeek, 'WEEK');
        const weekComment = comments[weekKey];
        if (weekComment) {
          sheet.getRange(i + 1, 19).setValue(weekComment);
        }
      } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
        const sourceAppDisplayName = nameOrRange;
        const sourceAppKey = this.getCommentKey(currentApp, currentWeek, 'SOURCE_APP', sourceAppDisplayName);
        const sourceAppComment = comments[sourceAppKey];
        if (sourceAppComment) {
          sheet.getRange(i + 1, 19).setValue(sourceAppComment);
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek) {
        const sourceAppName = nameOrRange;
        const campaignIdValue = idOrEmpty && typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK') 
          ? this.extractCampaignIdFromHyperlink(idOrEmpty) 
          : idOrEmpty;
        
        const campaignKey = this.getCommentKey(currentApp, currentWeek, 'CAMPAIGN', campaignIdValue, sourceAppName);
        const campaignComment = comments[campaignKey];
        if (campaignComment) {
          sheet.getRange(i + 1, 19).setValue(campaignComment);
        }
      }
    }
  }

  syncCommentsFromSheetQuiet() {
    this.syncCommentsFromSheet();
  }
}

function collapseAllGroupsRecursively(sheet) {
  console.log('Starting recursive collapse (one-by-one method)...');
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    console.log('No data to process');
    return;
  }
  
  const groups = identifyGroups(data);
  
  console.log(`Found ${groups.apps.length} app groups, ${groups.weeks.length} week groups, ${groups.sourceApps.length} source app groups`);
  
  let totalCollapsed = 0;
  
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
    
    console.log(`Groups expanded after ${attempts} attempts`);
  } catch (e) {
    console.log('Error expanding groups:', e);
  }
}