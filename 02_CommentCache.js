class CommentCache {
  constructor(projectName = null) {
    this.projectName = projectName || CURRENT_PROJECT;
    this.config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
    this.cacheSheet = this.getOrCreateCacheSheet();
    this._columnCache = {};
    this._headersCache = {};
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

  findColumnByHeader(sheet, headerText, useCache = true) {
    const sheetId = sheet.getSheetId();
    const cacheKey = `${sheetId}_${headerText}`;
    
    if (useCache && this._columnCache[cacheKey]) {
      const cachedData = this._columnCache[cacheKey];
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const headersKey = currentHeaders.join('|||');
      
      if (cachedData.headersKey === headersKey) {
        return cachedData.column;
      }
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headersKey = headers.join('|||');
    
    for (let i = 0; i < headers.length; i++) {
      if (headers[i].toString().toLowerCase().trim() === headerText.toLowerCase().trim()) {
        this._columnCache[cacheKey] = {
          column: i + 1,
          headersKey: headersKey
        };
        return i + 1;
      }
    }
    
    return -1;
  }

  getCommentColumn(sheet) {
    let column = this.findColumnByHeader(sheet, 'Comments');
    if (column === -1) {
      column = this.findColumnByHeader(sheet, 'Comment');
    }
    if (column === -1) {
      console.error(`Column 'Comments' not found in sheet ${sheet.getName()}`);
      console.log('Available headers:', sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
      throw new Error(`Column 'Comments' not found in sheet ${sheet.getName()}`);
    }
    return column;
  }
  
  getLevelColumn(sheet) {
    const column = this.findColumnByHeader(sheet, 'Level');
    return column === -1 ? 1 : column;
  }
  
  getNameColumn(sheet) {
    let column = this.findColumnByHeader(sheet, 'Week Range / Source App');
    if (column === -1) {
      column = this.findColumnByHeader(sheet, 'Week Range/Source App');
    }
    return column === -1 ? 2 : column;
  }
  
  getIdColumn(sheet) {
    const column = this.findColumnByHeader(sheet, 'ID');
    return column === -1 ? 3 : column;
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

  saveComment(appName, weekRange, level, comment, identifier = null, sourceApp = null) {
    if (comment === null || comment === undefined || comment === '') return;
    
    let commentStr = comment;
    
    if (typeof comment !== 'string') {
      console.log(`WARNING: Non-string comment in ${this.projectName}`);
      console.log(`  Type: ${typeof comment}, Value: ${comment}`);
      console.log(`  Location: App="${appName}", Week="${weekRange}", Level="${level}"`);
      
      if (comment instanceof Date) {
        commentStr = Utilities.formatDate(comment, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      } else if (typeof comment === 'object') {
        commentStr = JSON.stringify(comment);
      } else {
        commentStr = String(comment);
      }
    }
    
    if (!commentStr.trim()) return;
    
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
        if (commentStr.length > existingComment.length) {
          this.cacheSheet.getRange(i + 1, 6, 1, 2).setValues([[commentStr, new Date()]]);
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
        commentStr, 
        new Date()
      ]]);
    }
  }

  syncCommentsFromSheet() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(this.config.SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return;
    
    console.log(`Starting sync for ${this.projectName}`);
    
    const levelCol = this.getLevelColumn(sheet) - 1;
    const nameCol = this.getNameColumn(sheet) - 1;
    const idCol = this.getIdColumn(sheet) - 1;
    const commentCol = this.getCommentColumn(sheet) - 1;
    
    console.log(`Column indices - Level: ${levelCol}, Name: ${nameCol}, ID: ${idCol}, Comment: ${commentCol}`);
    
    const data = sheet.getDataRange().getValues();
    let currentApp = '';
    let currentWeek = '';
    let savedCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      try {
        const level = data[i][levelCol];
        const nameOrRange = data[i][nameCol];
        const idOrEmpty = data[i][idCol];
        const comment = data[i][commentCol];
        
        if (level === 'APP') {
          currentApp = nameOrRange;
          currentWeek = '';
        } else if (level === 'WEEK' && currentApp) {
          currentWeek = nameOrRange;
          if (comment) {
            this.saveComment(currentApp, currentWeek, 'WEEK', comment);
            savedCount++;
          }
        } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
          if (comment) {
            const sourceAppDisplayName = nameOrRange;
            this.saveComment(currentApp, currentWeek, 'SOURCE_APP', comment, sourceAppDisplayName);
            savedCount++;
          }
        } else if (level === 'CAMPAIGN' && currentApp && currentWeek && comment) {
          const sourceAppName = nameOrRange;
          const campaignIdValue = idOrEmpty && typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK') 
            ? this.extractCampaignIdFromHyperlink(idOrEmpty) 
            : idOrEmpty;
          
          this.saveComment(currentApp, currentWeek, 'CAMPAIGN', comment, campaignIdValue, sourceAppName);
          savedCount++;
        }
      } catch (e) {
        console.error(`Error processing row ${i + 1} in ${this.projectName}:`, e);
        console.log(`  Row data: Level="${data[i][levelCol]}", Name="${data[i][nameCol]}", Comment type=${typeof data[i][commentCol]}`);
      }
    }
    
    console.log(`${this.projectName}: Sync completed, saved ${savedCount} comments`);
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
    if (!sheet || sheet.getLastRow() < 2) return;
    
    const levelCol = this.getLevelColumn(sheet) - 1;
    const nameCol = this.getNameColumn(sheet) - 1;
    const idCol = this.getIdColumn(sheet) - 1;
    const commentCol = this.getCommentColumn(sheet);
    
    const comments = this.loadAllComments();
    const data = sheet.getDataRange().getValues();
    let currentApp = '';
    let currentWeek = '';
    
    for (let i = 1; i < data.length; i++) {
      const level = data[i][levelCol];
      const nameOrRange = data[i][nameCol];
      const idOrEmpty = data[i][idCol];
      
      if (level === 'APP') {
        currentApp = nameOrRange;
        currentWeek = '';
      } else if (level === 'WEEK' && currentApp) {
        currentWeek = nameOrRange;
        const weekKey = this.getCommentKey(currentApp, currentWeek, 'WEEK');
        const weekComment = comments[weekKey];
        if (weekComment) {
          sheet.getRange(i + 1, commentCol).setValue(weekComment);
        }
      } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
        const sourceAppDisplayName = nameOrRange;
        const sourceAppKey = this.getCommentKey(currentApp, currentWeek, 'SOURCE_APP', sourceAppDisplayName);
        const sourceAppComment = comments[sourceAppKey];
        if (sourceAppComment) {
          sheet.getRange(i + 1, commentCol).setValue(sourceAppComment);
        }
      } else if (level === 'CAMPAIGN' && currentApp && currentWeek) {
        const sourceAppName = nameOrRange;
        const campaignIdValue = idOrEmpty && typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK') 
          ? this.extractCampaignIdFromHyperlink(idOrEmpty) 
          : idOrEmpty;
        
        const campaignKey = this.getCommentKey(currentApp, currentWeek, 'CAMPAIGN', campaignIdValue, sourceAppName);
        const campaignComment = comments[campaignKey];
        if (campaignComment) {
          sheet.getRange(i + 1, commentCol).setValue(campaignComment);
        }
      }
    }
  }

  syncCommentsFromSheetQuiet() {
    this.syncCommentsFromSheet();
  }

  clearColumnCache() {
    this._columnCache = {};
    this._headersCache = {};
    console.log(`Column cache cleared for ${this.projectName}`);
  }
}

function collapseAllGroupsRecursively(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  
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
      } catch (e) {}
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
    } catch (e) {}
  });
  totalCollapsed += weekCollapsed;
  
  let appCollapsed = 0;
  groups.apps.forEach((group, index) => {
    try {
      sheet.getRange(group.start, 1, group.count, 1).collapseGroups();
      appCollapsed++;
      SpreadsheetApp.flush();
      Utilities.sleep(50);
    } catch (e) {}
  });
  totalCollapsed += appCollapsed;
}

function identifyGroups(data) {
  const groups = {
    apps: [],
    weeks: [],
    sourceApps: []
  };
  
  const levelCol = 0;
  
  let currentApp = null;
  let appStartRow = null;
  let weekStartRow = null;
  let sourceAppStartRow = null;
  
  for (let i = 1; i < data.length; i++) {
    const level = data[i][levelCol];
    
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
  } catch (e) {}
}

function testCommentCacheSync(projectName = 'TRICKY') {
  console.log(`=== Testing Comment Cache Sync for ${projectName} ===`);
  try {
    const cache = new CommentCache(projectName);
    cache.syncCommentsFromSheet();
    console.log('✅ Sync completed successfully');
  } catch (e) {
    console.error('❌ Sync failed:', e);
    console.log('Stack trace:', e.stack);
  }
}

function diagnoseCommentColumn(projectName = 'TRICKY') {
  console.log(`=== Diagnosing Comment Column for ${projectName} ===`);
  
  try {
    const config = getProjectConfig(projectName);
    const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
    const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
    
    if (!sheet) {
      console.log('❌ Sheet not found:', config.SHEET_NAME);
      return;
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log('Headers found:', headers);
    
    const cache = new CommentCache(projectName);
    const commentCol = cache.getCommentColumn(sheet) - 1;
    console.log(`Comment column index: ${commentCol} (column ${commentCol + 1})`);
    
    const data = sheet.getDataRange().getValues();
    console.log(`Total rows: ${data.length}`);
    
    let typesFound = {};
    for (let i = 1; i < Math.min(data.length, 20); i++) {
      const comment = data[i][commentCol];
      const type = typeof comment;
      
      if (!typesFound[type]) typesFound[type] = 0;
      typesFound[type]++;
      
      if (comment && type !== 'string') {
        console.log(`Row ${i + 1}: Type=${type}, Value=${comment}, Level=${data[i][0]}`);
      }
    }
    
    console.log('Comment types found in first 20 rows:', typesFound);
    
  } catch (e) {
    console.error('Diagnosis failed:', e);
  }
}