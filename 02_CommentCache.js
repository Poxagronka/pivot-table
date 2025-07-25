var COMMENT_CACHE_GLOBAL = {
  spreadsheetMetadata: null,
  spreadsheetMetadataTime: null,
  sheetData: {},
  sheetDataTime: {},
  CACHE_DURATION: 300000
};

class CommentCache {
  constructor(projectName = null) {
    this.projectName = projectName ? projectName.toUpperCase() : CURRENT_PROJECT; 
    this.config = projectName ? getProjectConfig(this.projectName) : getCurrentConfig();
    this.cacheSpreadsheetId = COMMENTS_CACHE_SPREADSHEET_ID;
    this.cacheSheetName = null;
    this.cacheSheetId = null;
    this._columnCache = {};
    this._headersCache = {};
  }

  getCachedSpreadsheetMetadata() {
    const now = Date.now();
    
    if (COMMENT_CACHE_GLOBAL.spreadsheetMetadata && 
        COMMENT_CACHE_GLOBAL.spreadsheetMetadataTime &&
        (now - COMMENT_CACHE_GLOBAL.spreadsheetMetadataTime) < COMMENT_CACHE_GLOBAL.CACHE_DURATION) {
      return COMMENT_CACHE_GLOBAL.spreadsheetMetadata;
    }
    
    const metadata = Sheets.Spreadsheets.get(this.cacheSpreadsheetId);
    
    COMMENT_CACHE_GLOBAL.spreadsheetMetadata = metadata;
    COMMENT_CACHE_GLOBAL.spreadsheetMetadataTime = now;
    
    return metadata;
  }

  getCachedSheetData(spreadsheetId, range) {
    const cacheKey = `${spreadsheetId}_${range}`;
    const now = Date.now();
    
    if (COMMENT_CACHE_GLOBAL.sheetData[cacheKey] && 
        COMMENT_CACHE_GLOBAL.sheetDataTime[cacheKey] &&
        (now - COMMENT_CACHE_GLOBAL.sheetDataTime[cacheKey]) < COMMENT_CACHE_GLOBAL.CACHE_DURATION) {
      return COMMENT_CACHE_GLOBAL.sheetData[cacheKey];
    }
    
    const data = Sheets.Spreadsheets.Values.get(spreadsheetId, range);
    
    COMMENT_CACHE_GLOBAL.sheetData[cacheKey] = data;
    COMMENT_CACHE_GLOBAL.sheetDataTime[cacheKey] = now;
    
    return data;
  }

  clearCache() {
    COMMENT_CACHE_GLOBAL.spreadsheetMetadata = null;
    COMMENT_CACHE_GLOBAL.spreadsheetMetadataTime = null;
    COMMENT_CACHE_GLOBAL.sheetData = {};
    COMMENT_CACHE_GLOBAL.sheetDataTime = {};
  }

  getOrCreateCacheSheet() {
    if (this.cacheSheetId && this.cacheSheetName) {
      return { name: this.cacheSheetName, id: this.cacheSheetId };
    }
    
    const sheetName = this.config.COMMENTS_CACHE_SHEET || `CommentsCache_${this.projectName}`;
    
    try {
      const spreadsheet = this.getCachedSpreadsheetMetadata();
      let sheet = spreadsheet.sheets.find(s => s.properties.title === sheetName);
      
      if (!sheet) {
        const addSheetRequest = {
          requests: [{
            addSheet: {
              properties: {
                title: sheetName
              }
            }
          }]
        };
        
        const response = Sheets.Spreadsheets.batchUpdate(addSheetRequest, this.cacheSpreadsheetId);
        const newSheet = response.replies[0].addSheet;
        
        const batchRequests = {
          requests: [
            {
              updateCells: {
                range: {
                  sheetId: newSheet.properties.sheetId,
                  startRowIndex: 0,
                  endRowIndex: 1,
                  startColumnIndex: 0,
                  endColumnIndex: 8
                },
                rows: [{
                  values: [
                    { userEnteredValue: { stringValue: 'AppName' } },
                    { userEnteredValue: { stringValue: 'WeekRange' } },
                    { userEnteredValue: { stringValue: 'Level' } },
                    { userEnteredValue: { stringValue: 'Identifier' } },
                    { userEnteredValue: { stringValue: 'SourceApp' } },
                    { userEnteredValue: { stringValue: 'Campaign' } },
                    { userEnteredValue: { stringValue: 'Comment' } },
                    { userEnteredValue: { stringValue: 'LastUpdated' } }
                  ]
                }],
                fields: 'userEnteredValue'
              }
            },
            {
              repeatCell: {
                range: {
                  sheetId: newSheet.properties.sheetId,
                  startRowIndex: 0,
                  endRowIndex: 1,
                  startColumnIndex: 0,
                  endColumnIndex: 8
                },
                cell: {
                  userEnteredFormat: {
                    textFormat: { bold: true },
                    backgroundColor: { red: 0.94, green: 0.94, blue: 0.94 }
                  }
                },
                fields: 'userEnteredFormat(textFormat,backgroundColor)'
              }
            }
          ]
        };
        
        Sheets.Spreadsheets.batchUpdate(batchRequests, this.cacheSpreadsheetId);
        
        this.cacheSheetName = sheetName;
        this.cacheSheetId = newSheet.properties.sheetId;
        
        COMMENT_CACHE_GLOBAL.spreadsheetMetadata = null;
        COMMENT_CACHE_GLOBAL.spreadsheetMetadataTime = null;
      } else {
        this.cacheSheetName = sheet.properties.title;
        this.cacheSheetId = sheet.properties.sheetId;
      }
      
      return { name: this.cacheSheetName, id: this.cacheSheetId };
    } catch (e) {
      console.error('Error creating/accessing cache sheet:', e);
      throw e;
    }
  }

  findColumnByHeader(headers, headerText) {
    for (let i = 0; i < headers.length; i++) {
      if (headers[i].toString().toLowerCase().trim() === headerText.toLowerCase().trim()) {
        return i + 1;
      }
    }
    return -1;
  }

  getSheetHeaders(sheetName) {
    const cacheKey = `${this.config.SHEET_ID}_${sheetName}_headers`;
    
    if (this._headersCache[cacheKey]) {
      return this._headersCache[cacheKey];
    }
    
    try {
      const range = `${sheetName}!1:1`;
      const response = this.getCachedSheetData(this.config.SHEET_ID, range);
      const headers = response.values ? response.values[0] : [];
      
      this._headersCache[cacheKey] = headers;
      return headers;
    } catch (e) {
      console.error('Error getting sheet headers:', e);
      return [];
    }
  }

  getCommentColumn(sheetName) {
  const headers = this.getSheetHeaders(sheetName);
  let column = this.findColumnByHeader(headers, 'Comments');
  if (column === -1) {
    column = this.findColumnByHeader(headers, 'Comment');
  }
  if (column === -1) {
    console.log(`Column 'Comments' not found in sheet ${sheetName} - sheet may be empty`);
    return null; // вместо throw new Error
  }
  return column;
}
  
  getLevelColumn(sheetName) {
    const headers = this.getSheetHeaders(sheetName);
    const column = this.findColumnByHeader(headers, 'Level');
    return column === -1 ? 1 : column;
  }
  
  getNameColumn(sheetName) {
    const headers = this.getSheetHeaders(sheetName);
    let column = this.findColumnByHeader(headers, 'Week Range / Source App');
    if (column === -1) {
      column = this.findColumnByHeader(headers, 'Week Range/Source App');
    }
    return column === -1 ? 2 : column;
  }
  
  getIdColumn(sheetName) {
    const headers = this.getSheetHeaders(sheetName);
    const column = this.findColumnByHeader(headers, 'ID');
    return column === -1 ? 3 : column;
  }

  getCommentKey(appName, weekRange, level, identifier = null, sourceApp = null, campaign = null) {
    return `${appName}|||${weekRange}|||${level}|||${identifier || 'N/A'}|||${sourceApp || 'N/A'}|||${campaign || 'N/A'}`;
  }

  loadAllComments() {
    this.getOrCreateCacheSheet();
    const comments = {};
    
    try {
      const range = `${this.cacheSheetName}!A:H`;
      const response = this.getCachedSheetData(this.cacheSpreadsheetId, range);
      
      if (!response.values || response.values.length <= 1) {
        return comments;
      }
      
      for (let i = 1; i < response.values.length; i++) {
        const row = response.values[i];
        if (row.length >= 7) {
          const [appName, weekRange, level, identifier, sourceApp, campaign, comment] = row;
          if (comment) {
            const key = this.getCommentKey(appName, weekRange, level, identifier, sourceApp, campaign);
            comments[key] = comment;
          }
        }
      }
    } catch (e) {
      console.error('Error loading comments from cache:', e);
    }
    
    return comments;
  }

  batchSaveComments(commentsArray) {
    if (!commentsArray || commentsArray.length === 0) return;
    
    this.getOrCreateCacheSheet();
    
    try {
      const range = `${this.cacheSheetName}!A:H`;
      const response = this.getCachedSheetData(this.cacheSpreadsheetId, range);
      const existingData = response.values || [];
      
      const existingDataMap = {};
      for (let i = 1; i < existingData.length; i++) {
        if (existingData[i].length >= 6) {
          const [appName, weekRange, level, identifier, sourceApp, campaign] = existingData[i];
          const key = this.getCommentKey(appName, weekRange, level, identifier, sourceApp, campaign);
          existingDataMap[key] = {
            rowIndex: i + 1,
            existingComment: existingData[i][6] || ''
          };
        }
      }
      
      const updateRequests = [];
      const newRowsToAdd = [];
      const timestamp = new Date().toISOString();
      
      commentsArray.forEach(commentData => {
        const { appName, weekRange, level, comment, identifier, sourceApp, campaign } = commentData;
        const key = this.getCommentKey(appName, weekRange, level, identifier, sourceApp, campaign);
        
        if (existingDataMap[key]) {
          const existing = existingDataMap[key];
          if (comment.length > existing.existingComment.length) {
            updateRequests.push({
              range: `${this.cacheSheetName}!G${existing.rowIndex}:H${existing.rowIndex}`,
              values: [[comment, timestamp]]
            });
          }
        } else {
          newRowsToAdd.push([
            appName,
            weekRange,
            level,
            identifier || 'N/A',
            sourceApp || 'N/A',
            campaign || 'N/A',
            comment,
            timestamp
          ]);
        }
      });
      
      const allRequests = [];
      
      if (updateRequests.length > 0) {
        allRequests.push(...updateRequests);
      }
      
      if (newRowsToAdd.length > 0) {
        const startRow = existingData.length + 1;
        const endRow = startRow + newRowsToAdd.length - 1;
        allRequests.push({
          range: `${this.cacheSheetName}!A${startRow}:H${endRow}`,
          values: newRowsToAdd
        });
      }
      
      if (allRequests.length > 0) {
        const batchUpdateRequest = {
          valueInputOption: 'RAW',
          data: allRequests
        };
        
        Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest, this.cacheSpreadsheetId);
        
        const cacheKey = `${this.cacheSpreadsheetId}_${this.cacheSheetName}!A:H`;
        delete COMMENT_CACHE_GLOBAL.sheetData[cacheKey];
        delete COMMENT_CACHE_GLOBAL.sheetDataTime[cacheKey];
      }
      
      console.log(`Batch saved: ${updateRequests.length} updates, ${newRowsToAdd.length} new comments for ${this.projectName}`);
      
    } catch (e) {
      console.error('Error in batch save comments:', e);
      throw e;
    }
  }

  saveComment(appName, weekRange, level, comment, identifier = null, sourceApp = null, campaign = null) {
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
    
    this.batchSaveComments([{
      appName,
      weekRange, 
      level,
      comment: commentStr,
      identifier,
      sourceApp,
      campaign
    }]);
  }

  syncCommentsFromSheet() {
  const sheetName = this.config.SHEET_NAME;
  
  try {
    const range = `${sheetName}!A:Z`;
    const response = this.getCachedSheetData(this.config.SHEET_ID, range);
    
    if (!response.values || response.values.length < 2) {
      console.log(`${this.projectName}: No data found in ${sheetName} - skipping comment sync`);
      return;
    }
    
    const data = response.values;
    const headers = data[0];
    
    // Добавить проверку на существование заголовков
    if (!headers || headers.length === 0) {
      console.log(`${this.projectName}: No headers found in ${sheetName} - skipping comment sync`);
      return;
    }
    
    const levelCol = this.findColumnByHeader(headers, 'Level') - 1;
    const nameCol = this.findColumnByHeader(headers, 'Week Range / Source App') - 1;
    const idCol = this.findColumnByHeader(headers, 'ID') - 1;
    let commentCol = this.findColumnByHeader(headers, 'Comments') - 1;
    
    if (commentCol === -2) {
      commentCol = this.findColumnByHeader(headers, 'Comment') - 1;
      if (commentCol === -2) {
        console.log(`${this.projectName}: Comments column not found in ${sheetName} - skipping comment sync`);
        return; // вместо throw new Error
      }
    }
      
      const commentsToSave = [];
      let currentApp = '';
      let currentWeek = '';
      let weekComments = 0;
      let sourceAppComments = 0;
      let campaignComments = 0;
      let networkComments = 0;
      
      for (let i = 1; i < data.length; i++) {
        try {
          const row = data[i];
          if (!row || row.length === 0) continue;
          
          const level = row[levelCol] || '';
          const nameOrRange = row[nameCol] || '';
          const idOrEmpty = row[idCol] || '';
          const comment = row[commentCol] || '';
          
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
                identifier: 'N/A',
                sourceApp: 'N/A',
                campaign: 'N/A'
              });
              weekComments++;
            }
          } else if (level === 'NETWORK' && currentApp && currentWeek) {
            if (comment) {
              commentsToSave.push({
                appName: currentApp,
                weekRange: currentWeek,
                level: 'NETWORK',
                comment: comment,
                identifier: idOrEmpty || 'N/A',
                sourceApp: 'N/A',
                campaign: nameOrRange || 'N/A'
              });
              networkComments++;
            }
          } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
            if (comment) {
              commentsToSave.push({
                appName: currentApp,
                weekRange: currentWeek,
                level: 'SOURCE_APP',
                comment: comment,
                identifier: nameOrRange || 'N/A',
                sourceApp: nameOrRange || 'N/A',
                campaign: 'N/A'
              });
              sourceAppComments++;
            }
          } else if (level === 'CAMPAIGN' && currentApp && currentWeek && comment) {
            const sourceAppName = nameOrRange;
            const campaignIdValue = idOrEmpty && typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK') 
              ? this.extractCampaignIdFromHyperlink(idOrEmpty) 
              : idOrEmpty;
            
            let campaignName = '';
            if (this.projectName === 'TRICKY') {
              campaignName = this.extractCampaignNameFromHyperlink(campaignIdValue) || campaignIdValue || 'Unknown';
            } else {
              campaignName = sourceAppName || 'Unknown';
            }
            
            commentsToSave.push({
              appName: currentApp,
              weekRange: currentWeek,
              level: 'CAMPAIGN',
              comment: comment,
              identifier: this.projectName === 'TRICKY' ? campaignIdValue : 'N/A',
              sourceApp: sourceAppName || 'N/A',
              campaign: campaignName
            });
            campaignComments++;
          }
        } catch (e) {
          console.error(`Error processing row ${i + 1} in ${this.projectName}:`, e);
        }
      }
      
      if (commentsToSave.length > 0) {
        this.batchSaveComments(commentsToSave);
        console.log(`${this.projectName}: Synced ${commentsToSave.length} comments (${weekComments} weeks, ${sourceAppComments} source apps, ${campaignComments} campaigns, ${networkComments} networks)`);
      } else {
        console.log(`${this.projectName}: No comments to sync`);
      }
    } catch (e) {
      console.error(`Error syncing comments for ${this.projectName}:`, e);
      throw e;
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

  extractCampaignNameFromHyperlink(campaignIdOrFormula) {
    try {
      if (typeof campaignIdOrFormula === 'string' && campaignIdOrFormula.includes('HYPERLINK')) {
        const nameMatch = campaignIdOrFormula.match(/"([^"]+)"\s*\)$/);
        return nameMatch ? nameMatch[1] : null;
      }
      return null;
    } catch (e) {
      return null;
    }
  }

  applyCommentsToSheet() {
  const sheetName = this.config.SHEET_NAME;
  
  try {
    const range = `${sheetName}!A:Z`;
    const response = this.getCachedSheetData(this.config.SHEET_ID, range);
    
    if (!response.values || response.values.length < 2) {
      console.log(`${this.projectName}: No data found in ${sheetName} - skipping comment application`);
      return;
    }
    
    const data = response.values;
    const headers = data[0];
    
    // Добавить проверку на существование заголовков
    if (!headers || headers.length === 0) {
      console.log(`${this.projectName}: No headers found in ${sheetName} - skipping comment application`);
      return;
    }
    
    const levelCol = this.findColumnByHeader(headers, 'Level') - 1;
    const nameCol = this.findColumnByHeader(headers, 'Week Range / Source App') - 1;
    const idCol = this.findColumnByHeader(headers, 'ID') - 1;
    let commentCol = this.findColumnByHeader(headers, 'Comments');
    
    if (commentCol === -1) {
      commentCol = this.findColumnByHeader(headers, 'Comment');
    }
    
    if (commentCol === -1) {
      console.log(`${this.projectName}: Comments column not found in ${sheetName} - skipping comment application`);
      return; // вместо throw new Error
    }
      
      const comments = this.loadAllComments();
      
      let currentApp = '';
      let currentWeek = '';
      const updatesToMake = [];
      let weekComments = 0;
      let sourceAppComments = 0;
      let campaignComments = 0;
      let networkComments = 0;
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;
        
        const level = row[levelCol] || '';
        const nameOrRange = row[nameCol] || '';
        const idOrEmpty = row[idCol] || '';
        
        if (level === 'APP') {
          currentApp = nameOrRange;
          currentWeek = '';
        } else if (level === 'WEEK' && currentApp) {
          currentWeek = nameOrRange;
          const weekKey = this.getCommentKey(currentApp, currentWeek, 'WEEK', 'N/A', 'N/A', 'N/A');
          const weekComment = comments[weekKey];
          if (weekComment) {
            const cellRange = `${sheetName}!${this.columnNumberToLetter(commentCol)}${i + 1}`;
            updatesToMake.push({
              range: cellRange,
              values: [[weekComment]]
            });
            weekComments++;
          }
        } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
          const sourceAppDisplayName = nameOrRange;
          const sourceAppKey = this.getCommentKey(currentApp, currentWeek, 'SOURCE_APP', sourceAppDisplayName, sourceAppDisplayName, 'N/A');
          const sourceAppComment = comments[sourceAppKey];
          if (sourceAppComment) {
            const cellRange = `${sheetName}!${this.columnNumberToLetter(commentCol)}${i + 1}`;
            updatesToMake.push({
              range: cellRange,
              values: [[sourceAppComment]]
            });
            sourceAppComments++;
          }
        } else if (level === 'CAMPAIGN' && currentApp && currentWeek) {
          const sourceAppName = nameOrRange;
          let campaignIdValue = idOrEmpty;
          
          if (typeof campaignIdValue === 'string' && campaignIdValue.includes('HYPERLINK')) {
            campaignIdValue = this.extractCampaignIdFromHyperlink(campaignIdValue);
          }
          
          let campaignName = '';
          if (this.projectName === 'TRICKY') {
            campaignName = this.extractCampaignNameFromHyperlink(idOrEmpty) || campaignIdValue || 'Unknown';
          } else {
            campaignName = sourceAppName || 'Unknown';
          }
          
          const campaignKey = this.getCommentKey(currentApp, currentWeek, 'CAMPAIGN', 
            this.projectName === 'TRICKY' ? campaignIdValue : 'N/A', 
            sourceAppName, 
            campaignName);
          const campaignComment = comments[campaignKey];
          if (campaignComment) {
            const cellRange = `${sheetName}!${this.columnNumberToLetter(commentCol)}${i + 1}`;
            updatesToMake.push({
              range: cellRange,
              values: [[campaignComment]]
            });
            campaignComments++;
          }
        } else if (level === 'NETWORK' && currentApp && currentWeek) {
          const networkName = nameOrRange;
          const networkKey = this.getCommentKey(currentApp, currentWeek, 'NETWORK', idOrEmpty || 'N/A', 'N/A', networkName);
          const networkComment = comments[networkKey];
          if (networkComment) {
            const cellRange = `${sheetName}!${this.columnNumberToLetter(commentCol)}${i + 1}`;
            updatesToMake.push({
              range: cellRange,
              values: [[networkComment]]
            });
            networkComments++;
          }
        }
      }
      
      if (updatesToMake.length > 0) {
        const batchUpdateRequest = {
          valueInputOption: 'RAW',
          data: updatesToMake
        };
        
        Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest, this.config.SHEET_ID);
        
        const cacheKey = `${this.config.SHEET_ID}_${range}`;
        delete COMMENT_CACHE_GLOBAL.sheetData[cacheKey];
        delete COMMENT_CACHE_GLOBAL.sheetDataTime[cacheKey];
        
        console.log(`${this.projectName}: Applied ${updatesToMake.length} comments (${weekComments} weeks, ${sourceAppComments} source apps, ${campaignComments} campaigns, ${networkComments} networks)`);
      } else {
        console.log(`${this.projectName}: No comments found to apply`);
      }
    } catch (e) {
      console.error(`Error applying comments to ${sheetName}:`, e);
      throw e;
    }
  }

  columnNumberToLetter(column) {
    let temp, letter = '';
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
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