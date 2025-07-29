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
                  endColumnIndex: 9
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
                    { userEnteredValue: { stringValue: 'LastUpdated' } },
                    { userEnteredValue: { stringValue: 'RowHash' } }
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
                  endColumnIndex: 9
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
        
        this.migrateToHashBasedCache();
      }
      
      return { name: this.cacheSheetName, id: this.cacheSheetId };
    } catch (e) {
      console.error('Error creating/accessing cache sheet:', e);
      throw e;
    }
  }

  migrateToHashBasedCache() {
    try {
      const range = `${this.cacheSheetName}!A:I`;
      const response = this.getCachedSheetData(this.cacheSpreadsheetId, range);
      
      if (!response.values || response.values.length <= 1) return;
      
      const headers = response.values[0];
      const hasHashColumn = headers.length >= 9 && headers[8] === 'RowHash';
      
      if (!hasHashColumn) {
        console.log(`${this.projectName}: Migrating cache to hash-based system...`);
        
        const updateRequests = [];
        updateRequests.push({
          range: `${this.cacheSheetName}!I1`,
          values: [['RowHash']]
        });
        
        for (let i = 1; i < response.values.length; i++) {
          const row = response.values[i];
          if (row.length >= 6) {
            const [appName, weekRange, level, identifier, sourceApp, campaign] = row;
            const hash = this.generateRowHash(level, appName, weekRange, identifier, sourceApp, campaign);
            updateRequests.push({
              range: `${this.cacheSheetName}!I${i + 1}`,
              values: [[hash]]
            });
          }
        }
        
        if (updateRequests.length > 0) {
          const batchUpdateRequest = {
            valueInputOption: 'RAW',
            data: updateRequests
          };
          
          Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest, this.cacheSpreadsheetId);
          console.log(`${this.projectName}: Migrated ${updateRequests.length - 1} comments to hash-based system`);
        }
      }
    } catch (e) {
      console.error('Error migrating cache:', e);
    }
  }

 generateRowHash(level, appName, weekRange, identifier = '', sourceApp = '', campaign = '') {
    // Используем единую функцию из Utilities
    if (level === 'WEEK') {
      return generateCommentHash(level, appName, weekRange, this.projectName);
    } else {
      return generateDetailedCommentHash(level, appName, weekRange, 
        identifier, sourceApp, campaign, this.projectName);
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
      console.error(`Column 'Comments' not found in sheet ${sheetName}`);
      console.log('Available headers:', headers);
      throw new Error(`Column 'Comments' not found in sheet ${sheetName}`);
    }
    return column;
  }
  
  getHashColumn(sheetName) {
    const headers = this.getSheetHeaders(sheetName);
    const column = this.findColumnByHeader(headers, 'RowHash');
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
    const commentsByHash = {};
    
    try {
      const range = `${this.cacheSheetName}!A:I`;
      const response = this.getCachedSheetData(this.cacheSpreadsheetId, range);
      
      if (!response.values || response.values.length <= 1) {
        return { comments, commentsByHash };
      }
      
      for (let i = 1; i < response.values.length; i++) {
        const row = response.values[i];
        if (row.length >= 7) {
          const [appName, weekRange, level, identifier, sourceApp, campaign, comment, lastUpdated, hash] = row;
          if (comment) {
            const key = this.getCommentKey(appName, weekRange, level, identifier, sourceApp, campaign);
            comments[key] = comment;
            
            if (hash) {
              commentsByHash[hash] = comment;
            }
          }
        }
      }
    } catch (e) {
      console.error('Error loading comments from cache:', e);
    }
    
    return { comments, commentsByHash };
  }

  batchSaveComments(commentsArray) {
    if (!commentsArray || commentsArray.length === 0) return;
    
    this.getOrCreateCacheSheet();
    
    try {
      const range = `${this.cacheSheetName}!A:I`;
      const response = this.getCachedSheetData(this.cacheSpreadsheetId, range);
      const existingData = response.values || [];
      
      const existingDataMap = {};
      const existingHashMap = {};
      
      for (let i = 1; i < existingData.length; i++) {
        if (existingData[i].length >= 6) {
          const [appName, weekRange, level, identifier, sourceApp, campaign, comment, lastUpdated, hash] = existingData[i];
          const key = this.getCommentKey(appName, weekRange, level, identifier, sourceApp, campaign);
          existingDataMap[key] = {
            rowIndex: i + 1,
            existingComment: comment || ''
          };
          
          if (hash) {
            existingHashMap[hash] = {
              rowIndex: i + 1,
              existingComment: comment || ''
            };
          }
        }
      }
      
      const updateRequests = [];
      const newRowsToAdd = [];
      const timestamp = new Date().toISOString();
      
      commentsArray.forEach(commentData => {
        const { appName, weekRange, level, comment, identifier, sourceApp, campaign, hash } = commentData;
        const key = this.getCommentKey(appName, weekRange, level, identifier, sourceApp, campaign);
        
        let existingRow = null;
        if (hash && existingHashMap[hash]) {
          existingRow = existingHashMap[hash];
        } else if (existingDataMap[key]) {
          existingRow = existingDataMap[key];
        }
        
        if (existingRow) {
          if (comment.length > existingRow.existingComment.length) {
            updateRequests.push({
              range: `${this.cacheSheetName}!G${existingRow.rowIndex}:I${existingRow.rowIndex}`,
              values: [[comment, timestamp, hash || '']]
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
            timestamp,
            hash || ''
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
          range: `${this.cacheSheetName}!A${startRow}:I${endRow}`,
          values: newRowsToAdd
        });
      }
      
      if (allRequests.length > 0) {
        const batchUpdateRequest = {
          valueInputOption: 'RAW',
          data: allRequests
        };
        
        Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest, this.cacheSpreadsheetId);
        
        const cacheKey = `${this.cacheSpreadsheetId}_${this.cacheSheetName}!A:I`;
        delete COMMENT_CACHE_GLOBAL.sheetData[cacheKey];
        delete COMMENT_CACHE_GLOBAL.sheetDataTime[cacheKey];
      }
      
      console.log(`Batch saved: ${updateRequests.length} updates, ${newRowsToAdd.length} new comments for ${this.projectName}`);
      
    } catch (e) {
      console.error('Error in batch save comments:', e);
      throw e;
    }
  }

  saveComment(appName, weekRange, level, comment, identifier = null, sourceApp = null, campaign = null, hash = null) {
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
    
    if (!hash) {
      hash = this.generateRowHash(level, appName, weekRange, identifier, sourceApp, campaign);
    }
    
    this.batchSaveComments([{
      appName,
      weekRange, 
      level,
      comment: commentStr,
      identifier,
      sourceApp,
      campaign,
      hash
    }]);
  }

  syncCommentsFromSheet() {
    const sheetName = this.config.SHEET_NAME;
    
    try {
      const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
      const sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        console.log(`${this.projectName}: Sheet ${sheetName} not found`);
        return;
      }
      
      if (sheet.getLastRow() < 1) {
        console.log(`${this.projectName}: Sheet ${sheetName} is empty`);
        return;
      }
      
      const range = `${sheetName}!A:Z`;
      const response = this.getCachedSheetData(this.config.SHEET_ID, range);
      
      if (!response.values || response.values.length < 2) {
        console.log(`No data found in ${sheetName}`);
        return;
      }
      
      const data = response.values;
      const headers = data[0];
      
      console.log(`${this.projectName}: Headers found:`, headers.filter(h => h).join(', '));
      
      const levelCol = this.findColumnByHeader(headers, 'Level') - 1;
      const nameCol = this.findColumnByHeader(headers, 'Week Range / Source App') - 1;
      const idCol = this.findColumnByHeader(headers, 'ID') - 1;
      let commentCol = this.findColumnByHeader(headers, 'Comments') - 1;
      let hashCol = this.findColumnByHeader(headers, 'RowHash') - 1;
      
      if (commentCol === -2) {
        commentCol = this.findColumnByHeader(headers, 'Comment') - 1;
        if (commentCol === -2) {
          console.log(`${this.projectName}: Comments column not found. Available headers:`, headers);
          return;
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
          const hash = hashCol >= 0 ? (row[hashCol] || '') : '';
          
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
                campaign: 'N/A',
                hash: hash || this.generateRowHash('WEEK', currentApp, currentWeek, '', '', '')
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
                campaign: nameOrRange || 'N/A',
                hash: hash || this.generateRowHash('NETWORK', currentApp, currentWeek, idOrEmpty || '', '', nameOrRange || '')
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
                campaign: 'N/A',
                hash: hash || this.generateRowHash('SOURCE_APP', currentApp, currentWeek, nameOrRange || '', nameOrRange || '', '')
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
              campaignName = this.extractCampaignNameFromHyperlink(idOrEmpty) || campaignIdValue || 'Unknown';
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
              campaign: campaignName,
              hash: hash || this.generateRowHash('CAMPAIGN', currentApp, currentWeek, campaignIdValue, sourceAppName || '', campaignName)
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
        console.log(`No data found in ${sheetName}`);
        return;
      }
      
      const data = response.values;
      const headers = data[0];
      
      const levelCol = this.findColumnByHeader(headers, 'Level') - 1;
      const nameCol = this.findColumnByHeader(headers, 'Week Range / Source App') - 1;
      const idCol = this.findColumnByHeader(headers, 'ID') - 1;
      let commentCol = this.findColumnByHeader(headers, 'Comments');
      let hashCol = this.findColumnByHeader(headers, 'RowHash') - 1;
      
      if (commentCol === -1) {
        commentCol = this.findColumnByHeader(headers, 'Comment');
      }
      
      if (commentCol === -1) {
        throw new Error('Comments column not found');
      }
      
      const { comments, commentsByHash } = this.loadAllComments();
      
      let currentApp = '';
      let currentWeek = '';
      const updatesToMake = [];
      let weekComments = 0;
      let sourceAppComments = 0;
      let campaignComments = 0;
      let networkComments = 0;
      let hashMatches = 0;
      let fallbackMatches = 0;
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;
        
        const level = row[levelCol] || '';
        const nameOrRange = row[nameCol] || '';
        const idOrEmpty = row[idCol] || '';
        const hash = hashCol >= 0 ? (row[hashCol] || '') : '';
        
        let foundComment = null;
        let matchType = '';
        
        if (level === 'APP') {
          currentApp = nameOrRange;
          currentWeek = '';
        } else if (level === 'WEEK' && currentApp) {
          currentWeek = nameOrRange;
          
          if (hash && commentsByHash[hash]) {
            foundComment = commentsByHash[hash];
            matchType = 'hash';
            hashMatches++;
          } else {
            const weekKey = this.getCommentKey(currentApp, currentWeek, 'WEEK', 'N/A', 'N/A', 'N/A');
            if (comments[weekKey]) {
              foundComment = comments[weekKey];
              matchType = 'fallback';
              fallbackMatches++;
            }
          }
          
          if (foundComment) {
            const cellRange = `${sheetName}!${this.columnNumberToLetter(commentCol)}${i + 1}`;
            updatesToMake.push({
              range: cellRange,
              values: [[foundComment]]
            });
            weekComments++;
          }
        } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
          const sourceAppDisplayName = nameOrRange;
          
          if (hash && commentsByHash[hash]) {
            foundComment = commentsByHash[hash];
            matchType = 'hash';
            hashMatches++;
          } else {
            const sourceAppKey = this.getCommentKey(currentApp, currentWeek, 'SOURCE_APP', sourceAppDisplayName, sourceAppDisplayName, 'N/A');
            if (comments[sourceAppKey]) {
              foundComment = comments[sourceAppKey];
              matchType = 'fallback';
              fallbackMatches++;
            }
          }
          
          if (foundComment) {
            const cellRange = `${sheetName}!${this.columnNumberToLetter(commentCol)}${i + 1}`;
            updatesToMake.push({
              range: cellRange,
              values: [[foundComment]]
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
          
          if (hash && commentsByHash[hash]) {
            foundComment = commentsByHash[hash];
            matchType = 'hash';
            hashMatches++;
          } else {
            const campaignKey = this.getCommentKey(currentApp, currentWeek, 'CAMPAIGN', 
              this.projectName === 'TRICKY' ? campaignIdValue : 'N/A', 
              sourceAppName, 
              campaignName);
            if (comments[campaignKey]) {
              foundComment = comments[campaignKey];
              matchType = 'fallback';
              fallbackMatches++;
            }
          }
          
          if (foundComment) {
            const cellRange = `${sheetName}!${this.columnNumberToLetter(commentCol)}${i + 1}`;
            updatesToMake.push({
              range: cellRange,
              values: [[foundComment]]
            });
            campaignComments++;
          }
        } else if (level === 'NETWORK' && currentApp && currentWeek) {
          const networkName = nameOrRange;
          
          if (hash && commentsByHash[hash]) {
            foundComment = commentsByHash[hash];
            matchType = 'hash';
            hashMatches++;
          } else {
            const networkKey = this.getCommentKey(currentApp, currentWeek, 'NETWORK', idOrEmpty || 'N/A', 'N/A', networkName);
            if (comments[networkKey]) {
              foundComment = comments[networkKey];
              matchType = 'fallback';
              fallbackMatches++;
            }
          }
          
          if (foundComment) {
            const cellRange = `${sheetName}!${this.columnNumberToLetter(commentCol)}${i + 1}`;
            updatesToMake.push({
              range: cellRange,
              values: [[foundComment]]
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
        console.log(`${this.projectName}: Match types - Hash: ${hashMatches}, Fallback: ${fallbackMatches}`);
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