//пук22

class CommentCache {
  constructor(projectName = null) {
    this.projectName = projectName || CURRENT_PROJECT;
    this.config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
    this.cacheSheet = this.getOrCreateCacheSheet();
    this.commentColumnIndex = null;
  }

  getOrCreateCacheSheet() {
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    let sheet = spreadsheet.getSheetByName(this.config.COMMENTS_CACHE_SHEET);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(this.config.COMMENTS_CACHE_SHEET);
      sheet.hideSheet();
      const headers = [['AppName', 'WeekRange', 'Level', 'Identifier', 'SourceApp', 'Comment', 'LastUpdated']];
      Sheets.Spreadsheets.Values.update({
        majorDimension: 'ROWS',
        values: headers
      }, this.config.SHEET_ID, `${this.config.COMMENTS_CACHE_SHEET}!A1:G1`, {
        valueInputOption: 'RAW'
      });
    }
    return sheet;
  }

  findCommentColumnIndex() {
    if (this.commentColumnIndex !== null) return this.commentColumnIndex;
    
    try {
      const result = Sheets.Spreadsheets.Values.get(
        this.config.SHEET_ID,
        `${this.config.SHEET_NAME}!1:1`
      );
      
      if (result.values && result.values[0]) {
        const headers = result.values[0];
        for (let i = 0; i < headers.length; i++) {
          if (headers[i] === 'Comments') {
            this.commentColumnIndex = i;
            return i;
          }
        }
      }
    } catch (e) {}
    
    this.commentColumnIndex = 18;
    return 18;
  }

  getCommentKey(appName, weekRange, level, identifier = null, sourceApp = null) {
    return `${appName}|||${weekRange}|||${level}|||${identifier || 'N/A'}|||${sourceApp || 'N/A'}`;
  }

  loadAllComments() {
    const comments = {};
    try {
      const result = Sheets.Spreadsheets.Values.get(
        this.config.SHEET_ID, 
        `${this.config.COMMENTS_CACHE_SHEET}!A2:G`
      );
      if (result.values) {
        result.values.forEach(row => {
          if (row[5]) {
            const key = this.getCommentKey(row[0], row[1], row[2], row[3], row[4]);
            comments[key] = row[5];
          }
        });
      }
    } catch (e) {}
    return comments;
  }

  syncCommentsFromSheet() {
    try {
      const commentColIndex = this.findCommentColumnIndex();
      const lastCol = this.getColumnLetter(commentColIndex);
      
      const result = Sheets.Spreadsheets.Values.get(
        this.config.SHEET_ID,
        `${this.config.SHEET_NAME}!A:${lastCol}`
      );
      
      if (!result.values || result.values.length < 2) return;
      
      const data = result.values;
      let currentApp = '';
      let currentWeek = '';
      
      const commentsToSave = [];
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length < 1) continue;
        
        const level = row[0];
        const nameOrRange = row[1] || '';
        const idOrEmpty = row[2] || '';
        const comment = row.length > commentColIndex ? row[commentColIndex] : null;
        
        if (level === 'APP') {
          currentApp = nameOrRange;
          currentWeek = '';
        } else if (level === 'WEEK' && currentApp) {
          currentWeek = nameOrRange;
          if (comment && comment.toString().trim()) {
            commentsToSave.push({
              appName: currentApp,
              weekRange: currentWeek,
              level: 'WEEK',
              comment: comment.toString(),
              identifier: null,
              sourceApp: null
            });
          }
        } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
          if (comment && comment.toString().trim()) {
            commentsToSave.push({
              appName: currentApp,
              weekRange: currentWeek,
              level: 'SOURCE_APP',
              comment: comment.toString(),
              identifier: nameOrRange,
              sourceApp: null
            });
          }
        } else if (level === 'CAMPAIGN' && currentApp && currentWeek) {
          if (comment && comment.toString().trim()) {
            const sourceAppName = nameOrRange;
            const campaignIdValue = idOrEmpty && typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK') 
              ? this.extractCampaignIdFromHyperlink(idOrEmpty) 
              : idOrEmpty;
            
            commentsToSave.push({
              appName: currentApp,
              weekRange: currentWeek,
              level: 'CAMPAIGN',
              comment: comment.toString(),
              identifier: campaignIdValue,
              sourceApp: sourceAppName
            });
          }
        }
      }
      
      if (commentsToSave.length > 0) {
        this.batchSaveComments(commentsToSave);
      }
    } catch (e) {}
  }

  getColumnLetter(columnIndex) {
    let letter = '';
    let tempIndex = columnIndex;
    
    while (tempIndex >= 0) {
      letter = String.fromCharCode(65 + (tempIndex % 26)) + letter;
      tempIndex = Math.floor(tempIndex / 26) - 1;
    }
    
    return letter;
  }

  batchSaveComments(commentsToSave) {
    try {
      const existingResult = Sheets.Spreadsheets.Values.get(
        this.config.SHEET_ID,
        `${this.config.COMMENTS_CACHE_SHEET}!A:G`
      );
      
      const existingData = existingResult.values || [['AppName', 'WeekRange', 'Level', 'Identifier', 'SourceApp', 'Comment', 'LastUpdated']];
      const existingMap = new Map();
      
      for (let i = 1; i < existingData.length; i++) {
        const row = existingData[i];
        if (row && row.length >= 5) {
          const key = `${row[0]}|||${row[1]}|||${row[2]}|||${row[3]}|||${row[4]}`;
          existingMap.set(key, i);
        }
      }
      
      const updates = [];
      const appends = [];
      
      commentsToSave.forEach(item => {
        const key = this.getCommentKey(item.appName, item.weekRange, item.level, item.identifier, item.sourceApp);
        const rowIndex = existingMap.get(key);
        
        const rowData = [
          item.appName,
          item.weekRange,
          item.level,
          item.identifier || 'N/A',
          item.sourceApp || 'N/A',
          item.comment,
          new Date().toISOString()
        ];
        
        if (rowIndex !== undefined) {
          updates.push({
            range: `${this.config.COMMENTS_CACHE_SHEET}!A${rowIndex + 1}:G${rowIndex + 1}`,
            values: [rowData]
          });
        } else {
          appends.push(rowData);
        }
      });
      
      if (updates.length > 0) {
        const batchUpdateRequest = {
          data: updates,
          valueInputOption: 'RAW'
        };
        Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest, this.config.SHEET_ID);
      }
      
      if (appends.length > 0) {
        Sheets.Spreadsheets.Values.append({
          majorDimension: 'ROWS',
          values: appends
        }, this.config.SHEET_ID, `${this.config.COMMENTS_CACHE_SHEET}!A:G`, {
          valueInputOption: 'RAW',
          insertDataOption: 'INSERT_ROWS'
        });
      }
    } catch (e) {}
  }

  applyCommentsToSheet() {
    const comments = this.loadAllComments();
    if (Object.keys(comments).length === 0) return;
    
    try {
      const commentColIndex = this.findCommentColumnIndex();
      const commentColLetter = this.getColumnLetter(commentColIndex);
      
      const result = Sheets.Spreadsheets.Values.get(
        this.config.SHEET_ID,
        `${this.config.SHEET_NAME}!A:C`
      );
      
      if (!result.values || result.values.length < 2) return;
      
      const data = result.values;
      const updates = [];
      let currentApp = '';
      let currentWeek = '';
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length < 1) continue;
        
        const level = row[0];
        const nameOrRange = row[1];
        const idOrEmpty = row[2];
        
        if (level === 'APP') {
          currentApp = nameOrRange;
          currentWeek = '';
        } else if (level === 'WEEK' && currentApp) {
          currentWeek = nameOrRange;
          const weekKey = this.getCommentKey(currentApp, currentWeek, 'WEEK');
          if (comments[weekKey]) {
            updates.push({
              range: `${this.config.SHEET_NAME}!${commentColLetter}${i + 1}`,
              values: [[comments[weekKey]]]
            });
          }
        } else if (level === 'SOURCE_APP' && currentApp && currentWeek) {
          const sourceAppKey = this.getCommentKey(currentApp, currentWeek, 'SOURCE_APP', nameOrRange);
          if (comments[sourceAppKey]) {
            updates.push({
              range: `${this.config.SHEET_NAME}!${commentColLetter}${i + 1}`,
              values: [[comments[sourceAppKey]]]
            });
          }
        } else if (level === 'CAMPAIGN' && currentApp && currentWeek) {
          const sourceAppName = nameOrRange;
          const campaignIdValue = idOrEmpty && typeof idOrEmpty === 'string' && idOrEmpty.includes('HYPERLINK') 
            ? this.extractCampaignIdFromHyperlink(idOrEmpty) 
            : idOrEmpty;
          
          const campaignKey = this.getCommentKey(currentApp, currentWeek, 'CAMPAIGN', campaignIdValue, sourceAppName);
          if (comments[campaignKey]) {
            updates.push({
              range: `${this.config.SHEET_NAME}!${commentColLetter}${i + 1}`,
              values: [[comments[campaignKey]]]
            });
          }
        }
      }
      
      if (updates.length > 0) {
        const batchUpdateRequest = {
          data: updates,
          valueInputOption: 'RAW'
        };
        Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest, this.config.SHEET_ID);
      }
    } catch (e) {}
  }

  extractCampaignIdFromHyperlink(hyperlinkFormula) {
    try {
      const match = hyperlinkFormula.match(/campaigns\/([^"]+)/);
      return match ? match[1] : 'Unknown';
    } catch (e) {
      return 'Unknown';
    }
  }

  saveComment(appName, weekRange, level, comment, identifier = null, sourceApp = null) {
    if (!comment || !comment.trim()) return;
    
    try {
      const rowData = [[
        appName, 
        weekRange, 
        level,
        identifier || 'N/A', 
        sourceApp || 'N/A', 
        comment, 
        new Date().toISOString()
      ]];
      
      Sheets.Spreadsheets.Values.append({
        majorDimension: 'ROWS',
        values: rowData
      }, this.config.SHEET_ID, `${this.config.COMMENTS_CACHE_SHEET}!A:G`, {
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS'
      });
    } catch (e) {}
  }

  syncCommentsFromSheetQuiet() {
    this.syncCommentsFromSheet();
  }
}