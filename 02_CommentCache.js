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
    this._cols = null;
  }

  getCachedSpreadsheetMetadata() {
    const now = Date.now();
    const cache = COMMENT_CACHE_GLOBAL;
    
    if (cache.spreadsheetMetadata && cache.spreadsheetMetadataTime &&
        (now - cache.spreadsheetMetadataTime) < cache.CACHE_DURATION) {
      return cache.spreadsheetMetadata;
    }
    
    const metadata = Sheets.Spreadsheets.get(this.cacheSpreadsheetId);
    cache.spreadsheetMetadata = metadata;
    cache.spreadsheetMetadataTime = now;
    return metadata;
  }

  getFreshSheetData(spreadsheetId, range) {
    try {
      const data = Sheets.Spreadsheets.Values.get(spreadsheetId, range);
      const key = `${spreadsheetId}_${range}`;
      COMMENT_CACHE_GLOBAL.sheetData[key] = data;
      COMMENT_CACHE_GLOBAL.sheetDataTime[key] = Date.now();
      return data;
    } catch (e) {
      console.error(`Error fetching ${range}:`, e);
      throw e;
    }
  }

  clearCache() {
    COMMENT_CACHE_GLOBAL = {
      spreadsheetMetadata: null,
      spreadsheetMetadataTime: null,
      sheetData: {},
      sheetDataTime: {},
      CACHE_DURATION: 300000
    };
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
        const response = Sheets.Spreadsheets.batchUpdate({
          requests: [{
            addSheet: { properties: { title: sheetName } }
          }]
        }, this.cacheSpreadsheetId);
        
        sheet = response.replies[0].addSheet;
        
        Sheets.Spreadsheets.batchUpdate({
          requests: [
            {
              updateCells: {
                range: { sheetId: sheet.properties.sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 8 },
                rows: [{
                  values: ['AppName','WeekRange','Level','Identifier','SourceApp','Campaign','Comment','LastUpdated']
                    .map(v => ({ userEnteredValue: { stringValue: v } }))
                }],
                fields: 'userEnteredValue'
              }
            },
            {
              repeatCell: {
                range: { sheetId: sheet.properties.sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 8 },
                cell: { userEnteredFormat: { textFormat: { bold: true }, backgroundColor: { red: 0.94, green: 0.94, blue: 0.94 } } },
                fields: 'userEnteredFormat(textFormat,backgroundColor)'
              }
            }
          ]
        }, this.cacheSpreadsheetId);
        
        COMMENT_CACHE_GLOBAL.spreadsheetMetadata = null;
      }
      
      this.cacheSheetName = sheetName;
      this.cacheSheetId = sheet.properties?.sheetId || sheet.sheetId;
      return { name: this.cacheSheetName, id: this.cacheSheetId };
      
    } catch (e) {
      console.error('Error creating/accessing cache sheet:', e);
      throw e;
    }
  }

  // Объединенный метод для всех колонок - экономия 88 строк!
  getColumns() {
    if (this._cols) return this._cols;
    
    try {
      const headers = this.getFreshSheetData(this.config.SHEET_ID, `${this.config.SHEET_NAME}!1:1`).values?.[0] || [];
      const find = (names) => {
        for (const name of names) {
          const idx = headers.findIndex(h => h?.toString().toLowerCase().trim() === name.toLowerCase());
          if (idx !== -1) return idx + 1;
        }
        return -1;
      };
      
      this._cols = {
        comment: find(['Comments', 'Comment']),
        level: find(['Level']) || 1,
        name: find(['Week Range / Source App', 'Week Range/Source App']) || 2,
        id: find(['ID']) || 3
      };
      
      return this._cols;
    } catch (e) {
      console.error('Error getting columns:', e);
      return { comment: -1, level: 1, name: 2, id: 3 };
    }
  }

  // Legacy методы для обратной совместимости - по 2 строки каждый
  findColumnByHeader(headers, text) { 
    return headers.findIndex(h => h?.toString().toLowerCase().trim() === text.toLowerCase().trim()) + 1; 
  }
  getSheetHeaders(sheetName) { 
    return this.getFreshSheetData(this.config.SHEET_ID, `${sheetName}!1:1`).values?.[0] || []; 
  }
  getCommentColumn(sheetName) { return this.getColumns().comment; }
  getLevelColumn(sheetName) { return this.getColumns().level; }
  getNameColumn(sheetName) { return this.getColumns().name; }
  getIdColumn(sheetName) { return this.getColumns().id; }

  getCommentKey(appName, weekRange, level, identifier = null, sourceApp = null, campaign = null, country = null) {
    return [appName, weekRange, level, identifier || 'N/A', sourceApp || 'N/A', campaign || 'N/A', country || 'N/A'].join('|||');
  }

  loadAllComments() {
  this.getOrCreateCacheSheet();
  const comments = {};
  
  try {
    const response = this.getFreshSheetData(this.cacheSpreadsheetId, `${this.cacheSheetName}!A:H`);
    if (!response.values || response.values.length <= 1) return comments;
    
    response.values.slice(1).forEach(row => {
      if (row.length >= 7 && row[6]) {
        comments[this.getCommentKey(...row.slice(0, 6))] = row[6];
      }
    });
    
    // Добавляем legacy ключи для обратной совместимости
    if (typeof APP_NAME_LEGACY !== 'undefined') {
      Object.keys(APP_NAME_LEGACY).forEach(newName => {
        const oldName = APP_NAME_LEGACY[newName];
        response.values.slice(1).forEach(row => {
          if (row.length >= 7 && row[6] && row[0] === oldName) {
            const legacyKey = this.getCommentKey(newName, row[1], row[2], row[3], row[4], row[5]);
            if (!comments[legacyKey]) comments[legacyKey] = row[6];
          }
        });
      });
    }
  } catch (e) {
    console.error('Error loading comments:', e);
  }
  
  return comments;
}

  batchSaveComments(commentsArray) {
  if (!commentsArray?.length) return;
  
  this.getOrCreateCacheSheet();
  
  // Нормализация названий приложений - заменяем старые названия на новые
  if (typeof APP_NAME_LEGACY !== 'undefined') {
    commentsArray = commentsArray.map(item => {
      // Ищем, не является ли appName старым названием
      const newName = Object.keys(APP_NAME_LEGACY).find(key => APP_NAME_LEGACY[key] === item.appName);
      return {
        ...item,
        appName: newName || item.appName // Если нашли - заменяем на новое, иначе оставляем как есть
      };
    });
  }
  
  try {
    const existing = this.getFreshSheetData(this.cacheSpreadsheetId, `${this.cacheSheetName}!A:H`).values || [];
    const existingMap = new Map();
    
    existing.slice(1).forEach((row, i) => {
      if (row.length >= 6) {
        existingMap.set(this.getCommentKey(...row.slice(0, 6)), {
          rowIndex: i + 2,
          comment: row[6] || ''
        });
      }
    });
    
    const timestamp = new Date().toISOString();
    const updates = [];
    const newRows = [];
    
    commentsArray.forEach(({ appName, weekRange, level, comment, identifier, sourceApp, campaign }) => {
      const key = this.getCommentKey(appName, weekRange, level, identifier, sourceApp, campaign);
      const existing = existingMap.get(key);
      
      if (existing) {
        if (comment.length > existing.comment.length) {
          updates.push({
            range: `${this.cacheSheetName}!G${existing.rowIndex}:H${existing.rowIndex}`,
            values: [[comment, timestamp]]
          });
        }
      } else {
        newRows.push([appName, weekRange, level, identifier || 'N/A', sourceApp || 'N/A', 
                     campaign || 'N/A', comment, timestamp]);
      }
    });
    
    const requests = [...updates];
    if (newRows.length) {
      const start = existing.length + 1;
      requests.push({
        range: `${this.cacheSheetName}!A${start}:H${start + newRows.length - 1}`,
        values: newRows
      });
    }
    
    if (requests.length) {
      Sheets.Spreadsheets.Values.batchUpdate({
        valueInputOption: 'RAW',
        data: requests
      }, this.cacheSpreadsheetId);
      
      delete COMMENT_CACHE_GLOBAL.sheetData[`${this.cacheSpreadsheetId}_${this.cacheSheetName}!A:H`];
    }
    
    console.log(`${this.projectName}: Saved ${updates.length + newRows.length} comments`);
  } catch (e) {
    console.error('Error saving comments:', e);
    throw e;
  }
}

  // Упрощенные методы валидации и синхронизации
  validateRowContent(row, level, expected) {
    if (!row?.length >= 3 || row[0] !== level) return false;
    
    const validators = {
      WEEK: () => row[1] === expected.weekRange,
      SOURCE_APP: () => row[1]?.includes(expected.sourceAppName),
      CAMPAIGN: () => {
        const id = this.extractCampaignIdFromHyperlink(row[2]) || row[2];
        return row[1] === expected.sourceAppName && id === expected.campaignId;
      },
      NETWORK: () => row[1] === expected.networkName && (!row[2] || row[2] === expected.networkId)
    };
    
    return validators[level]?.() || false;
  }

  syncCommentsFromSheet() {
    try {
      const data = this.getFreshSheetData(this.config.SHEET_ID, `${this.config.SHEET_NAME}!A:Z`).values;
      if (!data || data.length < 2) return console.log(`${this.projectName}: No data`);
      
      const cols = this.getColumns();
      if (cols.comment === -1) return console.log(`${this.projectName}: Comments column not found`);
      
      const comments = [];
      let currentApp = '', currentWeek = '';
      
      data.slice(1).forEach(row => {
        if (!row?.length) return;
        
        const [level, nameOrRange, idOrEmpty] = [row[cols.level-1], row[cols.name-1], row[cols.id-1]];
        const comment = row[cols.comment-1];
        
        if (level === 'APP') {
          currentApp = nameOrRange;
          currentWeek = '';
        } else if (level === 'WEEK' && currentApp) {
          currentWeek = nameOrRange;
          if (comment) comments.push({ appName: currentApp, weekRange: currentWeek, level: 'WEEK', 
                                       comment, identifier: 'N/A', sourceApp: 'N/A', campaign: 'N/A' });
        } else if (comment && currentApp && currentWeek) {
          let config = null;
          
          if (level === 'COUNTRY' && this.projectName === 'APPLOVIN_TEST') {
            // Для APPLOVIN_TEST найдем текущую кампанию
            let currentCampaign = '';
            for (let j = data.indexOf(row) - 1; j >= 0; j--) {
              if (data[j] && data[j][cols.level - 1] === 'CAMPAIGN') {
                currentCampaign = data[j][cols.name - 1];
                break;
              }
            }
            config = {
              identifier: idOrEmpty || 'N/A',
              sourceApp: currentCampaign || 'N/A',
              campaign: nameOrRange || 'N/A', // Название страны
              country: nameOrRange || 'N/A'
            };
          } else if (this.projectName === 'INCENT_TRAFFIC') {
            // Специальная обработка для новой структуры INCENT_TRAFFIC
            if (level === 'COUNTRY') {
              config = {
                identifier: idOrEmpty || 'N/A', // код страны из колонки GEO
                sourceApp: 'N/A',
                campaign: nameOrRange || 'N/A', // полное название страны
                country: idOrEmpty || 'N/A' // код страны
              };
            } else if (level === 'CAMPAIGN') {
              config = {
                identifier: idOrEmpty || 'N/A', // campaign ID
                sourceApp: nameOrRange || 'N/A', // campaign name
                campaign: idOrEmpty || 'N/A' // campaign ID
              };
            } else {
              config = {
                identifier: idOrEmpty || 'N/A',
                sourceApp: 'N/A',
                campaign: nameOrRange || 'N/A'
              };
            }
          } else {
            config = {
              NETWORK: { identifier: idOrEmpty || 'N/A', sourceApp: 'N/A', campaign: nameOrRange || 'N/A' },
              SOURCE_APP: { identifier: nameOrRange || 'N/A', sourceApp: nameOrRange || 'N/A', campaign: 'N/A' },
              CAMPAIGN: { 
                identifier: this.projectName === 'TRICKY' ? (this.extractCampaignIdFromHyperlink(idOrEmpty) || idOrEmpty) : 'N/A',
                sourceApp: nameOrRange || 'N/A',
                campaign: this.projectName === 'TRICKY' ? (this.extractCampaignIdFromHyperlink(idOrEmpty) || idOrEmpty || 'Unknown') : (nameOrRange || 'Unknown')
              },
              COUNTRY: { 
                identifier: idOrEmpty || 'N/A',
                sourceApp: nameOrRange || 'N/A',
                campaign: 'N/A',
                country: nameOrRange || 'N/A'
              }
            }[level];
          }
          
          if (config) comments.push({ appName: currentApp, weekRange: currentWeek, level, comment, ...config });
        }
      });
      
      if (comments.length) {
        this.batchSaveComments(comments);
        console.log(`${this.projectName}: Synced ${comments.length} comments`);
      }
    } catch (e) {
      console.error(`Error syncing comments for ${this.projectName}:`, e);
      throw e;
    }
  }

  applyCommentsToSheet() {
    try {
      const data = this.getFreshSheetData(this.config.SHEET_ID, `${this.config.SHEET_NAME}!A:Z`).values;
      if (!data || data.length < 2) return console.log(`${this.projectName}: No data`);
      
      const cols = this.getColumns();
      if (cols.comment === -1) throw new Error('Comments column not found');
      
      const comments = this.loadAllComments();
      const updates = [];
      let currentApp = '', currentWeek = '';
      
      data.slice(1).forEach((row, i) => {
        if (!row?.length) return;
        
        const [level, nameOrRange, idOrEmpty] = [row[cols.level-1], row[cols.name-1], row[cols.id-1]];
        
        if (level === 'APP') {
          currentApp = nameOrRange;
          currentWeek = '';
        } else if (level === 'WEEK' && currentApp) {
          currentWeek = nameOrRange;
        }
        
        if (!currentApp || !currentWeek) return;
        
        const keys = {
          WEEK: () => this.getCommentKey(currentApp, currentWeek, 'WEEK', 'N/A', 'N/A', 'N/A'),
          SOURCE_APP: () => this.getCommentKey(currentApp, currentWeek, 'SOURCE_APP', nameOrRange, nameOrRange, 'N/A'),
          CAMPAIGN: () => {
            if (this.projectName === 'INCENT_TRAFFIC') {
              return this.getCommentKey(currentApp, currentWeek, 'CAMPAIGN', idOrEmpty || 'N/A', nameOrRange, idOrEmpty);
            }
            const id = this.extractCampaignIdFromHyperlink(idOrEmpty) || idOrEmpty;
            const name = this.projectName === 'TRICKY' ? id : nameOrRange;
            return this.getCommentKey(currentApp, currentWeek, 'CAMPAIGN', 
                                     this.projectName === 'TRICKY' ? id : 'N/A', nameOrRange, name);
          },
          NETWORK: () => this.getCommentKey(currentApp, currentWeek, 'NETWORK', idOrEmpty || 'N/A', 'N/A', nameOrRange),
          COUNTRY: () => {
            // Для INCENT_TRAFFIC страны имеют свою структуру
            if (this.projectName === 'INCENT_TRAFFIC') {
              return this.getCommentKey(currentApp, currentWeek, 'COUNTRY', idOrEmpty || 'N/A', 'N/A', nameOrRange, idOrEmpty);
            }
            // Для APPLOVIN_TEST страны связаны с кампаниями
            else if (this.projectName === 'APPLOVIN_TEST') {
              // Нужно найти текущую кампанию - ищем вверх до ближайшей кампании
              for (let j = i - 1; j >= 0; j--) {
                const prevRow = data[j];
                if (prevRow && prevRow[cols.level - 1] === 'CAMPAIGN') {
                  const campaignName = prevRow[cols.name - 1];
                  return this.getCommentKey(currentApp, currentWeek, 'COUNTRY', idOrEmpty || 'N/A', campaignName, nameOrRange);
                }
              }
            }
            return this.getCommentKey(currentApp, currentWeek, 'COUNTRY', idOrEmpty || 'N/A', nameOrRange, 'N/A', nameOrRange);
          }
        };
        
        const key = keys[level]?.();
        const comment = key && comments[key];
        
        if (comment) {
          updates.push({
            range: `${this.config.SHEET_NAME}!${String.fromCharCode(64 + cols.comment)}${i + 2}`,
            values: [[comment]]
          });
        }
      });
      
      if (updates.length) {
        Sheets.Spreadsheets.Values.batchUpdate({
          valueInputOption: 'RAW',
          data: updates
        }, this.config.SHEET_ID);
        
        console.log(`${this.projectName}: Applied ${updates.length} comments`);
      }
    } catch (e) {
      console.error(`Error applying comments:`, e);
      throw e;
    }
  }

  // Утилиты - по 1-2 строки
  extractCampaignIdFromHyperlink(formula) {
    return formula?.match?.(/campaigns\/([^"]+)/)?.[1] || 'Unknown';
  }
  extractCampaignNameFromHyperlink(formula) {
    return formula?.match?.(/["']([^"']+)["']\s*\)$/)?.[1] || null;
  }
  columnNumberToLetter(n) {
    return String.fromCharCode(64 + n);
  }
  findRowForComment() { return { rowIndex: -1, validated: false }; } // deprecated
  syncCommentsFromSheetQuiet() { this.syncCommentsFromSheet(); }
  clearColumnCache() { this._cols = null; }
  saveComment(appName, weekRange, level, comment, identifier, sourceApp, campaign) {
    if (comment) this.batchSaveComments([{ appName, weekRange, level, comment: String(comment), identifier, sourceApp, campaign }]);
  }
}