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
                range: { sheetId: sheet.properties.sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 9 },
                rows: [{
                  values: ['AppName','WeekRange','Level','Identifier','SourceApp','Campaign','Comment','LastUpdated','Country']
                    .map(v => ({ userEnteredValue: { stringValue: v } }))
                }],
                fields: 'userEnteredValue'
              }
            },
            {
              repeatCell: {
                range: { sheetId: sheet.properties.sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 9 },
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
        comment: COLUMN_CONFIG.COLUMNS.COMMENTS,  // Используем из конфига
        level: COLUMN_CONFIG.COLUMNS.LEVEL,
        name: COLUMN_CONFIG.COLUMNS.WEEK_RANGE,
        id: COLUMN_CONFIG.COLUMNS.ID
      };
      
      return this._cols;
    } catch (e) {
      console.error('Error getting columns:', e);
      return { comment: -1, level: 1, name: 2, id: 3 };
    }
  }

  // Legacy методы для обратной совместимости
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
      const response = this.getFreshSheetData(this.cacheSpreadsheetId, `${this.cacheSheetName}!A:I`);
      if (!response.values || response.values.length <= 1) return comments;
      
      response.values.slice(1).forEach(row => {
        if (row.length >= 7 && row[6]) {
          // Country в самой последней колонке (позиция 8)
          const country = row.length >= 9 ? row[8] : 'N/A';
          comments[this.getCommentKey(...row.slice(0, 6), country)] = row[6];
        }
      });
      
      // Legacy ключи для обратной совместимости
      if (typeof APP_NAME_LEGACY !== 'undefined') {
        Object.keys(APP_NAME_LEGACY).forEach(newName => {
          const oldName = APP_NAME_LEGACY[newName];
          response.values.slice(1).forEach(row => {
            if (row.length >= 7 && row[6] && row[0] === oldName) {
              const country = row.length >= 9 ? row[8] : 'N/A';
              const legacyKey = this.getCommentKey(newName, row[1], row[2], row[3], row[4], row[5], country);
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
    
    // Нормализация названий приложений
    if (typeof APP_NAME_LEGACY !== 'undefined') {
      commentsArray = commentsArray.map(item => {
        const newName = Object.keys(APP_NAME_LEGACY).find(key => APP_NAME_LEGACY[key] === item.appName);
        return {
          ...item,
          appName: newName || item.appName
        };
      });
    }
    
    try {
      const existing = this.getFreshSheetData(this.cacheSpreadsheetId, `${this.cacheSheetName}!A:I`).values || [];
      const existingMap = new Map();
      
      existing.slice(1).forEach((row, i) => {
        if (row.length >= 7) {
          const country = row.length >= 9 ? row[8] : 'N/A';
          existingMap.set(this.getCommentKey(...row.slice(0, 6), country), {
            rowIndex: i + 2,
            comment: row[6] || ''
          });
        }
      });
      
      const timestamp = new Date().toISOString();
      const updates = [];
      const newRows = [];
      
      commentsArray.forEach(({ appName, weekRange, level, comment, identifier, sourceApp, campaign, country }) => {
        const key = this.getCommentKey(appName, weekRange, level, identifier, sourceApp, campaign, country);
        const existing = existingMap.get(key);
        
        if (existing) {
          if (comment.length > existing.comment.length) {
            updates.push({
              range: `${this.cacheSheetName}!G${existing.rowIndex}:I${existing.rowIndex}`,
              values: [[comment, timestamp, country || 'N/A']]
            });
          }
        } else {
          newRows.push([appName, weekRange, level, identifier || 'N/A', sourceApp || 'N/A', 
                       campaign || 'N/A', comment, timestamp, country || 'N/A']);
        }
      });
      
      const requests = [...updates];
      if (newRows.length) {
        const start = existing.length + 1;
        requests.push({
          range: `${this.cacheSheetName}!A${start}:I${start + newRows.length - 1}`,
          values: newRows
        });
      }
      
      if (requests.length) {
        Sheets.Spreadsheets.Values.batchUpdate({
          valueInputOption: 'RAW',
          data: requests
        }, this.cacheSpreadsheetId);
        
        delete COMMENT_CACHE_GLOBAL.sheetData[`${this.cacheSpreadsheetId}_${this.cacheSheetName}!A:I`];
      }
      
      console.log(`${this.projectName}: Saved ${updates.length + newRows.length} comments`);
    } catch (e) {
      console.error('Error saving comments:', e);
      throw e;
    }
  }

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
      let currentApp = '', currentWeek = '', currentCampaign = '';
      
      // Для INCENT_TRAFFIC: отслеживаем network, country, campaign
      let currentNetwork = '', currentCountry = '';
      
      data.slice(1).forEach(row => {
        if (!row?.length) return;
        
        const [level, nameOrRange, idOrEmpty] = [row[cols.level-1], row[cols.name-1], row[cols.id-1]];
        const comment = row[cols.comment-1];
        
        if (level === 'APP') {
          currentApp = nameOrRange;
          currentWeek = '';
          currentCampaign = '';
        } else if (level === 'NETWORK' && this.projectName === 'INCENT_TRAFFIC') {
          // Для INCENT_TRAFFIC: NETWORK - это аналог APP
          currentNetwork = nameOrRange;
          currentCountry = '';
          currentCampaign = '';
          currentWeek = '';
          
          if (comment) {
            comments.push({ 
              appName: currentNetwork, // Используем network как appName 
              weekRange: '', 
              level: 'NETWORK', 
              comment, 
              identifier: 'N/A',
              sourceApp: nameOrRange,
              campaign: 'N/A' 
            });
          }
        } else if (level === 'COUNTRY' && this.projectName === 'INCENT_TRAFFIC') {
          currentCountry = nameOrRange;
          currentCampaign = '';
          currentWeek = '';
          
          if (comment && currentNetwork) {
            comments.push({ 
              appName: currentNetwork, 
              weekRange: '', 
              level: 'COUNTRY', 
              comment, 
              identifier: idOrEmpty || 'N/A', // countryCode
              sourceApp: nameOrRange, // countryName
              campaign: 'N/A',
              country: idOrEmpty || 'N/A'
            });
          }
        } else if (level === 'CAMPAIGN' && this.projectName === 'APPLOVIN_TEST') {
          // Для APPLOVIN_TEST: CAMPAIGN на втором уровне после APP
          currentCampaign = nameOrRange;
          currentWeek = ''; // Сбрасываем неделю при новой кампании
          if (comment && currentApp) {
            comments.push({ 
              appName: currentApp, 
              weekRange: '', // Для кампании weekRange пустой
              level: 'CAMPAIGN', 
              comment, 
              identifier: idOrEmpty || 'N/A', 
              sourceApp: nameOrRange || 'N/A', 
              campaign: idOrEmpty || 'N/A' 
            });
          }
        } else if (level === 'CAMPAIGN' && this.projectName === 'INCENT_TRAFFIC') {
          // Для INCENT_TRAFFIC: CAMPAIGN внутри COUNTRY внутри NETWORK
          currentCampaign = nameOrRange;
          currentWeek = '';
          
          if (comment && currentNetwork && currentCountry) {
            comments.push({ 
              appName: currentNetwork, 
              weekRange: '', 
              level: 'CAMPAIGN', 
              comment, 
              identifier: idOrEmpty || 'N/A', // campaignId
              sourceApp: nameOrRange, // campaignName
              campaign: idOrEmpty || 'N/A', // campaignId
              country: currentCountry
            });
          }
        } else if (level === 'WEEK') {
          if (this.projectName === 'APPLOVIN_TEST') {
            // Для APPLOVIN_TEST: WEEK внутри CAMPAIGN
            currentWeek = nameOrRange;
            if (comment && currentApp && currentCampaign) {
              comments.push({ 
                appName: currentApp, 
                weekRange: currentWeek, 
                level: 'WEEK', 
                comment, 
                identifier: currentCampaign, // Сохраняем ID кампании
                sourceApp: currentCampaign, // Имя кампании
                campaign: 'N/A' 
              });
            }
          } else if (this.projectName === 'INCENT_TRAFFIC') {
            // Для INCENT_TRAFFIC: WEEK внутри CAMPAIGN внутри COUNTRY внутри NETWORK
            currentWeek = nameOrRange;
            if (comment && currentNetwork && currentCountry && currentCampaign) {
              comments.push({ 
                appName: currentNetwork, 
                weekRange: currentWeek, 
                level: 'WEEK', 
                comment, 
                identifier: currentCampaign, // campaignId
                sourceApp: currentCampaign, // campaignName 
                campaign: currentCampaign, // campaignId
                country: currentCountry
              });
            }
          } else {
            // Стандартная логика для других проектов
            currentWeek = nameOrRange;
            if (comment && currentApp) {
              comments.push({ 
                appName: currentApp, 
                weekRange: currentWeek, 
                level: 'WEEK', 
                comment, 
                identifier: 'N/A', 
                sourceApp: 'N/A', 
                campaign: 'N/A' 
              });
            }
          }
        } else if (comment && (currentApp || (this.projectName === 'INCENT_TRAFFIC' && currentNetwork))) {
          // Обработка других уровней
          let config = null;
          
          if (level === 'COUNTRY' && this.projectName === 'APPLOVIN_TEST') {
            // Для APPLOVIN_TEST: COUNTRY внутри WEEK внутри CAMPAIGN
            config = {
              identifier: `${currentCampaign}_${idOrEmpty || 'N/A'}`, // CampaignId_CountryCode
              sourceApp: currentCampaign || 'N/A', // Имя кампании
              campaign: nameOrRange || 'N/A', // Название страны
              country: idOrEmpty || 'N/A' // Код страны
            };
          } else if (this.projectName === 'INCENT_TRAFFIC') {
            // Для INCENT_TRAFFIC все случаи уже обработаны выше, пропускаем
            config = null;
          } else {
            // Стандартная логика
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
          
          if (config && currentWeek) {
            const appName = this.projectName === 'INCENT_TRAFFIC' ? currentNetwork : currentApp;
            comments.push({ appName, weekRange: currentWeek, level, comment, ...config });
          } else if (config && this.projectName === 'APPLOVIN_TEST' && level === 'COUNTRY') {
            // Для APPLOVIN_TEST страны могут быть без явного weekRange в текущей строке
            comments.push({ appName: currentApp, weekRange: currentWeek || '', level, comment, ...config });
          }
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
      let currentApp = '', currentWeek = '', currentCampaign = '';
      
      // Для INCENT_TRAFFIC: отслеживаем network, country, campaign
      let currentNetwork = '', currentCountry = '';
      
      data.slice(1).forEach((row, i) => {
        if (!row?.length) return;
        
        const [level, nameOrRange, idOrEmpty] = [row[cols.level-1], row[cols.name-1], row[cols.id-1]];
        
        if (level === 'APP') {
          currentApp = nameOrRange;
          currentWeek = '';
          currentCampaign = '';
        } else if (level === 'NETWORK' && this.projectName === 'INCENT_TRAFFIC') {
          // Для INCENT_TRAFFIC: NETWORK - это аналог APP
          currentNetwork = nameOrRange;
          currentCountry = '';
          currentCampaign = '';
          currentWeek = '';
          
          // Применяем комментарий к NETWORK если есть
          const key = this.getCommentKey(currentNetwork, '', 'NETWORK', 'N/A', nameOrRange, 'N/A');
          const comment = comments[key];
          
          if (comment) {
            updates.push({
              range: `${this.config.SHEET_NAME}!${String.fromCharCode(64 + cols.comment)}${i + 2}`,
              values: [[comment]]
            });
          }
        } else if (level === 'COUNTRY' && this.projectName === 'INCENT_TRAFFIC') {
          currentCountry = nameOrRange;
          currentCampaign = '';
          currentWeek = '';
          
          // Применяем комментарий к COUNTRY если есть
          const key = this.getCommentKey(currentNetwork, '', 'COUNTRY', idOrEmpty || 'N/A', nameOrRange, 'N/A', idOrEmpty);
          const comment = comments[key];
          
          if (comment) {
            updates.push({
              range: `${this.config.SHEET_NAME}!${String.fromCharCode(64 + cols.comment)}${i + 2}`,
              values: [[comment]]
            });
          }
        } else if (level === 'CAMPAIGN' && this.projectName === 'APPLOVIN_TEST') {
          currentCampaign = nameOrRange;
          const campaignId = idOrEmpty || 'N/A';
          currentWeek = ''; // Сбрасываем неделю для новой кампании
          
          // Генерация ключа для уровня CAMPAIGN в APPLOVIN_TEST
          const key = this.getCommentKey(currentApp, '', 'CAMPAIGN', campaignId, nameOrRange, campaignId);
          const comment = comments[key];
          
          if (comment) {
            updates.push({
              range: `${this.config.SHEET_NAME}!${String.fromCharCode(64 + cols.comment)}${i + 2}`,
              values: [[comment]]
            });
          }
        } else if (level === 'CAMPAIGN' && this.projectName === 'INCENT_TRAFFIC') {
          // Для INCENT_TRAFFIC: CAMPAIGN внутри COUNTRY внутри NETWORK
          currentCampaign = nameOrRange;
          currentWeek = '';
          
          // Применяем комментарий к CAMPAIGN если есть
          const key = this.getCommentKey(currentNetwork, '', 'CAMPAIGN', idOrEmpty || 'N/A', nameOrRange, idOrEmpty, currentCountry);
          const comment = comments[key];
          
          if (comment) {
            updates.push({
              range: `${this.config.SHEET_NAME}!${String.fromCharCode(64 + cols.comment)}${i + 2}`,
              values: [[comment]]
            });
          }
        } else if (level === 'WEEK') {
          if (this.projectName === 'APPLOVIN_TEST') {
            currentWeek = nameOrRange;
            // Для APPLOVIN_TEST: WEEK связана с кампанией
            const key = this.getCommentKey(currentApp, currentWeek, 'WEEK', currentCampaign, currentCampaign, 'N/A');
            const comment = comments[key];
            
            if (comment) {
              updates.push({
                range: `${this.config.SHEET_NAME}!${String.fromCharCode(64 + cols.comment)}${i + 2}`,
                values: [[comment]]
              });
            }
          } else if (this.projectName === 'INCENT_TRAFFIC') {
            // Для INCENT_TRAFFIC: WEEK внутри CAMPAIGN внутри COUNTRY внутри NETWORK
            currentWeek = nameOrRange;
            
            // Применяем комментарий к WEEK если есть
            const key = this.getCommentKey(currentNetwork, currentWeek, 'WEEK', currentCampaign, currentCampaign, currentCampaign, currentCountry);
            const comment = comments[key];
            
            if (comment) {
              updates.push({
                range: `${this.config.SHEET_NAME}!${String.fromCharCode(64 + cols.comment)}${i + 2}`,
                values: [[comment]]
              });
            }
          } else {
            currentWeek = nameOrRange;
          }
        }
        
        if (!currentApp && !(this.projectName === 'INCENT_TRAFFIC' && currentNetwork)) return;
        
        // Генерация ключей для остальных уровней
        const keys = {
          WEEK: () => {
            if (this.projectName === 'APPLOVIN_TEST' || this.projectName === 'INCENT_TRAFFIC') {
              // Уже обработано выше
              return null;
            }
            return this.getCommentKey(currentApp, currentWeek, 'WEEK', 'N/A', 'N/A', 'N/A');
          },
          SOURCE_APP: () => this.getCommentKey(currentApp, currentWeek, 'SOURCE_APP', nameOrRange, nameOrRange, 'N/A'),
          CAMPAIGN: () => {
            if (this.projectName === 'APPLOVIN_TEST' || this.projectName === 'INCENT_TRAFFIC') {
              // Уже обработано выше
              return null;
            }
            const id = this.extractCampaignIdFromHyperlink(idOrEmpty) || idOrEmpty;
            const name = this.projectName === 'TRICKY' ? id : nameOrRange;
            return this.getCommentKey(currentApp, currentWeek, 'CAMPAIGN', 
                                     this.projectName === 'TRICKY' ? id : 'N/A', nameOrRange, name);
          },
          NETWORK: () => {
            if (this.projectName === 'INCENT_TRAFFIC') {
              // Уже обработано выше
              return null;
            }
            return this.getCommentKey(currentApp, currentWeek, 'NETWORK', idOrEmpty || 'N/A', 'N/A', nameOrRange);
          },
          COUNTRY: () => {
            if (this.projectName === 'INCENT_TRAFFIC') {
              // Уже обработано выше
              return null;
            } else if (this.projectName === 'APPLOVIN_TEST') {
              // Для APPLOVIN_TEST: страны связаны с кампаниями и неделями
              const countryCode = idOrEmpty || 'N/A';
              const countryName = nameOrRange || 'N/A';
              return this.getCommentKey(currentApp, currentWeek, 'COUNTRY', 
                                       `${currentCampaign}_${countryCode}`, currentCampaign, countryName);
            }
            return this.getCommentKey(currentApp, currentWeek, 'COUNTRY', idOrEmpty || 'N/A', nameOrRange, 'N/A', nameOrRange);
          }
        };
        
        const key = keys[level]?.();
        if (key) {
          const comment = comments[key];
          
          if (comment) {
            updates.push({
              range: `${this.config.SHEET_NAME}!${String.fromCharCode(64 + cols.comment)}${i + 2}`,
              values: [[comment]]
            });
          }
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

  // Утилиты
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