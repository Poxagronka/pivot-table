/**
 * Apps Database Cache Management - ОПТИМИЗИРОВАНО: кеш 1 час + Sheets API v4 + batch операции
 */

class AppsDatabase {
  constructor(projectName = null) {
    this.projectName = projectName || CURRENT_PROJECT;
    this.config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
    
    if (this.projectName !== 'TRICKY') {
      this.cacheSheet = null;
      return;
    }
    
    this.cacheSpreadsheetId = MAIN_SHEET_ID;
    this.cacheSheetName = this.config.APPS_CACHE_SHEET;
    this.memoryCache = null;
    this.memoryCacheTime = null;
    this.CACHE_DURATION = 3600000;
  }

  getOrCreateCacheSheet() {
    if (!this.config.APPS_CACHE_SHEET) return null;
    
    try {
      const spreadsheet = Sheets.Spreadsheets.get(this.cacheSpreadsheetId);
      let sheet = spreadsheet.sheets.find(s => s.properties.title === this.cacheSheetName);
      
      if (!sheet) {
        const addSheetRequest = {
          requests: [{
            addSheet: {
              properties: {
                title: this.cacheSheetName
              }
            }
          }]
        };
        
        const response = Sheets.Spreadsheets.batchUpdate(addSheetRequest, this.cacheSpreadsheetId);
        const newSheet = response.replies[0].addSheet;
        
        const headerRequest = {
          requests: [{
            updateCells: {
              range: {
                sheetId: newSheet.properties.sheetId,
                startRowIndex: 0,
                endRowIndex: 1,
                startColumnIndex: 0,
                endColumnIndex: 5
              },
              rows: [{
                values: [
                  { userEnteredValue: { stringValue: 'Bundle ID' } },
                  { userEnteredValue: { stringValue: 'Publisher' } },
                  { userEnteredValue: { stringValue: 'App Name' } },
                  { userEnteredValue: { stringValue: 'Link App' } },
                  { userEnteredValue: { stringValue: 'Last Updated' } }
                ]
              }],
              fields: 'userEnteredValue'
            }
          }]
        };
        
        Sheets.Spreadsheets.batchUpdate(headerRequest, this.cacheSpreadsheetId);
        return { sheetId: newSheet.properties.sheetId, title: this.cacheSheetName };
      } else {
        return { sheetId: sheet.properties.sheetId, title: sheet.properties.title };
      }
    } catch (e) {
      logError('Error creating/accessing cache sheet:', e);
      throw e;
    }
  }

  loadFromCache() {
    const now = new Date().getTime();
    
    if (this.memoryCache && this.memoryCacheTime && (now - this.memoryCacheTime) < this.CACHE_DURATION) {
      logDebugInfo('Apps Database: Using memory cache');
      return this.memoryCache;
    }
    
    if (!this.cacheSheetName) return {};
    
    const apps = {};
    
    try {
      const range = `${this.cacheSheetName}!A:E`;
      const response = Sheets.Spreadsheets.Values.get(this.cacheSpreadsheetId, range);
      
      if (!response.values || response.values.length <= 1) {
        this.memoryCache = apps;
        this.memoryCacheTime = now;
        return apps;
      }
      
      for (let i = 1; i < response.values.length; i++) {
        const [bundleId, publisher, appName, linkApp, lastUpdated] = response.values[i];
        if (bundleId) {
          apps[bundleId] = {
            publisher: publisher || 'Unknown Publisher',
            appName: appName || 'Unknown App',
            linkApp: linkApp || '',
            lastUpdated: lastUpdated
          };
        }
      }
      
      this.memoryCache = apps;
      this.memoryCacheTime = now;
      
      logInfo(`Apps Database: Loaded ${Object.keys(apps).length} apps into cache`);
    } catch (e) {
      logError('Error loading Apps Database cache:', e);
    }
    
    return apps;
  }

  updateCacheFromExternalTable() {
    if (!this.cacheSheetName) return false;
    
    try {
      logInfo('Apps Database: Starting cache update from external table');
      
      const externalRange = `${APPS_DATABASE_SHEET}!A:Z`;
      const externalResponse = Sheets.Spreadsheets.Values.get(APPS_DATABASE_ID, externalRange);
      
      if (!externalResponse.values || externalResponse.values.length < 2) {
        logError('Apps Database: No data in external table');
        return false;
      }
      
      const externalData = externalResponse.values;
      const headers = externalData[0];
      const linkAppCol = this.findColumnIndex(headers, ['Link App', 'link_app', 'linkapp', 'links']);
      const publisherCol = this.findColumnIndex(headers, ['Publisher', 'publisher']);
      const appNameCol = this.findColumnIndex(headers, ['App Name', 'app_name', 'appname']);
      
      if (linkAppCol === -1) {
        logError('Apps Database: Link App column not found');
        return false;
      }
      
      this.getOrCreateCacheSheet();
      
      const clearRequest = {
        requests: [{
          updateCells: {
            range: {
              sheetId: null,
              startRowIndex: 1,
              endRowIndex: 100000,
              startColumnIndex: 0,
              endColumnIndex: 5
            },
            fields: 'userEnteredValue'
          }
        }]
      };
      
      try {
        const spreadsheet = Sheets.Spreadsheets.get(this.cacheSpreadsheetId);
        const sheet = spreadsheet.sheets.find(s => s.properties.title === this.cacheSheetName);
        if (sheet) {
          clearRequest.requests[0].updateCells.range.sheetId = sheet.properties.sheetId;
          Sheets.Spreadsheets.batchUpdate(clearRequest, this.cacheSpreadsheetId);
        }
      } catch (e) {
        logError('Error clearing cache sheet:', e);
      }
      
      const cacheData = [];
      const now = new Date();
      let lastPublisher = '';
      let lastAppName = '';
      
      for (let i = 1; i < externalData.length; i++) {
        const row = externalData[i];
        const linkApp = row[linkAppCol];
        
        if (linkApp && linkApp.trim && linkApp.trim()) {
          const bundleId = this.extractBundleIdFromLink(linkApp);
          
          if (bundleId) {
            let publisherRaw = publisherCol !== -1 ? row[publisherCol] : '';
            let appNameRaw = appNameCol !== -1 ? row[appNameCol] : '';
            
            let publisher = (publisherRaw != null && publisherRaw.toString) ? publisherRaw.toString().trim() : '';
            let appName = (appNameRaw != null && appNameRaw.toString) ? appNameRaw.toString().trim() : '';
            
            if (publisher) {
              lastPublisher = publisher;
            } else if (lastPublisher) {
              publisher = lastPublisher;
            }
            
            if (appName) {
              lastAppName = appName;
            } else if (lastAppName) {
              appName = lastAppName;
            }
            
            if (!publisher && !appName) {
              publisher = 'Unknown Publisher';
              appName = 'Unknown App';
            } else if (!publisher) {
              publisher = appName;
            } else if (!appName) {
              appName = publisher;
            }
            
            cacheData.push([bundleId, publisher, appName, linkApp, now.toISOString()]);
          }
        } else {
          let publisherRaw = publisherCol !== -1 ? row[publisherCol] : '';
          let appNameRaw = appNameCol !== -1 ? row[appNameCol] : '';
          
          let publisher = (publisherRaw != null && publisherRaw.toString) ? publisherRaw.toString().trim() : '';
          let appName = (appNameRaw != null && appNameRaw.toString) ? appNameRaw.toString().trim() : '';
          
          if (publisher) lastPublisher = publisher;
          if (appName) lastAppName = appName;
        }
      }
      
      if (cacheData.length > 0) {
        const batchSize = 1000;
        const batches = [];
        
        for (let i = 0; i < cacheData.length; i += batchSize) {
          batches.push(cacheData.slice(i, i + batchSize));
        }
        
        for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
          const batch = batches[batchIndex];
          const startRow = 1 + (batchIndex * batchSize);
          const endRow = startRow + batch.length;
          
          const batchRequest = {
            valueInputOption: 'RAW',
            data: [{
              range: `${this.cacheSheetName}!A${startRow + 1}:E${endRow}`,
              values: batch
            }]
          };
          
          Sheets.Spreadsheets.Values.batchUpdate(batchRequest, this.cacheSpreadsheetId);
          
          if (batchIndex < batches.length - 1) {
            Utilities.sleep(100);
          }
        }
        
        logInfo(`Apps Database: Cached ${cacheData.length} apps in ${batches.length} batches`);
      }
      
      this.memoryCache = null;
      this.memoryCacheTime = null;
      
      return true;
      
    } catch (e) {
      logError('Error updating Apps Database cache:', e);
      return false;
    }
  }

  extractBundleIdFromLink(linkApp) {
    if (!linkApp || typeof linkApp !== 'string') return null;
    
    const url = linkApp.trim();
    
    if (url === 'no app found' || url.length < 10 || url.includes('idx') || url.endsWith('id...')) return null;
    
    try {
      const iosPatterns = [
        /apps\.apple\.com\/.*\/app\/[^\/]*\/id(\d{8,})/,
        /apps\.apple\.com\/.*\/id(\d{8,})/,
        /apps\.apple\.com\/app\/id(\d{8,})/,
        /\/id(\d{8,})(?:[^\d]|$)/
      ];
      
      for (const pattern of iosPatterns) {
        const match = url.match(pattern);
        if (match && match[1] && match[1].length >= 8) {
          if (/^\d+$/.test(match[1])) {
            return match[1];
          }
        }
      }
      
      const androidPatterns = [
        /play\.google\.com\/store\/apps\/details\?id=([a-zA-Z][a-zA-Z0-9._]*[a-zA-Z0-9])(?:&|$)/,
        /play\.google\.com\/store\/apps\/details\?[^&]*&id=([a-zA-Z][a-zA-Z0-9._]*[a-zA-Z0-9])(?:&|$)/,
        /play\.google\.com\/store\/apps\/details\/[^?]*\?id=([a-zA-Z][a-zA-Z0-9._]*[a-zA-Z0-9])(?:&|$)/,
        /[?&]id=([a-zA-Z][a-zA-Z0-9._]*[a-zA-Z0-9])(?:&|$)/
      ];
      
      for (const pattern of androidPatterns) {
        const match = url.match(pattern);
        if (match && match[1]) {
          const bundleId = match[1];
          if (bundleId.includes('.') && bundleId.length >= 3 && /^[a-zA-Z]/.test(bundleId)) {
            return bundleId;
          }
        }
      }
      
    } catch (e) {
      logError('Error extracting bundle ID from link:', e);
    }
    
    return null;
  }

  findColumnIndex(headers, possibleNames) {
    for (let i = 0; i < headers.length; i++) {
      const header = (headers[i] || '').toString().toLowerCase().trim();
      for (const name of possibleNames) {
        if (header === name.toLowerCase()) {
          return i;
        }
      }
    }
    return -1;
  }

  getAppInfo(bundleId) {
    if (!bundleId || this.projectName !== 'TRICKY') {
      return { publisher: 'Unknown Publisher', appName: 'Unknown App', linkApp: '' };
    }
    
    const cache = this.loadFromCache();
    const appInfo = cache[bundleId];
    
    if (appInfo) {
      return {
        publisher: appInfo.publisher,
        appName: appInfo.appName,
        linkApp: appInfo.linkApp
      };
    }
    
    return {
      publisher: bundleId,
      appName: '',
      linkApp: ''
    };
  }

  getSourceAppDisplayName(bundleId) {
    if (!bundleId || this.projectName !== 'TRICKY') {
      return bundleId || 'Unknown';
    }
    
    const appInfo = this.getAppInfo(bundleId);
    
    if (appInfo.publisher !== bundleId) {
      const publisher = appInfo.publisher || '';
      const appName = appInfo.appName || '';
      
      if (publisher && appName && publisher !== appName) {
        return `${publisher} ${appName}`;
      } else if (publisher) {
        return publisher;
      } else if (appName) {
        return appName;
      }
    }
    
    return bundleId;
  }

  shouldUpdateCache() {
    if (!this.cacheSheetName || this.projectName !== 'TRICKY') return false;
    
    const cache = this.loadFromCache();
    const bundleIds = Object.keys(cache);
    
    if (bundleIds.length === 0) return true;
    
    const oneHourAgo = new Date(Date.now() - this.CACHE_DURATION);
    
    for (const bundleId of bundleIds) {
      const lastUpdated = cache[bundleId].lastUpdated;
      if (!lastUpdated || new Date(lastUpdated) < oneHourAgo) {
        return true;
      }
    }
    
    return false;
  }

  ensureCacheUpToDate() {
    if (this.projectName !== 'TRICKY') return;
    
    try {
      if (this.shouldUpdateCache()) {
        logInfo('Apps Database: Cache is outdated, updating...');
        this.updateCacheFromExternalTable();
      } else {
        logDebugInfo('Apps Database: Cache is up to date');
      }
    } catch (e) {
      logError('Error ensuring Apps Database cache is up to date:', e);
    }
  }

  preloadCache() {
    if (this.projectName !== 'TRICKY') return;
    
    logInfo('Apps Database: Preloading cache...');
    this.ensureCacheUpToDate();
    this.loadFromCache();
    logInfo('Apps Database: Cache preloaded successfully');
  }

  clearMemoryCache() {
    this.memoryCache = null;
    this.memoryCacheTime = null;
    logDebugInfo('Apps Database: Memory cache cleared');
  }
}

function extractBundleIdFromCampaign(campaignName) {
  if (!campaignName) return null;
  
  const equalsIndex = campaignName.indexOf('=');
  
  if (equalsIndex !== -1) {
    const beforeEquals = campaignName.substring(0, equalsIndex).trim();
    
    if (isValidBundleId(beforeEquals)) {
      return beforeEquals;
    }
    
    let startIndex = equalsIndex + 1;
    
    while (startIndex < campaignName.length && campaignName[startIndex] === ' ') {
      startIndex++;
    }
    
    if (startIndex >= campaignName.length) return null;
    
    const spaceIndex = campaignName.indexOf(' ', startIndex);
    let bundleId;
    
    if (spaceIndex === -1) {
      bundleId = campaignName.substring(startIndex).trim();
    } else {
      bundleId = campaignName.substring(startIndex, spaceIndex).trim();
    }
    
    if (isValidBundleId(bundleId)) {
      return bundleId;
    } else {
      return null;
    }
  }
  
  const spaceIndex = campaignName.indexOf(' ');
  
  if (spaceIndex === -1) {
    const wholeName = campaignName.trim();
    if (isValidBundleId(wholeName)) {
      return wholeName;
    }
  } else {
    const bundleId = campaignName.substring(0, spaceIndex).trim();
    
    if (isValidBundleId(bundleId)) {
      return bundleId;
    }
  }
  
  return null;
}

function isValidBundleId(bundleId) {
  if (!bundleId || typeof bundleId !== 'string') return false;
  
  const trimmed = bundleId.trim();
  if (trimmed.length < 3) return false;
  
  if (trimmed.includes('.') && /^[a-zA-Z]/.test(trimmed)) {
    if (/^[a-zA-Z][a-zA-Z0-9._]*[a-zA-Z0-9]$/.test(trimmed)) {
      return true;
    }
  }
  
  if (/^\d{8,}$/.test(trimmed)) {
    return true;
  }
  
  return false;
}

function refreshAppsDatabase() {
  const ui = SpreadsheetApp.getUi();
  
  if (CURRENT_PROJECT !== 'TRICKY') {
    ui.alert('Apps Database', 'Apps Database is only used for TRICKY project.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const appsDb = new AppsDatabase('TRICKY');
    const success = appsDb.updateCacheFromExternalTable();
    
    if (success) {
      const cache = appsDb.loadFromCache();
      const count = Object.keys(cache).length;
      ui.alert('Apps Database Updated', `Successfully updated Apps Database cache.\n\n${count} apps loaded.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Update Failed', 'Failed to update Apps Database cache.\n\nCheck console for errors.', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', 'Error updating Apps Database: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function debugAppsDatabase() {
  setLogLevel('DEBUG');
  
  logDebugInfo('=== APPS DATABASE DEBUG START ===');
  
  if (CURRENT_PROJECT !== 'TRICKY') {
    logError('Apps Database только для TRICKY проекта');
    return;
  }
  
  try {
    const appsDb = new AppsDatabase('TRICKY');
    logDebugInfo('AppsDatabase объект создан');
    
    if (!appsDb.cacheSheetName) {
      logError('Cache sheet не найден');
      return;
    }
    logDebugInfo('Cache sheet найден:', appsDb.cacheSheetName);
    
    logDebugInfo('Подключаемся к внешней таблице:', APPS_DATABASE_ID);
    
    try {
      const externalRange = `${APPS_DATABASE_SHEET}!A:Z`;
      const externalResponse = Sheets.Spreadsheets.Values.get(APPS_DATABASE_ID, externalRange);
      logDebugInfo('Внешняя таблица подключена');
      
      if (!externalResponse.values || externalResponse.values.length < 2) {
        logError('Недостаточно данных в таблице');
        return;
      }
      
      logDebugInfo('Данных получено:', externalResponse.values.length, 'строк');
      
      const headers = externalResponse.values[0];
      logDebugInfo('Заголовки:', headers);
      
      const linkAppCol = appsDb.findColumnIndex(headers, ['Link App', 'link_app', 'linkapp', 'links']);
      const publisherCol = appsDb.findColumnIndex(headers, ['Publisher', 'publisher']);
      const appNameCol = appsDb.findColumnIndex(headers, ['App Name', 'app_name', 'appname']);
      
      logDebugInfo('Найденные колонки:');
      logDebugInfo('  - Link App:', linkAppCol, linkAppCol !== -1 ? `(${headers[linkAppCol]})` : '(НЕ НАЙДЕНА)');
      logDebugInfo('  - Publisher:', publisherCol, publisherCol !== -1 ? `(${headers[publisherCol]})` : '(НЕ НАЙДЕНА)');
      logDebugInfo('  - App Name:', appNameCol, appNameCol !== -1 ? `(${headers[appNameCol]})` : '(НЕ НАЙДЕНА)');
      
      if (linkAppCol === -1) {
        logError('КРИТИЧЕСКАЯ ОШИБКА: Link App колонка не найдена!');
        return;
      }
      
      logDebugInfo('Анализ первых 5 строк данных:');
      const sampleSize = Math.min(5, externalResponse.values.length - 1);
      
      for (let i = 1; i <= sampleSize; i++) {
        const row = externalResponse.values[i];
        const linkApp = row[linkAppCol];
        const publisher = publisherCol !== -1 ? row[publisherCol] : 'Unknown Publisher';
        const appName = appNameCol !== -1 ? row[appNameCol] : 'Unknown App';
        
        logDebugInfo(`\nСтрока ${i}:`);
        logDebugInfo('  Publisher:', `"${publisher}"`, publisher ? '✅' : '❌ ПУСТОЙ');
        logDebugInfo('  App Name:', `"${appName}"`, appName ? '✅' : '❌ ПУСТОЙ');
        logDebugInfo('  Link App:', linkApp?.substring(0, 100) + (linkApp?.length > 100 ? '...' : ''));
        
        if (linkApp && linkApp.trim()) {
          const bundleId = appsDb.extractBundleIdFromLink(linkApp);
          logDebugInfo('  Extracted Bundle ID:', bundleId || 'НЕ УДАЛОСЬ ИЗВЛЕЧЬ');
        } else {
          logDebugInfo('  ⚠️ Пустая ссылка');
        }
      }
      
      logDebugInfo('Начинаем полную обработку...');
      const success = appsDb.updateCacheFromExternalTable();
      
      if (success) {
        const cache = appsDb.loadFromCache();
        const count = Object.keys(cache).length;
        logInfo('УСПЕХ! Кеш обновлен:', count, 'приложений');
      } else {
        logError('ОШИБКА: Обновление кеша провалилось');
      }
      
    } catch (e) {
      logError('Ошибка работы с внешней таблицей:', e);
    }
    
  } catch (e) {
    logError('ИСКЛЮЧЕНИЕ:', e.toString());
  }
  
  logDebugInfo('=== APPS DATABASE DEBUG END ===');
  setLogLevel('INFO');
}

function testAppsDbBundleExtraction() {
  setLogLevel('DEBUG');
  logDebugInfo('=== TESTING BUNDLE ID EXTRACTION AND APPS DB LOOKUP ===');
  
  setCurrentProject('TRICKY');
  logDebugInfo('Project set to:', CURRENT_PROJECT);
  
  const testCampaigns = [
    '[pb tricky] | NPCW | USA | bm | I | subject = words.puzzle.wordgame.free.connect Bidmachine AMBO CPI skipctr',
    '[pb tricky] | NPCW | GBR | bm | I | subject = com.fanatee.cody Bidmachine skipctr AMBO CPI',
    '[pb tricky] | NPCNM | CAN | bm | I | subject = 6473832648 AMBO CPI abdoul Bidmachine skipctr autobudget',
    '[pb tricky] | NPKS | USA | subj=com.intensedev.classicsolitaire.gp Bidmachine AMBO CPA skipctr',
    'com.easybrain.number.puzzle.game=80 Pub/Smaato СPP_nc_easybrain'
  ];
  
  logDebugInfo('TESTING CAMPAIGN NAME PROCESSING:');
  
  testCampaigns.forEach((campaignName, index) => {
    logDebugInfo(`\n--- Test ${index + 1}: "${campaignName}" ---`);
    
    const bundleId = extractBundleIdFromCampaign(campaignName);
    logDebugInfo(`Extracted bundle ID: "${bundleId}"`);
    
    if (bundleId) {
      const appsDb = new AppsDatabase('TRICKY');
      logDebugInfo('Apps Database instance created');
      
      const displayName = appsDb.getSourceAppDisplayName(bundleId);
      logDebugInfo(`Final display name: "${displayName}"`);
      
      const cache = appsDb.loadFromCache();
      const cacheEntry = cache[bundleId];
      logDebugInfo(`Cache entry for "${bundleId}":`, cacheEntry);
      
    } else {
      logDebugInfo('❌ Bundle ID extraction failed');
    }
  });
  
  try {
    const appsDb = new AppsDatabase('TRICKY');
    const cache = appsDb.loadFromCache();
    const cacheSize = Object.keys(cache).length;
    logDebugInfo(`Apps Database cache size: ${cacheSize} entries`);
    
    if (cacheSize > 0) {
      const sampleKeys = Object.keys(cache).slice(0, 5);
      logDebugInfo('Sample cache entries:');
      sampleKeys.forEach(key => {
        const entry = cache[key];
        logDebugInfo(`  ${key} → ${entry.publisher} ${entry.appName}`);
      });
    }
  } catch (e) {
    logError('Error loading cache:', e);
  }
  
  logDebugInfo('=== TESTING COMPLETE ===');
  setLogLevel('INFO');
}