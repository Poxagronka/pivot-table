class AppsDatabase {
  constructor(projectName = null) {
    this.projectName = projectName || CURRENT_PROJECT;
    this.config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
    
    if (this.projectName !== 'TRICKY') {
      this.cacheSheet = null;
      return;
    }
    
    this.cacheSheet = this.getOrCreateCacheSheet();
    this.optimizedCache = null;
  }

  getOrCreateCacheSheet() {
    if (!this.config.APPS_CACHE_SHEET) return null;
    
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    let sheet = spreadsheet.getSheetByName(this.config.APPS_CACHE_SHEET);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(this.config.APPS_CACHE_SHEET);
      sheet.hideSheet();
      sheet.getRange(1, 1, 1, 5).setValues([['Bundle ID', 'Publisher', 'App Name', 'Link App', 'Last Updated']]);
    }
    return sheet;
  }

  loadFromCache() {
    if (!this.cacheSheet) return {};
    
    if (this.optimizedCache) {
      return this.optimizedCache;
    }
    
    console.log('Loading Apps Database from cache...');
    const apps = {};
    const data = this.cacheSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const [bundleId, publisher, appName, linkApp, lastUpdated] = data[i];
      if (bundleId) {
        apps[bundleId] = {
          publisher: publisher || 'Unknown Publisher',
          appName: appName || 'Unknown App',
          linkApp: linkApp || '',
          lastUpdated: lastUpdated
        };
      }
    }
    
    this.optimizedCache = apps;
    console.log(`Apps Database cached: ${Object.keys(apps).length} apps`);
    return apps;
  }

  loadFromCacheOptimized() {
    return this.loadFromCache();
  }

  updateCacheFromExternalTable() {
    if (!this.cacheSheet) return false;
    
    console.log('Updating Apps Database from external table...');
    
    try {
      const externalSpreadsheet = SpreadsheetApp.openById(APPS_DATABASE_ID);
      const externalSheet = externalSpreadsheet.getSheetByName(APPS_DATABASE_SHEET);
      
      if (!externalSheet) {
        console.log('External sheet not found');
        return false;
      }
      
      const externalData = externalSheet.getDataRange().getValues();
      if (externalData.length < 2) {
        console.log('Insufficient data in external sheet');
        return false;
      }
      
      const headers = externalData[0];
      const linkAppCol = this.findColumnIndex(headers, ['Link App', 'link_app', 'linkapp', 'links']);
      const publisherCol = this.findColumnIndex(headers, ['Publisher', 'publisher']);
      const appNameCol = this.findColumnIndex(headers, ['App Name', 'app_name', 'appname']);
      
      if (linkAppCol === -1) {
        console.log('Link App column not found');
        return false;
      }
      
      console.log(`Found columns - Link App: ${linkAppCol}, Publisher: ${publisherCol}, App Name: ${appNameCol}`);
      
      if (this.cacheSheet.getLastRow() > 1) {
        this.cacheSheet.deleteRows(2, this.cacheSheet.getLastRow() - 1);
      }
      
      const cacheData = [];
      const now = new Date();
      let lastPublisher = '';
      let lastAppName = '';
      let processedCount = 0;
      let bundleIdCount = 0;
      
      console.log(`Processing ${externalData.length - 1} rows...`);
      
      for (let i = 1; i < externalData.length; i++) {
        const row = externalData[i];
        const linkApp = row[linkAppCol];
        
        if (linkApp && linkApp.trim && linkApp.trim()) {
          const bundleId = this.extractBundleIdFromLinkOptimized(linkApp);
          
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
            
            cacheData.push([bundleId, publisher, appName, linkApp, now]);
            bundleIdCount++;
          }
        } else {
          let publisherRaw = publisherCol !== -1 ? row[publisherCol] : '';
          let appNameRaw = appNameCol !== -1 ? row[appNameCol] : '';
          
          let publisher = (publisherRaw != null && publisherRaw.toString) ? publisherRaw.toString().trim() : '';
          let appName = (appNameRaw != null && appNameRaw.toString) ? appNameRaw.toString().trim() : '';
          
          if (publisher) lastPublisher = publisher;
          if (appName) lastAppName = appName;
        }
        
        processedCount++;
        if (processedCount % 100 === 0) {
          console.log(`Processed ${processedCount}/${externalData.length - 1} rows, found ${bundleIdCount} bundle IDs`);
        }
      }
      
      if (cacheData.length > 0) {
        console.log(`Writing ${cacheData.length} entries to cache...`);
        const batchSize = 100;
        
        for (let i = 0; i < cacheData.length; i += batchSize) {
          const batch = cacheData.slice(i, i + batchSize);
          const lastRow = this.cacheSheet.getLastRow();
          this.cacheSheet.getRange(lastRow + 1, 1, batch.length, 5).setValues(batch);
          
          if (i + batchSize < cacheData.length) {
            Utilities.sleep(100);
          }
        }
      }
      
      this.optimizedCache = null;
      
      console.log(`Apps Database update completed: ${bundleIdCount} apps processed`);
      return true;
      
    } catch (e) {
      console.error('Error updating Apps Database:', e);
      return false;
    }
  }

  extractBundleIdFromLinkOptimized(linkApp) {
    if (!linkApp || typeof linkApp !== 'string') return null;
    
    const url = linkApp.trim();
    
    if (url === 'no app found' || url.length < 10 || url.includes('idx') || url.endsWith('id...')) return null;
    
    try {
      if (url.includes('apps.apple.com')) {
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
      }
      
      if (url.includes('play.google.com')) {
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
      }
      
    } catch (e) {
      // Silent fail
    }
    
    return null;
  }

  extractBundleIdFromLink(linkApp) {
    return this.extractBundleIdFromLinkOptimized(linkApp);
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

  getAppInfoOptimized(bundleId, cacheData = null) {
    if (!bundleId || this.projectName !== 'TRICKY') {
      return { publisher: 'Unknown Publisher', appName: 'Unknown App', linkApp: '' };
    }
    
    const cache = cacheData || this.loadFromCache();
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

  getSourceAppDisplayNameOptimized(bundleId, cacheData = null) {
    if (!bundleId || this.projectName !== 'TRICKY') {
      return bundleId || 'Unknown';
    }
    
    const appInfo = this.getAppInfoOptimized(bundleId, cacheData);
    
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
    if (!this.cacheSheet || this.projectName !== 'TRICKY') return false;
    
    const cache = this.loadFromCache();
    const bundleIds = Object.keys(cache);
    
    if (bundleIds.length === 0) return true;
    
    const oneDayAgo = new Date(Date.now() - 24 * 60 * 60 * 1000);
    
    for (const bundleId of bundleIds) {
      const lastUpdated = cache[bundleId].lastUpdated;
      if (!lastUpdated || new Date(lastUpdated) < oneDayAgo) {
        return true;
      }
    }
    
    return false;
  }

  ensureCacheUpToDate() {
    if (this.projectName !== 'TRICKY') return;
    
    try {
      if (this.shouldUpdateCache()) {
        console.log('Apps Database cache is outdated, updating...');
        this.updateCacheFromExternalTable();
      } else {
        console.log('Apps Database cache is up to date');
      }
    } catch (e) {
      console.error('Error ensuring cache is up to date:', e);
    }
  }

  clearOptimizedCache() {
    this.optimizedCache = null;
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
  console.log('=== APPS DATABASE DEBUG START ===');
  
  if (CURRENT_PROJECT !== 'TRICKY') {
    console.log('ОШИБКА: Apps Database только для TRICKY проекта');
    return;
  }
  
  try {
    const appsDb = new AppsDatabase('TRICKY');
    console.log('✅ AppsDatabase объект создан');
    
    if (!appsDb.cacheSheet) {
      console.log('❌ Cache sheet не найден');
      return;
    }
    console.log('✅ Cache sheet найден:', appsDb.config.APPS_CACHE_SHEET);
    
    console.log('📋 Подключаемся к внешней таблице:', APPS_DATABASE_ID);
    const externalSpreadsheet = SpreadsheetApp.openById(APPS_DATABASE_ID);
    console.log('✅ Внешняя таблица подключена');
    
    const externalSheet = externalSpreadsheet.getSheetByName(APPS_DATABASE_SHEET);
    if (!externalSheet) {
      console.log('❌ Лист не найден:', APPS_DATABASE_SHEET);
      return;
    }
    console.log('✅ Лист найден:', APPS_DATABASE_SHEET);
    
    const externalData = externalSheet.getDataRange().getValues();
    console.log('📊 Данных получено:', externalData.length, 'строк');
    
    if (externalData.length < 2) {
      console.log('❌ Недостаточно данных в таблице');
      return;
    }
    
    const headers = externalData[0];
    console.log('📋 Заголовки:', headers);
    
    const linkAppCol = appsDb.findColumnIndex(headers, ['Link App', 'link_app', 'linkapp', 'links']);
    const publisherCol = appsDb.findColumnIndex(headers, ['Publisher', 'publisher']);
    const appNameCol = appsDb.findColumnIndex(headers, ['App Name', 'app_name', 'appname']);
    
    console.log('🔍 Найденные колонки:');
    console.log('  - Link App:', linkAppCol, linkAppCol !== -1 ? `(${headers[linkAppCol]})` : '(НЕ НАЙДЕНА)');
    console.log('  - Publisher:', publisherCol, publisherCol !== -1 ? `(${headers[publisherCol]})` : '(НЕ НАЙДЕНА)');
    console.log('  - App Name:', appNameCol, appNameCol !== -1 ? `(${headers[appNameCol]})` : '(НЕ НАЙДЕНА)');
    
    if (linkAppCol === -1) {
      console.log('❌ КРИТИЧЕСКАЯ ОШИБКА: Link App колонка не найдена!');
      console.log('🔍 Возможные варианты названий:', ['Link App', 'link_app', 'linkapp', 'links']);
      return;
    }
    
    console.log('🔬 Анализ первых 5 строк данных:');
    const sampleSize = Math.min(5, externalData.length - 1);
    
    for (let i = 1; i <= sampleSize; i++) {
      const row = externalData[i];
      const linkApp = row[linkAppCol];
      const publisher = publisherCol !== -1 ? row[publisherCol] : 'Unknown Publisher';
      const appName = appNameCol !== -1 ? row[appNameCol] : 'Unknown App';
      
      console.log(`\n📱 Строка ${i}:`);
      console.log('  Publisher:', `"${publisher}"`, publisher ? '✅' : '❌ ПУСТОЙ');
      console.log('  App Name:', `"${appName}"`, appName ? '✅' : '❌ ПУСТОЙ');
      console.log('  Link App:', linkApp?.substring(0, 100) + (linkApp?.length > 100 ? '...' : ''));
      
      if (linkApp && linkApp.trim()) {
        const bundleId = appsDb.extractBundleIdFromLinkOptimized(linkApp);
        console.log('  Extracted Bundle ID:', bundleId || 'НЕ УДАЛОСЬ ИЗВЛЕЧЬ');
        
        if (bundleId) {
          let processedPublisher = (publisher && publisher.trim()) ? publisher.trim() : '';
          let processedAppName = (appName && appName.trim()) ? appName.trim() : '';
          
          if (!processedPublisher && !processedAppName) {
            processedPublisher = 'Unknown Publisher';
            processedAppName = 'Unknown App';
          } else if (!processedPublisher) {
            processedPublisher = processedAppName;
          } else if (!processedAppName) {
            processedAppName = processedPublisher;
          }
          
          console.log('  Final Display:', `${processedPublisher} ${processedAppName}`);
        } else {
          console.log('  🔍 Анализ ссылки:');
          console.log('    - Содержит apps.apple.com:', linkApp.includes('apps.apple.com'));
          console.log('    - Содержит play.google.com:', linkApp.includes('play.google.com'));
          console.log('    - iOS /id pattern:', /\/id(\d{8,})/.test(linkApp));
          console.log('    - Android id= pattern:', /[?&]id=([a-zA-Z][a-zA-Z0-9._]+)/.test(linkApp));
          console.log('    - Длина ссылки:', linkApp.length);
        }
      } else {
        console.log('  ⚠️ Пустая ссылка');
      }
    }
    
    console.log('\n🚀 Начинаем полную обработку...');
    const success = appsDb.updateCacheFromExternalTable();
    
    if (success) {
      const cache = appsDb.loadFromCache();
      const count = Object.keys(cache).length;
      console.log('✅ УСПЕХ! Кеш обновлен:', count, 'приложений');
      
      console.log('\n📊 Примеры кешированных данных:');
      const cacheKeys = Object.keys(cache).slice(0, 5);
      let emptyPublisherCount = 0;
      let emptyAppNameCount = 0;
      
      cacheKeys.forEach(bundleId => {
        const app = cache[bundleId];
        console.log(`  ${bundleId} → ${app.publisher} ${app.appName}`);
        
        if (app.publisher === 'Unknown Publisher') emptyPublisherCount++;
        if (app.appName === 'Unknown App') emptyAppNameCount++;
      });
      
      console.log('\n📈 СТАТИСТИКА КАЧЕСТВА ДАННЫХ:');
      console.log(`  - Всего в кеше: ${count} приложений`);
      console.log(`  - С пустыми Publisher: ~${Math.round(emptyPublisherCount * count / cacheKeys.length)}`);
      console.log(`  - С пустыми App Name: ~${Math.round(emptyAppNameCount * count / cacheKeys.length)}`);
      console.log(`  - Поддержка оптимизированной обработки: ✅ Включена`);
    } else {
      console.log('❌ ОШИБКА: Обновление кеша провалилось');
    }
    
  } catch (e) {
    console.log('💥 ИСКЛЮЧЕНИЕ:', e.toString());
    console.log('📋 Stack trace:', e.stack);
  }
  
  console.log('\n=== APPS DATABASE DEBUG END ===');
}

function testAppsDbBundleExtraction() {
  console.log('=== TESTING BUNDLE ID EXTRACTION AND APPS DB LOOKUP ===');
  
  setCurrentProject('TRICKY');
  console.log('Project set to:', CURRENT_PROJECT);
  
  const testCampaigns = [
    '[pb tricky] | NPCW | USA | bm | I | subject = words.puzzle.wordgame.free.connect Bidmachine AMBO CPI skipctr',
    '[pb tricky] | NPCW | GBR | bm | I | subject = com.fanatee.cody Bidmachine skipctr AMBO CPI',
    '[pb tricky] | NPCNM | CAN | bm | I | subject = 6473832648 AMBO CPI abdoul Bidmachine skipctr autobudget',
    '[pb tricky] | NPKS | USA | subj=com.intensedev.classicsolitaire.gp Bidmachine AMBO CPA skipctr',
    'com.easybrain.number.puzzle.game=80 Pub/Smaato СPP_nc_easybrain'
  ];
  
  console.log('\n🧪 TESTING OPTIMIZED CAMPAIGN NAME PROCESSING:');
  
  testCampaigns.forEach((campaignName, index) => {
    console.log(`\n--- Test ${index + 1}: "${campaignName}" ---`);
    
    const bundleId = extractBundleIdFromCampaign(campaignName);
    console.log(`Extracted bundle ID: "${bundleId}"`);
    
    if (bundleId) {
      const appsDb = new AppsDatabase('TRICKY');
      console.log('Apps Database instance created');
      
      const displayName = appsDb.getSourceAppDisplayNameOptimized(bundleId);
      console.log(`Final display name: "${displayName}"`);
      
      const cache = appsDb.loadFromCache();
      const cacheEntry = cache[bundleId];
      console.log(`Cache entry for "${bundleId}":`, cacheEntry);
      
    } else {
      console.log('❌ Bundle ID extraction failed');
    }
  });
  
  try {
    const appsDb = new AppsDatabase('TRICKY');
    const cache = appsDb.loadFromCache();
    const cacheSize = Object.keys(cache).length;
    console.log(`\n📊 Apps Database cache size: ${cacheSize} entries`);
    
    if (cacheSize > 0) {
      const sampleKeys = Object.keys(cache).slice(0, 5);
      console.log('📋 Sample optimized cache entries:');
      sampleKeys.forEach(key => {
        const entry = cache[key];
        console.log(`  ${key} → ${entry.publisher} ${entry.appName}`);
      });
    }
  } catch (e) {
    console.log('❌ Error loading cache:', e);
  }
  
  console.log('\n=== TESTING COMPLETE ===');
}