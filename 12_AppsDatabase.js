/**
 * Apps Database Cache Management - ИСПРАВЛЕНО: правильное извлечение bundle ID из названия кампании
 * Кеширует данные из внешней таблицы Apps Database для группировки по Publisher + App Name
 */

class AppsDatabase {
  constructor(projectName = null) {
    this.projectName = projectName || CURRENT_PROJECT;
    this.config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
    
    // Only TRICKY project uses Apps Database
    if (this.projectName !== 'TRICKY') {
      this.cacheSheet = null;
      return;
    }
    
    this.cacheSheet = this.getOrCreateCacheSheet();
  }

  /**
   * Get or create the apps cache sheet
   */
  getOrCreateCacheSheet() {
    if (!this.config.APPS_CACHE_SHEET) return null;
    
    const spreadsheet = SpreadsheetApp.openById(this.config.SHEET_ID);
    let sheet = spreadsheet.getSheetByName(this.config.APPS_CACHE_SHEET);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(this.config.APPS_CACHE_SHEET);
      sheet.hideSheet();
      // Headers: Bundle ID, Publisher, App Name, Link App, Last Updated
      sheet.getRange(1, 1, 1, 5).setValues([['Bundle ID', 'Publisher', 'App Name', 'Link App', 'Last Updated']]);
    }
    return sheet;
  }

  /**
   * Load apps data from cache
   */
  loadFromCache() {
    if (!this.cacheSheet) return {};
    
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
    
    return apps;
  }

  /**
   * Update cache from external Apps Database table
   */
  updateCacheFromExternalTable() {
    if (!this.cacheSheet) return false;
    
    try {
      console.log('Updating Apps Database cache from external table...');
      
      // Access external spreadsheet
      const externalSpreadsheet = SpreadsheetApp.openById(APPS_DATABASE_ID);
      const externalSheet = externalSpreadsheet.getSheetByName(APPS_DATABASE_SHEET);
      
      if (!externalSheet) {
        console.error('Apps Database sheet not found');
        return false;
      }
      
      const externalData = externalSheet.getDataRange().getValues();
      if (externalData.length < 2) {
        console.error('No data in Apps Database sheet');
        return false;
      }
      
      // Find column indices
      const headers = externalData[0];
      const linkAppCol = this.findColumnIndex(headers, ['Link App', 'link_app', 'linkapp', 'links']);
      const publisherCol = this.findColumnIndex(headers, ['Publisher', 'publisher']);
      const appNameCol = this.findColumnIndex(headers, ['App Name', 'app_name', 'appname']);
      
      if (linkAppCol === -1) {
        console.error('Link App column not found in Apps Database');
        console.log('Available headers:', headers);
        return false;
      }
      
      // Clear cache sheet except headers
      if (this.cacheSheet.getLastRow() > 1) {
        this.cacheSheet.deleteRows(2, this.cacheSheet.getLastRow() - 1);
      }
      
      // Process external data with merged cells support
      const cacheData = [];
      const now = new Date();
      let processedApps = 0;
      let skippedApps = 0;
      
      // Track last non-empty values for merged cells
      let lastPublisher = '';
      let lastAppName = '';
      
      for (let i = 1; i < externalData.length; i++) {
        const row = externalData[i];
        const linkApp = row[linkAppCol];
        
        if (linkApp && linkApp.trim && linkApp.trim()) {
          const bundleId = this.extractBundleIdFromLink(linkApp);
          
          if (bundleId) {
            // Get raw values
            let publisherRaw = publisherCol !== -1 ? row[publisherCol] : '';
            let appNameRaw = appNameCol !== -1 ? row[appNameCol] : '';
            
            // Convert to strings safely
            let publisher = (publisherRaw != null && publisherRaw.toString) ? publisherRaw.toString().trim() : '';
            let appName = (appNameRaw != null && appNameRaw.toString) ? appNameRaw.toString().trim() : '';
            
            // Handle merged cells - use last non-empty value if current is empty
            if (publisher) {
              lastPublisher = publisher; // Update last known publisher
            } else if (lastPublisher) {
              publisher = lastPublisher; // Use last known publisher for merged cell
            }
            
            if (appName) {
              lastAppName = appName; // Update last known app name
            } else if (lastAppName) {
              appName = lastAppName; // Use last known app name for merged cell
            }
            
            // Use fallbacks for completely empty values
            if (!publisher && !appName) {
              publisher = 'Unknown Publisher';
              appName = 'Unknown App';
            } else if (!publisher) {
              publisher = appName;
            } else if (!appName) {
              appName = publisher;
            }
            
            cacheData.push([bundleId, publisher, appName, linkApp, now]);
            processedApps++;
          } else {
            skippedApps++;
            console.log(`Could not extract bundle ID from: ${linkApp.substring(0, 100)}...`);
          }
        } else {
          skippedApps++;
          
          // Still need to update merged cell tracking even for rows without links
          let publisherRaw = publisherCol !== -1 ? row[publisherCol] : '';
          let appNameRaw = appNameCol !== -1 ? row[appNameCol] : '';
          
          let publisher = (publisherRaw != null && publisherRaw.toString) ? publisherRaw.toString().trim() : '';
          let appName = (appNameRaw != null && appNameRaw.toString) ? appNameRaw.toString().trim() : '';
          
          if (publisher) lastPublisher = publisher;
          if (appName) lastAppName = appName;
        }
      }
      
      // Write to cache
      if (cacheData.length > 0) {
        const lastRow = this.cacheSheet.getLastRow();
        this.cacheSheet.getRange(lastRow + 1, 1, cacheData.length, 5).setValues(cacheData);
      }
      
      console.log(`Apps Database cache updated: ${processedApps} apps processed, ${skippedApps} skipped`);
      console.log(`Merged cells handling: Publisher "${lastPublisher}", App Name "${lastAppName}"`);
      return true;
      
    } catch (e) {
      console.error('Error updating Apps Database cache:', e);
      return false;
    }
  }

  /**
   * Extract bundle ID from a single store link
   */
  extractBundleIdFromLink(linkApp) {
    if (!linkApp || typeof linkApp !== 'string') return null;
    
    const url = linkApp.trim();
    
    // Skip invalid links
    if (url === 'no app found' || url.length < 10 || url.includes('idx') || url.endsWith('id...')) return null;
    
    try {
      // iOS App Store patterns (multiple variations)
      const iosPatterns = [
        /apps\.apple\.com\/.*\/app\/[^\/]*\/id(\d{8,})/,  // Standard format
        /apps\.apple\.com\/.*\/id(\d{8,})/,               // Short format  
        /apps\.apple\.com\/app\/id(\d{8,})/,              // Direct format
        /\/id(\d{8,})(?:[^\d]|$)/                         // Any /id followed by 8+ digits
      ];
      
      for (const pattern of iosPatterns) {
        const match = url.match(pattern);
        if (match && match[1] && match[1].length >= 8) {
          // Additional validation: iOS IDs should be numeric only
          if (/^\d+$/.test(match[1])) {
            return match[1];
          }
        }
      }
      
      // Google Play Store patterns (multiple parameter orders)
      const androidPatterns = [
        /play\.google\.com\/store\/apps\/details\?id=([a-zA-Z][a-zA-Z0-9._]*[a-zA-Z0-9])(?:&|$)/,  // id first
        /play\.google\.com\/store\/apps\/details\?[^&]*&id=([a-zA-Z][a-zA-Z0-9._]*[a-zA-Z0-9])(?:&|$)/, // id not first
        /play\.google\.com\/store\/apps\/details\/[^?]*\?id=([a-zA-Z][a-zA-Z0-9._]*[a-zA-Z0-9])(?:&|$)/, // with path
        /[?&]id=([a-zA-Z][a-zA-Z0-9._]*[a-zA-Z0-9])(?:&|$)/  // any id parameter
      ];
      
      for (const pattern of androidPatterns) {
        const match = url.match(pattern);
        if (match && match[1]) {
          const bundleId = match[1];
          // Validate Android bundle ID format
          if (bundleId.includes('.') && bundleId.length >= 3 && /^[a-zA-Z]/.test(bundleId)) {
            return bundleId;
          }
        }
      }
      
    } catch (e) {
      console.error('Error extracting bundle ID from link:', e);
    }
    
    return null;
  }

  /**
   * Find column index by possible header names
   */
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

  /**
   * Get app info by bundle ID
   */
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
    
    // If not found, return bundle ID as fallback
    return {
      publisher: bundleId,
      appName: '',
      linkApp: ''
    };
  }

  /**
   * Get source app display name (Publisher + App Name or bundle ID)
   */
  getSourceAppDisplayName(bundleId) {
    // Debug logging
    console.log(`getSourceAppDisplayName called with bundleId: "${bundleId}", type: ${typeof bundleId}, project: ${this.projectName}`);
    
    if (!bundleId || this.projectName !== 'TRICKY') {
      console.log(`Returning early: bundleId="${bundleId}", project="${this.projectName}"`);
      return bundleId || 'Unknown';
    }
    
    const appInfo = this.getAppInfo(bundleId);
    console.log(`Apps Database lookup result:`, appInfo);
    
    // If we found data in database
    if (appInfo.publisher !== bundleId) {
      const publisher = appInfo.publisher || '';
      const appName = appInfo.appName || '';
      
      console.log(`Found in database - Publisher: "${publisher}", App Name: "${appName}"`);
      
      // Create display name based on available data
      if (publisher && appName && publisher !== appName) {
        const result = `${publisher} ${appName}`;
        console.log(`Returning combined: "${result}"`);
        return result;
      } else if (publisher) {
        console.log(`Returning publisher: "${publisher}"`);
        return publisher;
      } else if (appName) {
        console.log(`Returning app name: "${appName}"`);
        return appName;
      }
    }
    
    // Fallback to bundle ID
    console.log(`Fallback to bundle ID: "${bundleId}"`);
    return bundleId;
  }

  /**
   * Check if cache needs update (older than 24 hours)
   */
  shouldUpdateCache() {
    if (!this.cacheSheet || this.projectName !== 'TRICKY') return false;
    
    const cache = this.loadFromCache();
    const bundleIds = Object.keys(cache);
    
    if (bundleIds.length === 0) return true; // Empty cache
    
    // Check if any entry is older than 24 hours
    const oneDayAgo = new Date(Date.now() - 24 * 60 * 60 * 1000);
    
    for (const bundleId of bundleIds) {
      const lastUpdated = cache[bundleId].lastUpdated;
      if (!lastUpdated || new Date(lastUpdated) < oneDayAgo) {
        return true;
      }
    }
    
    return false;
  }

  /**
   * Ensure cache is up to date
   */
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
}

/**
 * Extract bundle ID from campaign name - ИСПРАВЛЕНО: правильная логика поиска
 * Ищем знак "=", потом берем текст от следующего символа до следующего пробела
 */
function extractBundleIdFromCampaign(campaignName) {
  console.log(`extractBundleIdFromCampaign called with: "${campaignName}"`);
  
  if (!campaignName) {
    console.log('Campaign name is empty, returning null');
    return null;
  }
  
  // Ищем знак "="
  const equalsIndex = campaignName.indexOf('=');
  if (equalsIndex === -1) {
    console.log('No "=" found, using fallback logic');
    // Fallback: берем текст до первого пробела
    const spaceIndex = campaignName.indexOf(' ');
    if (spaceIndex === -1) {
      console.log('No space found, returning whole campaign name:', campaignName.trim());
      return campaignName.trim();
    }
    
    const bundleId = campaignName.substring(0, spaceIndex).trim();
    console.log(`Extracted before space: "${bundleId}"`);
    
    // Валидация: bundle ID должен содержать точку или быть iOS ID (только цифры 8+ символов)
    if (bundleId.includes('.') || (/^\d{8,}$/.test(bundleId))) {
      console.log(`Bundle ID validated: "${bundleId}"`);
      return bundleId;
    } else {
      console.log(`Bundle ID rejected (no dot or not iOS ID): "${bundleId}"`);
      return null;
    }
  }
  
  // Найден "=", ищем следующий символ (исключая пробелы)
  let startIndex = equalsIndex + 1;
  
  // Пропускаем пробелы после "="
  while (startIndex < campaignName.length && campaignName[startIndex] === ' ') {
    startIndex++;
  }
  
  if (startIndex >= campaignName.length) {
    console.log('No content after "=", returning null');
    return null;
  }
  
  console.log(`Found "=" at position ${equalsIndex}, text after: "${campaignName.substring(startIndex)}"`);
  
  // Ищем следующий пробел после начального символа
  const spaceIndex = campaignName.indexOf(' ', startIndex);
  let bundleId;
  
  if (spaceIndex === -1) {
    // Нет пробела - берем до конца строки
    bundleId = campaignName.substring(startIndex).trim();
  } else {
    // Есть пробел - берем до него
    bundleId = campaignName.substring(startIndex, spaceIndex).trim();
  }
  
  console.log(`Extracted bundle ID from subject: "${bundleId}"`);
  
  // Валидация: bundle ID должен содержать точку (Android) или быть только цифрами 8+ символов (iOS)
  if (bundleId.includes('.') || (/^\d{8,}$/.test(bundleId))) {
    console.log(`Bundle ID validated: "${bundleId}"`);
    return bundleId;
  } else {
    console.log(`Bundle ID rejected (no dot or not iOS ID): "${bundleId}"`);
    return null;
  }
}

/**
 * Global function to refresh Apps Database cache manually
 */
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

/**
 * Debug function to update Apps Database with detailed logging
 */
function debugAppsDatabase() {
  console.log('=== APPS DATABASE DEBUG START ===');
  
  if (CURRENT_PROJECT !== 'TRICKY') {
    console.log('ОШИБКА: Apps Database только для TRICKY проекта');
    return;
  }
  
  try {
    const appsDb = new AppsDatabase('TRICKY');
    console.log('✅ AppsDatabase объект создан');
    
    // Check cache sheet
    if (!appsDb.cacheSheet) {
      console.log('❌ Cache sheet не найден');
      return;
    }
    console.log('✅ Cache sheet найден:', appsDb.config.APPS_CACHE_SHEET);
    
    // Access external spreadsheet
    console.log('📋 Подключаемся к внешней таблице:', APPS_DATABASE_ID);
    const externalSpreadsheet = SpreadsheetApp.openById(APPS_DATABASE_ID);
    console.log('✅ Внешняя таблица подключена');
    
    const externalSheet = externalSpreadsheet.getSheetByName(APPS_DATABASE_SHEET);
    if (!externalSheet) {
      console.log('❌ Лист не найден:', APPS_DATABASE_SHEET);
      return;
    }
    console.log('✅ Лист найден:', APPS_DATABASE_SHEET);
    
    // Get data
    const externalData = externalSheet.getDataRange().getValues();
    console.log('📊 Данных получено:', externalData.length, 'строк');
    
    if (externalData.length < 2) {
      console.log('❌ Недостаточно данных в таблице');
      return;
    }
    
    // Analyze headers
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
    
    // Process sample data
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
        const bundleId = appsDb.extractBundleIdFromLink(linkApp);
        console.log('  Extracted Bundle ID:', bundleId || 'НЕ УДАЛОСЬ ИЗВЛЕЧЬ');
        
        if (bundleId) {
          // Simulate processing logic
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
    
    // Full processing
    console.log('\n🚀 Начинаем полную обработку...');
    const success = appsDb.updateCacheFromExternalTable();
    
    if (success) {
      const cache = appsDb.loadFromCache();
      const count = Object.keys(cache).length;
      console.log('✅ УСПЕХ! Кеш обновлен:', count, 'приложений');
      
      // Show sample cached data with analysis
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
      
      // Overall statistics
      console.log('\n📈 СТАТИСТИКА КАЧЕСТВА ДАННЫХ:');
      console.log(`  - Всего в кеше: ${count} приложений`);
      console.log(`  - С пустыми Publisher: ~${Math.round(emptyPublisherCount * count / cacheKeys.length)}`);
      console.log(`  - С пустыми App Name: ~${Math.round(emptyAppNameCount * count / cacheKeys.length)}`);
      console.log(`  - Поддержка объединенных ячеек: ✅ Включена`);
    } else {
      console.log('❌ ОШИБКА: Обновление кеша провалилось');
    }
    
  } catch (e) {
    console.log('💥 ИСКЛЮЧЕНИЕ:', e.toString());
    console.log('📋 Stack trace:', e.stack);
  }
  
    console.log('\n=== APPS DATABASE DEBUG END ===');
}

/**
 * Test function to debug bundle ID extraction and Apps Database lookup
 */
function testAppsDbBundleExtraction() {
  console.log('=== TESTING BUNDLE ID EXTRACTION AND APPS DB LOOKUP ===');
  
  // Set project to TRICKY
  setCurrentProject('TRICKY');
  console.log('Project set to:', CURRENT_PROJECT);
  
  // Test campaign names from the table
  const testCampaigns = [
    '[pb tricky] | NPCW | USA | bm | I | subject = words.puzzle.wordgame.free.connect Bidmachine AMBO CPI skipctr',
    '[pb tricky] | NPCW | GBR | bm | I | subject = com.fanatee.cody Bidmachine skipctr AMBO CPI',
    '[pb tricky] | NPCNM | CAN | bm | I | subject = 6473832648 AMBO CPI abdoul Bidmachine skipctr autobudget',
    '[pb tricky] | NPKS | USA | subj=com.intensedev.classicsolitaire.gp Bidmachine AMBO CPA skipctr'
  ];
  
  console.log('\n🧪 TESTING CAMPAIGN NAME PROCESSING:');
  
  testCampaigns.forEach((campaignName, index) => {
    console.log(`\n--- Test ${index + 1}: "${campaignName}" ---`);
    
    // Step 1: Extract bundle ID
    const bundleId = extractBundleIdFromCampaign(campaignName);
    console.log(`Extracted bundle ID: "${bundleId}"`);
    
    if (bundleId) {
      // Step 2: Get Apps Database instance
      const appsDb = new AppsDatabase('TRICKY');
      console.log('Apps Database instance created');
      
      // Step 3: Get display name
      const displayName = appsDb.getSourceAppDisplayName(bundleId);
      console.log(`Final display name: "${displayName}"`);
      
      // Step 4: Check cache directly
      const cache = appsDb.loadFromCache();
      const cacheEntry = cache[bundleId];
      console.log(`Cache entry for "${bundleId}":`, cacheEntry);
      
    } else {
      console.log('❌ Bundle ID extraction failed');
    }
  });
  
  // Additional cache statistics
  try {
    const appsDb = new AppsDatabase('TRICKY');
    const cache = appsDb.loadFromCache();
    const cacheSize = Object.keys(cache).length;
    console.log(`\n📊 Apps Database cache size: ${cacheSize} entries`);
    
    if (cacheSize > 0) {
      const sampleKeys = Object.keys(cache).slice(0, 5);
      console.log('📋 Sample cache entries:');
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