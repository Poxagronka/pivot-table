/**
 * Apps Database Cache Management - оптимизированная версия
 * Кеширует данные из внешней таблицы Apps Database для группировки по Publisher + App Name
 */

class AppsDatabase {
  constructor(projectName = null) {
    this.projectName = projectName ? projectName.toUpperCase() : CURRENT_PROJECT;
    this.config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
    this.cacheSheet = (this.projectName === 'TRICKY') ? this.getOrCreateCacheSheet() : null;
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

  updateCacheFromExternalTable() {
    if (!this.cacheSheet) return false;
    try {
      const externalSpreadsheet = SpreadsheetApp.openById(APPS_DATABASE_ID);
      const externalSheet = externalSpreadsheet.getSheetByName(APPS_DATABASE_SHEET);
      if (!externalSheet) return false;
      
      const externalData = externalSheet.getDataRange().getValues();
      if (externalData.length < 2) return false;
      
      const headers = externalData[0];
      const linkAppCol = this.findColumnIndex(headers, ['Link App', 'link_app', 'linkapp', 'links']);
      const publisherCol = this.findColumnIndex(headers, ['Publisher', 'publisher']);
      const appNameCol = this.findColumnIndex(headers, ['App Name', 'app_name', 'appname']);
      if (linkAppCol === -1) return false;
      
      if (this.cacheSheet.getLastRow() > 1) {
        this.cacheSheet.deleteRows(2, this.cacheSheet.getLastRow() - 1);
      }
      
      const cacheData = [];
      const now = new Date();
      let lastPublisher = '', lastAppName = '';
      
      for (let i = 1; i < externalData.length; i++) {
        const row = externalData[i];
        const linkApp = row[linkAppCol];
        
        // Обработка значений publisher и appName
        const [publisher, appName] = this.processRowData(row, publisherCol, appNameCol, lastPublisher, lastAppName);
        if (publisher) lastPublisher = publisher;
        if (appName) lastAppName = appName;
        
        if (linkApp && linkApp.trim && linkApp.trim()) {
          const bundleId = this.extractBundleIdFromLink(linkApp);
          if (bundleId) {
            const finalPublisher = publisher || lastPublisher || 'Unknown Publisher';
            const finalAppName = appName || lastAppName || finalPublisher || 'Unknown App';
            cacheData.push([bundleId, finalPublisher, finalAppName, linkApp, now]);
          }
        }
      }
      
      if (cacheData.length > 0) {
        this.cacheSheet.getRange(this.cacheSheet.getLastRow() + 1, 1, cacheData.length, 5).setValues(cacheData);
      }
      return true;
    } catch (e) {
      return false;
    }
  }
  
  processRowData(row, publisherCol, appNameCol, lastPublisher, lastAppName) {
    const getCleanValue = (col) => {
      const raw = col !== -1 ? row[col] : '';
      return (raw && raw.toString) ? raw.toString().trim() : '';
    };
    return [getCleanValue(publisherCol), getCleanValue(appNameCol)];
  }

  extractBundleIdFromLink(linkApp) {
    if (!linkApp || typeof linkApp !== 'string') return null;
    const url = linkApp.trim();
    if (url === 'no app found' || url.length < 10 || url.includes('idx') || url.endsWith('id...')) return null;
    
    try {
      // iOS patterns - объединенные в один
      const iosMatch = url.match(/\/id(\d{8,})(?:[^\d]|$)/);
      if (iosMatch && iosMatch[1] && /^\d+$/.test(iosMatch[1])) {
        return iosMatch[1];
      }
      
      // Android patterns - объединенные в один
      const androidMatch = url.match(/[?&]id=([a-zA-Z][a-zA-Z0-9._]*[a-zA-Z0-9])(?:&|$)/);
      if (androidMatch && androidMatch[1] && androidMatch[1].includes('.') && androidMatch[1].length >= 3) {
        return androidMatch[1];
      }
    } catch (e) {}
    return null;
  }

  findColumnIndex(headers, possibleNames) {
    return headers.findIndex(h => 
      possibleNames.some(name => 
        (h || '').toString().toLowerCase().trim() === name.toLowerCase()
      )
    );
  }

  getAppInfo(bundleId) {
    if (!bundleId || this.projectName !== 'TRICKY') {
      return { publisher: 'Unknown Publisher', appName: 'Unknown App', linkApp: '' };
    }
    const appInfo = this.loadFromCache()[bundleId];
    return appInfo ? {
      publisher: appInfo.publisher,
      appName: appInfo.appName,
      linkApp: appInfo.linkApp
    } : { publisher: bundleId, appName: '', linkApp: '' };
  }

  getSourceAppDisplayName(bundleId) {
    if (!bundleId || this.projectName !== 'TRICKY') return bundleId || 'Unknown';
    const { publisher, appName } = this.getAppInfo(bundleId);
    if (publisher !== bundleId) {
      if (publisher && appName && publisher !== appName) return `${publisher} ${appName}`;
      if (publisher) return publisher;
      if (appName) return appName;
    }
    return bundleId;
  }

  shouldUpdateCache() {
    if (!this.cacheSheet || this.projectName !== 'TRICKY') return false;
    const cache = this.loadFromCache();
    const bundleIds = Object.keys(cache);
    if (bundleIds.length === 0) return true;
    const oneDayAgo = new Date(Date.now() - 24 * 60 * 60 * 1000);
    return bundleIds.some(id => {
      const lastUpdated = cache[id].lastUpdated;
      return !lastUpdated || new Date(lastUpdated) < oneDayAgo;
    });
  }

  ensureCacheUpToDate() {
    if (this.projectName === 'TRICKY' && this.shouldUpdateCache()) {
      try { this.updateCacheFromExternalTable(); } catch (e) {}
    }
  }
}

function extractBundleIdFromCampaign(campaignName) {
  if (!campaignName) return null;
  
  const equalsIndex = campaignName.indexOf('=');
  if (equalsIndex !== -1) {
    // Проверяем bundle ID до "="
    const beforeEquals = campaignName.substring(0, equalsIndex).trim();
    if (isValidBundleId(beforeEquals)) return beforeEquals;
    
    // Проверяем bundle ID после "="
    const afterEquals = campaignName.substring(equalsIndex + 1).trim();
    const spaceIndex = afterEquals.indexOf(' ');
    const bundleId = spaceIndex === -1 ? afterEquals : afterEquals.substring(0, spaceIndex).trim();
    if (isValidBundleId(bundleId)) return bundleId;
    return null;
  }
  
  // Проверяем bundle ID в начале строки до первого пробела
  const spaceIndex = campaignName.indexOf(' ');
  const bundleId = spaceIndex === -1 ? campaignName.trim() : campaignName.substring(0, spaceIndex).trim();
  return isValidBundleId(bundleId) ? bundleId : null;
}

function isValidBundleId(bundleId) {
  if (!bundleId || typeof bundleId !== 'string') return false;
  const trimmed = bundleId.trim();
  if (trimmed.length < 3) return false;
  // Android: содержит точку и начинается с буквы
  if (trimmed.includes('.') && /^[a-zA-Z][a-zA-Z0-9._]*[a-zA-Z0-9]$/.test(trimmed)) return true;
  // iOS: только цифры, 8+ символов  
  return /^\d{8,}$/.test(trimmed);
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
    const message = success 
      ? `Successfully updated Apps Database cache.\n\n${Object.keys(appsDb.loadFromCache()).length} apps loaded.`
      : 'Failed to update Apps Database cache.\n\nCheck console for errors.';
    ui.alert(success ? 'Apps Database Updated' : 'Update Failed', message, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Error updating Apps Database: ' + e.toString(), ui.ButtonSet.OK);
  }
}

