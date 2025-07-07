/**
 * Apps Database Cache Management
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
      const bundleIdCol = this.findColumnIndex(headers, ['Bundle ID', 'bundle_id', 'bundleid']);
      const publisherCol = this.findColumnIndex(headers, ['Publisher', 'publisher']);
      const appNameCol = this.findColumnIndex(headers, ['App Name', 'app_name', 'appname']);
      const linkAppCol = this.findColumnIndex(headers, ['Link App', 'link_app', 'linkapp']);
      
      if (bundleIdCol === -1) {
        console.error('Bundle ID column not found in Apps Database');
        return false;
      }
      
      // Clear cache sheet except headers
      if (this.cacheSheet.getLastRow() > 1) {
        this.cacheSheet.deleteRows(2, this.cacheSheet.getLastRow() - 1);
      }
      
      // Process external data
      const cacheData = [];
      const now = new Date();
      
      for (let i = 1; i < externalData.length; i++) {
        const row = externalData[i];
        const bundleId = row[bundleIdCol];
        
        if (bundleId && bundleId.trim()) {
          const publisher = publisherCol !== -1 ? (row[publisherCol] || 'Unknown Publisher') : 'Unknown Publisher';
          const appName = appNameCol !== -1 ? (row[appNameCol] || 'Unknown App') : 'Unknown App';
          const linkApp = linkAppCol !== -1 ? (row[linkAppCol] || '') : '';
          
          cacheData.push([bundleId.trim(), publisher, appName, linkApp, now]);
        }
      }
      
      // Write to cache
      if (cacheData.length > 0) {
        const lastRow = this.cacheSheet.getLastRow();
        this.cacheSheet.getRange(lastRow + 1, 1, cacheData.length, 5).setValues(cacheData);
      }
      
      console.log(`Apps Database cache updated: ${cacheData.length} apps`);
      return true;
      
    } catch (e) {
      console.error('Error updating Apps Database cache:', e);
      return false;
    }
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
    if (!bundleId || this.projectName !== 'TRICKY') {
      return bundleId || 'Unknown';
    }
    
    const appInfo = this.getAppInfo(bundleId);
    
    if (appInfo.publisher !== bundleId && appInfo.appName) {
      // Found in database: "Publisher App Name"
      return `${appInfo.publisher} ${appInfo.appName}`;
    } else {
      // Not found: use bundle ID
      return bundleId;
    }
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
 * Extract bundle ID from campaign name (text before first space)
 */
function extractBundleIdFromCampaign(campaignName) {
  if (!campaignName) return null;
  
  const spaceIndex = campaignName.indexOf(' ');
  if (spaceIndex === -1) return campaignName.trim();
  
  const bundleId = campaignName.substring(0, spaceIndex).trim();
  
  // Basic validation: bundle ID should contain a dot
  if (bundleId.includes('.')) {
    return bundleId;
  }
  
  return null;
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