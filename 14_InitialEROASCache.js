/**
 * Initial eROAS Cache Management - Кеширует первоначальные значения eROAS 730d в отдельной таблице
 */

const INITIAL_EROAS_CACHE_SPREADSHEET_ID = '1JBYtINHH7yLwdsfCPV3q3sj6NlP3WmftsPvbdfzTWdU';

class InitialEROASCache {
  constructor(projectName = null) {
    this.projectName = projectName ? projectName.toUpperCase() : CURRENT_PROJECT;
    this.cacheSpreadsheet = null;
    this.cacheSheet = null;
    this.memoryCache = null;
    
    try {
      this.cacheSpreadsheet = SpreadsheetApp.openById(INITIAL_EROAS_CACHE_SPREADSHEET_ID);
    } catch (e) {
      console.error('Failed to open Initial eROAS cache spreadsheet:', e);
      throw new Error('Cannot access Initial eROAS cache spreadsheet. Check ID and permissions.');
    }
  }

  getOrCreateCacheSheet() {
    if (this.cacheSheet) return this.cacheSheet;
    
    const sheetName = `InitialEROAS_${this.projectName}`;
    this.cacheSheet = this.cacheSpreadsheet.getSheetByName(sheetName);
    
    if (!this.cacheSheet) {
      this.cacheSheet = this.cacheSpreadsheet.insertSheet(sheetName);
      this.cacheSheet.getRange(1, 1, 1, 7).setValues([['Level', 'AppName', 'WeekRange', 'Identifier', 'SourceApp', 'InitialEROAS', 'DateRecorded']]);
      this.cacheSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#f0f0f0');
    }
    
    return this.cacheSheet;
  }

  createKey(level, appName, weekRange, identifier = '', sourceApp = '') {
    return `${level}|||${appName}|||${weekRange}|||${identifier}|||${sourceApp}`;
  }

  loadAllInitialValues() {
    if (this.memoryCache) {
      return this.memoryCache;
    }
    
    const sheet = this.getOrCreateCacheSheet();
    const cache = {};
    
    if (sheet.getLastRow() <= 1) {
      this.memoryCache = cache;
      return cache;
    }
    
    try {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
      
      data.forEach(row => {
        const [level, appName, weekRange, identifier, sourceApp, initialEROAS, dateRecorded] = row;
        if (initialEROAS !== '' && initialEROAS > 0) {
          const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
          cache[key] = parseFloat(initialEROAS);
        }
      });
      
      this.memoryCache = cache;
      console.log(`Loaded ${Object.keys(cache).length} initial eROAS values for ${this.projectName}`);
    } catch (e) {
      console.error('Error loading initial eROAS values:', e);
      this.memoryCache = {};
    }
    
    return this.memoryCache;
  }

  saveInitialValue(level, appName, weekRange, currentEROAS, identifier = '', sourceApp = '') {
    if (currentEROAS === null || currentEROAS === undefined || currentEROAS === '' || currentEROAS === 0) return;
    
    const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
    const cache = this.loadAllInitialValues();
    
    if (cache[key] !== undefined) {
      return;
    }
    
    try {
      const sheet = this.getOrCreateCacheSheet();
      const lastRow = sheet.getLastRow();
      
      sheet.getRange(lastRow + 1, 1, 1, 7).setValues([[
        level,
        appName,
        weekRange,
        identifier || '',
        sourceApp || '',
        currentEROAS,
        new Date()
      ]]);
      
      if (this.memoryCache) {
        this.memoryCache[key] = currentEROAS;
      }
    } catch (e) {
      console.error('Error saving initial eROAS value:', e);
    }
  }

  recordInitialValuesFromData(appData) {
    console.log(`${this.projectName}: Recording initial eROAS values for all weeks in data`);
    
    this.loadAllInitialValues();
    
    const newValues = [];
    
    if (this.projectName === 'INCENT_TRAFFIC') {
      Object.values(appData).forEach(network => {
        Object.values(network.weeks).forEach(week => {
          const weekRange = `${week.weekStart} - ${week.weekEnd}`;
          
          const allCampaigns = [];
          Object.values(week.apps).forEach(app => {
            allCampaigns.push(...app.campaigns);
          });
          const weekTotals = calculateWeekTotals(allCampaigns);
          
          const weekKey = this.createKey('WEEK', network.networkName, weekRange, '', '');
          if (!this.memoryCache[weekKey] && weekTotals.avgEROASD730 > 0) {
            newValues.push(['WEEK', network.networkName, weekRange, '', '', weekTotals.avgEROASD730, new Date()]);
          }
          
          Object.values(week.apps).forEach(app => {
            const appTotals = calculateWeekTotals(app.campaigns);
            const appKey = this.createKey('APP', network.networkName, weekRange, app.appId, app.appName);
            if (!this.memoryCache[appKey] && appTotals.avgEROASD730 > 0) {
              newValues.push(['APP', network.networkName, weekRange, app.appId, app.appName, appTotals.avgEROASD730, new Date()]);
            }
          });
        });
      });
    } else {
      Object.values(appData).forEach(app => {
        Object.values(app.weeks).forEach(week => {
          const weekRange = `${week.weekStart} - ${week.weekEnd}`;
          
          if (this.projectName === 'TRICKY' && week.sourceApps) {
            const allCampaigns = [];
            Object.values(week.sourceApps).forEach(sourceApp => {
              allCampaigns.push(...sourceApp.campaigns);
            });
            const weekTotals = calculateWeekTotals(allCampaigns);
            
            const weekKey = this.createKey('WEEK', app.appName, weekRange, '', '');
            if (!this.memoryCache[weekKey] && weekTotals.avgEROASD730 > 0) {
              newValues.push(['WEEK', app.appName, weekRange, '', '', weekTotals.avgEROASD730, new Date()]);
            }
            
            Object.values(week.sourceApps).forEach(sourceApp => {
              const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
              const sourceAppKey = this.createKey('SOURCE_APP', app.appName, weekRange, sourceApp.sourceAppId, sourceApp.sourceAppName);
              if (!this.memoryCache[sourceAppKey] && sourceAppTotals.avgEROASD730 > 0) {
                newValues.push(['SOURCE_APP', app.appName, weekRange, sourceApp.sourceAppId, sourceApp.sourceAppName, sourceAppTotals.avgEROASD730, new Date()]);
              }
              
              sourceApp.campaigns.forEach(campaign => {
                const campaignKey = this.createKey('CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp);
                if (!this.memoryCache[campaignKey] && campaign.eRoasForecastD730 > 0) {
                  newValues.push(['CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp, campaign.eRoasForecastD730, new Date()]);
                }
              });
            });
          } else if (this.projectName === 'OVERALL' && week.networks) {
            const allCampaigns = [];
            Object.values(week.networks).forEach(network => {
              allCampaigns.push(...network.campaigns);
            });
            const weekTotals = calculateWeekTotals(allCampaigns);
            
            const weekKey = this.createKey('WEEK', app.appName, weekRange, '', '');
            if (!this.memoryCache[weekKey] && weekTotals.avgEROASD730 > 0) {
              newValues.push(['WEEK', app.appName, weekRange, '', '', weekTotals.avgEROASD730, new Date()]);
            }
            
            Object.values(week.networks).forEach(network => {
              const networkTotals = calculateWeekTotals(network.campaigns);
              const networkKey = this.createKey('NETWORK', app.appName, weekRange, network.networkId, network.networkName);
              if (!this.memoryCache[networkKey] && networkTotals.avgEROASD730 > 0) {
                newValues.push(['NETWORK', app.appName, weekRange, network.networkId, network.networkName, networkTotals.avgEROASD730, new Date()]);
              }
            });
          } else {
            const weekTotals = calculateWeekTotals(week.campaigns);
            
            const weekKey = this.createKey('WEEK', app.appName, weekRange, '', '');
            if (!this.memoryCache[weekKey] && weekTotals.avgEROASD730 > 0) {
              newValues.push(['WEEK', app.appName, weekRange, '', '', weekTotals.avgEROASD730, new Date()]);
            }
            
            if (week.campaigns) {
              week.campaigns.forEach(campaign => {
                const campaignKey = this.createKey('CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp);
                if (!this.memoryCache[campaignKey] && campaign.eRoasForecastD730 > 0) {
                  newValues.push(['CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp, campaign.eRoasForecastD730, new Date()]);
                }
              });
            }
          }
        });
      });
    }
    
    if (newValues.length > 0) {
      try {
        const sheet = this.getOrCreateCacheSheet();
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1, newValues.length, 7).setValues(newValues);
        
        newValues.forEach(row => {
          const [level, appName, weekRange, identifier, sourceApp, initialEROAS] = row;
          const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
          if (this.memoryCache) {
            this.memoryCache[key] = initialEROAS;
          }
        });
        
        console.log(`${this.projectName}: Recorded ${newValues.length} new initial eROAS values`);
      } catch (e) {
        console.error(`Error batch saving initial eROAS values for ${this.projectName}:`, e);
      }
    } else {
      console.log(`${this.projectName}: No new initial eROAS values to record`);
    }
  }

  formatEROASWithInitial(level, appName, weekRange, currentEROAS, identifier = '', sourceApp = '') {
    const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
    const cache = this.loadAllInitialValues();
    
    const initialValue = cache[key];
    const currentValue = Math.round(currentEROAS);
    
    if (initialValue !== undefined) {
      const initialRounded = Math.round(initialValue);
      return `${initialRounded}% → ${currentValue}%`;
    } else {
      return `${currentValue}% → ${currentValue}%`;
    }
  }

  clearMemoryCache() {
    this.memoryCache = null;
  }
}

function clearAllInitialEROASMemoryCaches() {
  console.log('Clearing all initial eROAS memory caches...');
}