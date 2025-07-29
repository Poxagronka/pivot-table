/**
 * Initial Metrics Cache Management - Кеширует первоначальные значения eROAS 730d и eProfit 730d в отдельной таблице
 */

const INITIAL_METRICS_CACHE_SPREADSHEET_ID = '1JBYtINHH7yLwdsfCPV3q3sj6NlP3WmftsPvbdfzTWdU';

class InitialMetricsCache {
  constructor(projectName = null) {
    this.projectName = projectName ? projectName.toUpperCase() : CURRENT_PROJECT;
    this.cacheSpreadsheet = null;
    this.cacheSheet = null;
    this.memoryCache = null;
    
    try {
      this.cacheSpreadsheet = SpreadsheetApp.openById(INITIAL_METRICS_CACHE_SPREADSHEET_ID);
    } catch (e) {
      console.error('Failed to open Initial Metrics cache spreadsheet:', e);
      throw new Error('Cannot access Initial Metrics cache spreadsheet. Check ID and permissions.');
    }
  }

  getOrCreateCacheSheet() {
    if (this.cacheSheet) return this.cacheSheet;
    
    const sheetName = `InitialMetrics_${this.projectName}`;
    this.cacheSheet = this.cacheSpreadsheet.getSheetByName(sheetName);
    
    if (!this.cacheSheet) {
      this.cacheSheet = this.cacheSpreadsheet.insertSheet(sheetName);
      this.cacheSheet.getRange(1, 1, 1, 8).setValues([['Level', 'AppName', 'WeekRange', 'Identifier', 'SourceApp', 'InitialEROAS', 'DateRecorded', 'InitialProfit']]);
      this.cacheSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f0f0f0');
    } else {
      this.migrateSheetStructure();
    }
    
    return this.cacheSheet;
  }

  migrateSheetStructure() {
    const headers = this.cacheSheet.getRange(1, 1, 1, this.cacheSheet.getLastColumn()).getValues()[0];
    
    if (headers.length < 8 || headers[7] !== 'InitialProfit') {
      console.log(`Migrating ${this.projectName} sheet structure to include InitialProfit`);
      
      if (headers.length < 8) {
        this.cacheSheet.getRange(1, 8).setValue('InitialProfit');
        this.cacheSheet.getRange(1, 8).setFontWeight('bold').setBackground('#f0f0f0');
      }
    }
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
      const lastCol = Math.max(8, sheet.getLastColumn());
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
      
      data.forEach(row => {
        const [level, appName, weekRange, identifier, sourceApp, initialEROAS, dateRecorded, initialProfit] = row;
        if ((initialEROAS !== '' && initialEROAS > 0) || (initialProfit !== '' && initialProfit > 0)) {
          const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
          cache[key] = {
            eROAS: initialEROAS && initialEROAS > 0 ? parseFloat(initialEROAS) : null,
            profit: initialProfit && initialProfit > 0 ? parseFloat(initialProfit) : null
          };
        }
      });
      
      this.memoryCache = cache;
      console.log(`Loaded ${Object.keys(cache).length} initial metrics values for ${this.projectName}`);
    } catch (e) {
      console.error('Error loading initial metrics values:', e);
      this.memoryCache = {};
    }
    
    return this.memoryCache;
  }

  saveInitialValue(level, appName, weekRange, currentEROAS, currentProfit, identifier = '', sourceApp = '') {
    if ((currentEROAS === null || currentEROAS === undefined || currentEROAS === '' || currentEROAS === 0) &&
        (currentProfit === null || currentProfit === undefined || currentProfit === '' || currentProfit === 0)) return;
    
    const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
    const cache = this.loadAllInitialValues();
    
    if (cache[key] !== undefined) {
      return;
    }
    
    try {
      const sheet = this.getOrCreateCacheSheet();
      const lastRow = sheet.getLastRow();
      
      const validEROAS = currentEROAS && currentEROAS > 0 ? currentEROAS : '';
      const validProfit = currentProfit && currentProfit > 0 ? currentProfit : '';
      
      sheet.getRange(lastRow + 1, 1, 1, 8).setValues([[
        level,
        appName,
        weekRange,
        identifier || '',
        sourceApp || '',
        validEROAS,
        new Date(),
        validProfit
      ]]);
      
      if (this.memoryCache) {
        this.memoryCache[key] = {
          eROAS: validEROAS ? parseFloat(validEROAS) : null,
          profit: validProfit ? parseFloat(validProfit) : null
        };
      }
    } catch (e) {
      console.error('Error saving initial metrics value:', e);
    }
  }

  recordInitialValuesFromData(appData) {
    console.log(`${this.projectName}: Recording initial metrics values for all weeks in data`);
    
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
          if (!this.memoryCache[weekKey] && (weekTotals.avgEROASD730 > 0 || weekTotals.totalProfit > 0)) {
            newValues.push(['WEEK', network.networkName, weekRange, '', '', 
              weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : '', 
              new Date(), 
              weekTotals.totalProfit > 0 ? weekTotals.totalProfit : ''
            ]);
          }
          
          Object.values(week.apps).forEach(app => {
            const appTotals = calculateWeekTotals(app.campaigns);
            const appKey = this.createKey('APP', network.networkName, weekRange, app.appId, app.appName);
            if (!this.memoryCache[appKey] && (appTotals.avgEROASD730 > 0 || appTotals.totalProfit > 0)) {
              newValues.push(['APP', network.networkName, weekRange, app.appId, app.appName, 
                appTotals.avgEROASD730 > 0 ? appTotals.avgEROASD730 : '', 
                new Date(), 
                appTotals.totalProfit > 0 ? appTotals.totalProfit : ''
              ]);
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
            if (!this.memoryCache[weekKey] && (weekTotals.avgEROASD730 > 0 || weekTotals.totalProfit > 0)) {
              newValues.push(['WEEK', app.appName, weekRange, '', '', 
                weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : '', 
                new Date(), 
                weekTotals.totalProfit > 0 ? weekTotals.totalProfit : ''
              ]);
            }
            
            Object.values(week.sourceApps).forEach(sourceApp => {
              const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
              const sourceAppKey = this.createKey('SOURCE_APP', app.appName, weekRange, sourceApp.sourceAppId, sourceApp.sourceAppName);
              if (!this.memoryCache[sourceAppKey] && (sourceAppTotals.avgEROASD730 > 0 || sourceAppTotals.totalProfit > 0)) {
                newValues.push(['SOURCE_APP', app.appName, weekRange, sourceApp.sourceAppId, sourceApp.sourceAppName, 
                  sourceAppTotals.avgEROASD730 > 0 ? sourceAppTotals.avgEROASD730 : '', 
                  new Date(), 
                  sourceAppTotals.totalProfit > 0 ? sourceAppTotals.totalProfit : ''
                ]);
              }
              
              sourceApp.campaigns.forEach(campaign => {
                const campaignKey = this.createKey('CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp);
                if (!this.memoryCache[campaignKey] && (campaign.eRoasForecastD730 > 0 || campaign.eProfitForecast > 0)) {
                  newValues.push(['CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp, 
                    campaign.eRoasForecastD730 > 0 ? campaign.eRoasForecastD730 : '', 
                    new Date(), 
                    campaign.eProfitForecast > 0 ? campaign.eProfitForecast : ''
                  ]);
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
            if (!this.memoryCache[weekKey] && (weekTotals.avgEROASD730 > 0 || weekTotals.totalProfit > 0)) {
              newValues.push(['WEEK', app.appName, weekRange, '', '', 
                weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : '', 
                new Date(), 
                weekTotals.totalProfit > 0 ? weekTotals.totalProfit : ''
              ]);
            }
            
            Object.values(week.networks).forEach(network => {
              const networkTotals = calculateWeekTotals(network.campaigns);
              const networkKey = this.createKey('NETWORK', app.appName, weekRange, network.networkId, network.networkName);
              if (!this.memoryCache[networkKey] && (networkTotals.avgEROASD730 > 0 || networkTotals.totalProfit > 0)) {
                newValues.push(['NETWORK', app.appName, weekRange, network.networkId, network.networkName, 
                  networkTotals.avgEROASD730 > 0 ? networkTotals.avgEROASD730 : '', 
                  new Date(), 
                  networkTotals.totalProfit > 0 ? networkTotals.totalProfit : ''
                ]);
              }
            });
          } else {
            const weekTotals = calculateWeekTotals(week.campaigns);
            
            const weekKey = this.createKey('WEEK', app.appName, weekRange, '', '');
            if (!this.memoryCache[weekKey] && (weekTotals.avgEROASD730 > 0 || weekTotals.totalProfit > 0)) {
              newValues.push(['WEEK', app.appName, weekRange, '', '', 
                weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : '', 
                new Date(), 
                weekTotals.totalProfit > 0 ? weekTotals.totalProfit : ''
              ]);
            }
            
            if (week.campaigns) {
              week.campaigns.forEach(campaign => {
                const campaignKey = this.createKey('CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp);
                if (!this.memoryCache[campaignKey] && (campaign.eRoasForecastD730 > 0 || campaign.eProfitForecast > 0)) {
                  newValues.push(['CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp, 
                    campaign.eRoasForecastD730 > 0 ? campaign.eRoasForecastD730 : '', 
                    new Date(), 
                    campaign.eProfitForecast > 0 ? campaign.eProfitForecast : ''
                  ]);
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
        sheet.getRange(lastRow + 1, 1, newValues.length, 8).setValues(newValues);
        
        newValues.forEach(row => {
          const [level, appName, weekRange, identifier, sourceApp, initialEROAS, dateRecorded, initialProfit] = row;
          const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
          if (this.memoryCache) {
            this.memoryCache[key] = {
              eROAS: initialEROAS && initialEROAS > 0 ? parseFloat(initialEROAS) : null,
              profit: initialProfit && initialProfit > 0 ? parseFloat(initialProfit) : null
            };
          }
        });
        
        console.log(`${this.projectName}: Recorded ${newValues.length} new initial metrics values`);
      } catch (e) {
        console.error(`Error batch saving initial metrics values for ${this.projectName}:`, e);
      }
    } else {
      console.log(`${this.projectName}: No new initial metrics values to record`);
    }
  }

  formatEROASWithInitial(level, appName, weekRange, currentEROAS, identifier = '', sourceApp = '') {
    const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
    const cache = this.loadAllInitialValues();
    
    const metrics = cache[key];
    const currentValue = Math.round(currentEROAS);
    
    if (metrics && metrics.eROAS !== null) {
      const initialRounded = Math.round(metrics.eROAS);
      return `${initialRounded}% → ${currentValue}%`;
    } else {
      return `${currentValue}% → ${currentValue}%`;
    }
  }

  formatProfitWithInitial(level, appName, weekRange, currentProfit, identifier = '', sourceApp = '') {
    const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
    const cache = this.loadAllInitialValues();
    
    const metrics = cache[key];
    const currentValue = Math.round(currentProfit);
    
    if (metrics && metrics.profit !== null) {
      const initialRounded = Math.round(metrics.profit);
      return `${initialRounded}$ → ${currentValue}$`;
    } else {
      return `${currentValue}$ → ${currentValue}$`;
    }
  }

  clearMemoryCache() {
    this.memoryCache = null;
  }
}

function clearAllInitialMetricsMemoryCaches() {
  console.log('Clearing all initial metrics memory caches...');
}