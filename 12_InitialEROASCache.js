/**
 * Initial Metrics Cache Management - ОПТИМИЗИРОВАННАЯ ВЕРСИЯ
 * Сохранена полная обратная совместимость со всеми файлами
 */

const INITIAL_METRICS_CACHE_SPREADSHEET_ID = '1JBYtINHH7yLwdsfCPV3q3sj6NlP3WmftsPvbdfzTWdU';

class InitialMetricsCache {
  constructor(projectName = null) {
    this.projectName = projectName ? projectName.toUpperCase() : CURRENT_PROJECT;
    this.cacheSpreadsheet = null;
    this.cacheSheet = null;
    this.memoryCache = null;
    this.rowIndexCache = null;
    this.dateCache = null;
    
    try {
      this.cacheSpreadsheet = SpreadsheetApp.openById(INITIAL_METRICS_CACHE_SPREADSHEET_ID);
    } catch (e) {
      console.error('Failed to open Initial Metrics cache spreadsheet:', e);
      throw new Error('Cannot access Initial Metrics cache spreadsheet. Check ID and permissions.');
    }
  }

  // ========== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ==========
  isValidEROAS(value) {
    return value && value !== '' && value > 0;
  }
  
  isValidProfit(value) {
    return value !== '' && value !== null && value !== undefined;
  }
  
  formatMetricWithInitial(level, appName, weekRange, currentValue, metricType, identifier = '', sourceApp = '') {
    const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
    const cache = this.loadAllInitialValues();
    const metrics = cache[key];
    const rounded = Math.round(currentValue);
    
    if (metrics && metrics[metricType] !== null && metrics[metricType] !== undefined) {
      const initial = Math.round(metrics[metricType]);
      const suffix = metricType === 'eROAS' ? '%' : '$';
      return `${initial}${suffix} → ${rounded}${suffix}`;
    }
    const suffix = metricType === 'eROAS' ? '%' : '$';
    return `${rounded}${suffix} → ${rounded}${suffix}`;
  }
  
  initDateCache() {
    const today = new Date();
    const dayOfWeek = today.getDay();
    this.dateCache = {
      today: today,
      isAllowedToRecord: dayOfWeek >= 2 || dayOfWeek === 0,
      currentWeekStart: formatDateForAPI(getMondayOfWeek(today)),
      lastWeekStart: formatDateForAPI(getMondayOfWeek(new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000)))
    };
    return this.dateCache;
  }
  
  canRecordInitialForWeek(weekStart) {
    if (!this.dateCache) this.initDateCache();
    
    if (weekStart >= this.dateCache.currentWeekStart) return false;
    if (weekStart === this.dateCache.lastWeekStart) return this.dateCache.isAllowedToRecord;
    return true;
  }
  
  shouldRecordMetric(eROAS, profit) {
    return this.isValidEROAS(eROAS) || (profit !== 0 && profit !== null && profit !== undefined);
  }
  
  needsUpdate(existingEntry, eROAS, profit) {
    return (this.isValidEROAS(eROAS) && !existingEntry.eROAS) || 
           (profit !== 0 && profit !== null && profit !== undefined && !existingEntry.profit);
  }

  // ========== ОСНОВНЫЕ МЕТОДЫ ==========
  getOrCreateCacheSheet() {
    if (this.cacheSheet) return this.cacheSheet;
    
    const sheetName = `InitialEROAS_${this.projectName}`;
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
    this.rowIndexCache = {};
    
    if (sheet.getLastRow() <= 1) {
      this.memoryCache = cache;
      return cache;
    }
    
    try {
      const lastCol = 8; // Оптимизация - всегда 8 колонок
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
      
      data.forEach((row, index) => {
        const [level, appName, weekRange, identifier, sourceApp, initialEROAS, dateRecorded, initialProfit] = row;
        if (this.isValidEROAS(initialEROAS) || this.isValidProfit(initialProfit)) {
          const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
          cache[key] = {
            eROAS: this.isValidEROAS(initialEROAS) ? parseFloat(initialEROAS) : null,
            profit: this.isValidProfit(initialProfit) ? parseFloat(initialProfit) : null
          };
          this.rowIndexCache[key] = index + 2;
        }
      });
      
      // Legacy ключи для старых названий
      if (typeof APP_NAME_LEGACY !== 'undefined') {
        Object.keys(APP_NAME_LEGACY).forEach(newName => {
          const oldName = APP_NAME_LEGACY[newName];
          data.forEach((row, index) => {
            const [level, appName, weekRange, identifier, sourceApp, initialEROAS, dateRecorded, initialProfit] = row;
            if (appName === oldName && (this.isValidEROAS(initialEROAS) || this.isValidProfit(initialProfit))) {
              const legacyKey = this.createKey(level, newName, weekRange, identifier, sourceApp);
              if (!cache[legacyKey]) {
                cache[legacyKey] = {
                  eROAS: this.isValidEROAS(initialEROAS) ? parseFloat(initialEROAS) : null,
                  profit: this.isValidProfit(initialProfit) ? parseFloat(initialProfit) : null
                };
              }
            }
          });
        });
      }
      
      this.memoryCache = cache;
      console.log(`Loaded ${Object.keys(cache).length} initial metrics values for ${this.projectName}`);
    } catch (e) {
      console.error('Error loading initial metrics values:', e);
      this.memoryCache = {};
      this.rowIndexCache = {};
    }
    
    return this.memoryCache;
  }

  saveInitialValue(level, appName, weekRange, currentEROAS, currentProfit, identifier = '', sourceApp = '') {
    if (!this.shouldRecordMetric(currentEROAS, currentProfit)) return;
    
    const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
    const cache = this.loadAllInitialValues();
    
    if (cache[key] !== undefined) {
      return;
    }
    
    try {
      const sheet = this.getOrCreateCacheSheet();
      const lastRow = sheet.getLastRow();
      
      const validEROAS = this.isValidEROAS(currentEROAS) ? currentEROAS : '';
      const validProfit = this.isValidProfit(currentProfit) ? currentProfit : '';
      
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
          profit: validProfit !== '' ? parseFloat(validProfit) : null
        };
        this.rowIndexCache[key] = lastRow + 1;
      }
    } catch (e) {
      console.error('Error saving initial metrics value:', e);
    }
  }

  // ========== СТРАТЕГИИ ОБРАБОТКИ ДАННЫХ ==========
  processWeekTotals(week, appName, campaigns, newValues, updateRequests) {
    const weekRange = `${week.weekStart} - ${week.weekEnd}`;
    const weekStartDate = weekRange.split(' - ')[0];
    
    if (!this.canRecordInitialForWeek(weekStartDate)) return;
    
    const weekTotals = calculateWeekTotals(campaigns);
    const weekKey = this.createKey('WEEK', appName, weekRange, '', '');
    const existingEntry = this.memoryCache[weekKey];
    
    if (!existingEntry && this.shouldRecordMetric(weekTotals.avgEROASD730, weekTotals.totalProfit)) {
      newValues.push([
        'WEEK', appName, weekRange, '', '',
        this.isValidEROAS(weekTotals.avgEROASD730) ? weekTotals.avgEROASD730 : '',
        new Date(),
        weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : ''
      ]);
    } else if (existingEntry && this.needsUpdate(existingEntry, weekTotals.avgEROASD730, weekTotals.totalProfit)) {
      updateRequests.push({
        key: weekKey,
        rowIndex: this.rowIndexCache[weekKey],
        eROAS: existingEntry.eROAS || (this.isValidEROAS(weekTotals.avgEROASD730) ? weekTotals.avgEROASD730 : null),
        profit: existingEntry.profit || (weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : null)
      });
    }
  }

  processNetworkData(week, appName, network, newValues, updateRequests) {
    const weekRange = `${week.weekStart} - ${week.weekEnd}`;
    const weekStartDate = weekRange.split(' - ')[0];
    
    if (!this.canRecordInitialForWeek(weekStartDate)) return;
    
    const networkTotals = calculateWeekTotals(network.campaigns);
    const networkKey = this.createKey('NETWORK', appName, weekRange, network.networkId, network.networkName);
    const existingEntry = this.memoryCache[networkKey];
    
    if (!existingEntry && this.shouldRecordMetric(networkTotals.avgEROASD730, networkTotals.totalProfit)) {
      newValues.push([
        'NETWORK', appName, weekRange, network.networkId, network.networkName,
        this.isValidEROAS(networkTotals.avgEROASD730) ? networkTotals.avgEROASD730 : '',
        new Date(),
        networkTotals.totalProfit !== 0 ? networkTotals.totalProfit : ''
      ]);
    } else if (existingEntry && this.needsUpdate(existingEntry, networkTotals.avgEROASD730, networkTotals.totalProfit)) {
      updateRequests.push({
        key: networkKey,
        rowIndex: this.rowIndexCache[networkKey],
        eROAS: existingEntry.eROAS || (this.isValidEROAS(networkTotals.avgEROASD730) ? networkTotals.avgEROASD730 : null),
        profit: existingEntry.profit || (networkTotals.totalProfit !== 0 ? networkTotals.totalProfit : null)
      });
    }
    
    // Обработка кампаний
    network.campaigns.forEach(campaign => {
      this.processCampaignData(week, appName, campaign, newValues, updateRequests);
    });
  }

  processSourceAppData(week, appName, sourceApp, newValues, updateRequests) {
    const weekRange = `${week.weekStart} - ${week.weekEnd}`;
    const weekStartDate = weekRange.split(' - ')[0];
    
    if (!this.canRecordInitialForWeek(weekStartDate)) return;
    
    const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
    const sourceAppKey = this.createKey('SOURCE_APP', appName, weekRange, sourceApp.sourceAppId, sourceApp.sourceAppName);
    const existingEntry = this.memoryCache[sourceAppKey];
    
    if (!existingEntry && this.shouldRecordMetric(sourceAppTotals.avgEROASD730, sourceAppTotals.totalProfit)) {
      newValues.push([
        'SOURCE_APP', appName, weekRange, sourceApp.sourceAppId, sourceApp.sourceAppName,
        this.isValidEROAS(sourceAppTotals.avgEROASD730) ? sourceAppTotals.avgEROASD730 : '',
        new Date(),
        sourceAppTotals.totalProfit !== 0 ? sourceAppTotals.totalProfit : ''
      ]);
    } else if (existingEntry && this.needsUpdate(existingEntry, sourceAppTotals.avgEROASD730, sourceAppTotals.totalProfit)) {
      updateRequests.push({
        key: sourceAppKey,
        rowIndex: this.rowIndexCache[sourceAppKey],
        eROAS: existingEntry.eROAS || (this.isValidEROAS(sourceAppTotals.avgEROASD730) ? sourceAppTotals.avgEROASD730 : null),
        profit: existingEntry.profit || (sourceAppTotals.totalProfit !== 0 ? sourceAppTotals.totalProfit : null)
      });
    }
    
    // Обработка кампаний
    sourceApp.campaigns.forEach(campaign => {
      this.processCampaignData(week, appName, campaign, newValues, updateRequests);
    });
  }

  processCampaignData(week, appName, campaign, newValues, updateRequests) {
    const weekRange = `${week.weekStart} - ${week.weekEnd}`;
    const weekStartDate = weekRange.split(' - ')[0];
    
    if (!this.canRecordInitialForWeek(weekStartDate)) return;
    
    const campaignKey = this.createKey('CAMPAIGN', appName, weekRange, campaign.campaignId, campaign.sourceApp);
    const existingEntry = this.memoryCache[campaignKey];
    
    if (!existingEntry && this.shouldRecordMetric(campaign.eRoasForecastD730, campaign.eProfitForecast)) {
      newValues.push([
        'CAMPAIGN', appName, weekRange, campaign.campaignId, campaign.sourceApp,
        this.isValidEROAS(campaign.eRoasForecastD730) ? campaign.eRoasForecastD730 : '',
        new Date(),
        campaign.eProfitForecast !== 0 ? campaign.eProfitForecast : ''
      ]);
    } else if (existingEntry && this.needsUpdate(existingEntry, campaign.eRoasForecastD730, campaign.eProfitForecast)) {
      updateRequests.push({
        key: campaignKey,
        rowIndex: this.rowIndexCache[campaignKey],
        eROAS: existingEntry.eROAS || (this.isValidEROAS(campaign.eRoasForecastD730) ? campaign.eRoasForecastD730 : null),
        profit: existingEntry.profit || (campaign.eProfitForecast !== 0 ? campaign.eProfitForecast : null)
      });
    }
  }

  processIncentTrafficData(appData, newValues, updateRequests) {
    Object.values(appData).forEach(network => {
      Object.values(network.countries).forEach(country => {
        Object.values(country.campaigns).forEach(campaign => {
          Object.values(campaign.weeks).forEach(week => {
            const weekRange = `${week.weekStart} - ${week.weekEnd}`;
            const weekStartDate = weekRange.split(' - ')[0];
            
            if (!this.canRecordInitialForWeek(weekStartDate)) return;
            
            const allDataPoints = week.data || [];
            const weekTotals = calculateWeekTotals(allDataPoints.map(d => ({
              ...d,
              campaignId: campaign.campaignId,
              campaignName: campaign.campaignName
            })));
            
            const weekKey = this.createKey('WEEK', network.networkName, weekRange, '', '');
            const existingEntry = this.memoryCache[weekKey];
            
            if (!existingEntry && this.shouldRecordMetric(weekTotals.avgEROASD730, weekTotals.totalProfit)) {
              newValues.push([
                'WEEK', network.networkName, weekRange, '', '',
                this.isValidEROAS(weekTotals.avgEROASD730) ? weekTotals.avgEROASD730 : '',
                new Date(),
                weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : ''
              ]);
            } else if (existingEntry && this.needsUpdate(existingEntry, weekTotals.avgEROASD730, weekTotals.totalProfit)) {
              updateRequests.push({
                key: weekKey,
                rowIndex: this.rowIndexCache[weekKey],
                eROAS: existingEntry.eROAS || (this.isValidEROAS(weekTotals.avgEROASD730) ? weekTotals.avgEROASD730 : null),
                profit: existingEntry.profit || (weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : null)
              });
            }
          });
        });
      });
    });
  }

  processOverallData(appData, newValues, updateRequests) {
    Object.values(appData).forEach(app => {
      Object.values(app.weeks).forEach(week => {
        const allCampaigns = [];
        Object.values(week.networks).forEach(network => {
          allCampaigns.push(...network.campaigns);
        });
        
        this.processWeekTotals(week, app.appName, allCampaigns, newValues, updateRequests);
        
        Object.values(week.networks).forEach(network => {
          this.processNetworkData(week, app.appName, network, newValues, updateRequests);
        });
      });
    });
  }

  processTrickyData(appData, newValues, updateRequests) {
    Object.values(appData).forEach(app => {
      Object.values(app.weeks).forEach(week => {
        if (week.sourceApps) {
          const allCampaigns = [];
          Object.values(week.sourceApps).forEach(sourceApp => {
            allCampaigns.push(...sourceApp.campaigns);
            this.processSourceAppData(week, app.appName, sourceApp, newValues, updateRequests);
          });
          
          this.processWeekTotals(week, app.appName, allCampaigns, newValues, updateRequests);
        } else if (week.campaigns) {
          this.processWeekTotals(week, app.appName, week.campaigns, newValues, updateRequests);
          
          week.campaigns.forEach(campaign => {
            this.processCampaignData(week, app.appName, campaign, newValues, updateRequests);
          });
        }
      });
    });
  }

  processDefaultData(appData, newValues, updateRequests) {
    Object.values(appData).forEach(app => {
      Object.values(app.weeks).forEach(week => {
        const campaigns = week.campaigns || [];
        this.processWeekTotals(week, app.appName, campaigns, newValues, updateRequests);
        
        campaigns.forEach(campaign => {
          this.processCampaignData(week, app.appName, campaign, newValues, updateRequests);
        });
      });
    });
  }

  recordInitialValuesFromData(appData) {
    console.log(`${this.projectName}: Recording initial metrics values for all weeks in data`);
    
    this.initDateCache();
    this.loadAllInitialValues();
    
    const newValues = [];
    const updateRequests = [];
    
    // Выбор стратегии по проекту
    switch(this.projectName) {
      case 'INCENT_TRAFFIC':
        this.processIncentTrafficData(appData, newValues, updateRequests);
        break;
      case 'OVERALL':
        this.processOverallData(appData, newValues, updateRequests);
        break;
      case 'TRICKY':
        this.processTrickyData(appData, newValues, updateRequests);
        break;
      default:
        this.processDefaultData(appData, newValues, updateRequests);
    }
    
    // Сохранение новых значений
    if (newValues.length > 0) {
      try {
        const sheet = this.getOrCreateCacheSheet();
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1, newValues.length, 8).setValues(newValues);
        
        newValues.forEach((row, index) => {
          const [level, appName, weekRange, identifier, sourceApp, initialEROAS, dateRecorded, initialProfit] = row;
          const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
          if (this.memoryCache) {
            this.memoryCache[key] = {
              eROAS: this.isValidEROAS(initialEROAS) ? parseFloat(initialEROAS) : null,
              profit: this.isValidProfit(initialProfit) ? parseFloat(initialProfit) : null
            };
            this.rowIndexCache[key] = lastRow + 1 + index;
          }
        });
        
        console.log(`${this.projectName}: Added ${newValues.length} new initial metrics values`);
      } catch (e) {
        console.error(`Error batch saving initial metrics values for ${this.projectName}:`, e);
      }
    }
    
    // Обработка обновлений
    if (updateRequests.length > 0) {
      try {
        this.processUpdateRequests(updateRequests);
        console.log(`${this.projectName}: Updated ${updateRequests.length} existing entries with missing values`);
      } catch (e) {
        console.error(`Error updating existing entries for ${this.projectName}:`, e);
      }
    }
    
    if (newValues.length === 0 && updateRequests.length === 0) {
      console.log(`${this.projectName}: No new initial metrics values to record`);
    }
  }

  processUpdateRequests(updateRequests) {
    const validRequests = updateRequests.filter(update => 
      update.rowIndex && update.rowIndex > 1 && 
      (update.eROAS !== null || update.profit !== null)
    );
    
    if (validRequests.length === 0) return;
    
    const spreadsheetId = this.cacheSpreadsheet.getId();
    const batchUpdateData = [];
    
    validRequests.forEach(update => {
      if (update.eROAS !== null && update.eROAS !== undefined) {
        batchUpdateData.push({
          range: `${this.cacheSheet.getName()}!F${update.rowIndex}`,
          values: [[update.eROAS]]
        });
      }
      if (update.profit !== null && update.profit !== undefined) {
        batchUpdateData.push({
          range: `${this.cacheSheet.getName()}!H${update.rowIndex}`,
          values: [[update.profit]]
        });
      }
    });
    
    if (batchUpdateData.length > 0) {
      const batchUpdateRequest = {
        valueInputOption: 'RAW',
        data: batchUpdateData
      };
      
      Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest, spreadsheetId);
      
      validRequests.forEach(request => {
        if (this.memoryCache && this.memoryCache[request.key]) {
          if (request.eROAS !== null) this.memoryCache[request.key].eROAS = request.eROAS;
          if (request.profit !== null) this.memoryCache[request.key].profit = request.profit;
        }
      });
    }
  }

  // ========== МЕТОДЫ ФОРМАТИРОВАНИЯ (сохранены для совместимости) ==========
  formatEROASWithInitial(level, appName, weekRange, currentEROAS, identifier = '', sourceApp = '') {
    return this.formatMetricWithInitial(level, appName, weekRange, currentEROAS, 'eROAS', identifier, sourceApp);
  }

  formatProfitWithInitial(level, appName, weekRange, currentProfit, identifier = '', sourceApp = '') {
    return this.formatMetricWithInitial(level, appName, weekRange, currentProfit, 'profit', identifier, sourceApp);
  }

  clearMemoryCache() {
    this.memoryCache = null;
    this.rowIndexCache = null;
    this.dateCache = null;
  }
}

// Глобальная функция для совместимости
function clearAllInitialMetricsMemoryCaches() {
  console.log('Clearing all initial metrics memory caches...');
}