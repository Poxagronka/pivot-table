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

  /**
   * Получает или создает лист для проекта
   */
  getOrCreateCacheSheet() {
    if (this.cacheSheet) return this.cacheSheet;
    
    const sheetName = `InitialEROAS_${this.projectName}`;
    this.cacheSheet = this.cacheSpreadsheet.getSheetByName(sheetName);
    
    if (!this.cacheSheet) {
      this.cacheSheet = this.cacheSpreadsheet.insertSheet(sheetName);
      // Headers: Level, AppName, WeekRange, Identifier, SourceApp, InitialEROAS, DateRecorded
      this.cacheSheet.getRange(1, 1, 1, 7).setValues([['Level', 'AppName', 'WeekRange', 'Identifier', 'SourceApp', 'InitialEROAS', 'DateRecorded']]);
      this.cacheSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#f0f0f0');
    }
    
    return this.cacheSheet;
  }

  /**
   * Создает уникальный ключ для записи
   */
  createKey(level, appName, weekRange, identifier = '', sourceApp = '') {
    return `${level}|||${appName}|||${weekRange}|||${identifier}|||${sourceApp}`;
  }

  /**
   * Загружает все закешированные значения в память
   */
  loadAllInitialValues() {
    // Если уже загружено в память, возвращаем кеш
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

  /**
   * Сохраняет первоначальное значение eROAS
   */
  saveInitialValue(level, appName, weekRange, currentEROAS, identifier = '', sourceApp = '') {
    if (currentEROAS === null || currentEROAS === undefined || currentEROAS === '' || currentEROAS === 0) return;
    
    const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
    const cache = this.loadAllInitialValues();
    
    // Если значение уже существует в кеше, не перезаписываем
    if (cache[key] !== undefined) {
      return;
    }
    
    try {
      const sheet = this.getOrCreateCacheSheet();
      const lastRow = sheet.getLastRow();
      
      // Добавляем новую запись
      sheet.getRange(lastRow + 1, 1, 1, 7).setValues([[
        level,
        appName,
        weekRange,
        identifier || '',
        sourceApp || '',
        currentEROAS,
        new Date()
      ]]);
      
      // Обновляем память кеш
      if (this.memoryCache) {
        this.memoryCache[key] = currentEROAS;
      }
    } catch (e) {
      console.error('Error saving initial eROAS value:', e);
    }
  }

  /**
   * Записывает первоначальные значения из обработанных данных
   */
  recordInitialValuesFromData(appData) {
    const today = new Date();
    const dayOfWeek = today.getDay();
    const shouldIncludeLastWeek = dayOfWeek >= 2 || dayOfWeek === 0;
    
    if (!shouldIncludeLastWeek) {
      console.log(`${this.projectName}: Skipping initial eROAS recording - not Tuesday yet`);
      return;
    }
    
    const lastWeekStart = formatDateForAPI(getMondayOfWeek(new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000)));
    console.log(`${this.projectName}: Checking for initial eROAS values for week ${lastWeekStart}`);
    
    // Загружаем все существующие значения в память один раз
    this.loadAllInitialValues();
    
    // Собираем все новые значения в массив для батч-записи
    const newValues = [];
    
    if (this.projectName === 'INCENT_TRAFFIC') {
      // Специальная обработка для INCENT_TRAFFIC (network → week → app)
      Object.values(appData).forEach(network => {
        Object.values(network.weeks).forEach(week => {
          if (week.weekStart === lastWeekStart) {
            const weekRange = `${week.weekStart} - ${week.weekEnd}`;
            
            // Week level для сетки
            const allCampaigns = [];
            Object.values(week.apps).forEach(app => {
              allCampaigns.push(...app.campaigns);
            });
            const weekTotals = calculateWeekTotals(allCampaigns);
            
            const weekKey = this.createKey('WEEK', network.networkName, weekRange, '', '');
            if (!this.memoryCache[weekKey] && weekTotals.avgEROASD730 > 0) {
              newValues.push(['WEEK', network.networkName, weekRange, '', '', weekTotals.avgEROASD730, new Date()]);
            }
            
            // App level
            Object.values(week.apps).forEach(app => {
              const appTotals = calculateWeekTotals(app.campaigns);
              const appKey = this.createKey('APP', network.networkName, weekRange, app.appId, app.appName);
              if (!this.memoryCache[appKey] && appTotals.avgEROASD730 > 0) {
                newValues.push(['APP', network.networkName, weekRange, app.appId, app.appName, appTotals.avgEROASD730, new Date()]);
              }
            });
          }
        });
      });
    } else {
      // Обычная обработка для других проектов
      Object.values(appData).forEach(app => {
        Object.values(app.weeks).forEach(week => {
          if (week.weekStart === lastWeekStart) {
            const weekRange = `${week.weekStart} - ${week.weekEnd}`;
            
            if (this.projectName === 'TRICKY' && week.sourceApps) {
              // TRICKY с source apps
              const allCampaigns = [];
              Object.values(week.sourceApps).forEach(sourceApp => {
                allCampaigns.push(...sourceApp.campaigns);
              });
              const weekTotals = calculateWeekTotals(allCampaigns);
              
              // Week level
              const weekKey = this.createKey('WEEK', app.appName, weekRange, '', '');
              if (!this.memoryCache[weekKey] && weekTotals.avgEROASD730 > 0) {
                newValues.push(['WEEK', app.appName, weekRange, '', '', weekTotals.avgEROASD730, new Date()]);
              }
              
              // Source app level
              Object.values(week.sourceApps).forEach(sourceApp => {
                const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
                const sourceAppKey = this.createKey('SOURCE_APP', app.appName, weekRange, sourceApp.sourceAppId, sourceApp.sourceAppName);
                if (!this.memoryCache[sourceAppKey] && sourceAppTotals.avgEROASD730 > 0) {
                  newValues.push(['SOURCE_APP', app.appName, weekRange, sourceApp.sourceAppId, sourceApp.sourceAppName, sourceAppTotals.avgEROASD730, new Date()]);
                }
                
                // Campaign level
                sourceApp.campaigns.forEach(campaign => {
                  const campaignKey = this.createKey('CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp);
                  if (!this.memoryCache[campaignKey] && campaign.eRoasForecastD730 > 0) {
                    newValues.push(['CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp, campaign.eRoasForecastD730, new Date()]);
                  }
                });
              });
            } else if (this.projectName === 'OVERALL' && week.networks) {
              // OVERALL с networks
              const allCampaigns = [];
              Object.values(week.networks).forEach(network => {
                allCampaigns.push(...network.campaigns);
              });
              const weekTotals = calculateWeekTotals(allCampaigns);
              
              // Week level
              const weekKey = this.createKey('WEEK', app.appName, weekRange, '', '');
              if (!this.memoryCache[weekKey] && weekTotals.avgEROASD730 > 0) {
                newValues.push(['WEEK', app.appName, weekRange, '', '', weekTotals.avgEROASD730, new Date()]);
              }
              
              // Network level
              Object.values(week.networks).forEach(network => {
                const networkTotals = calculateWeekTotals(network.campaigns);
                const networkKey = this.createKey('NETWORK', app.appName, weekRange, network.networkId, network.networkName);
                if (!this.memoryCache[networkKey] && networkTotals.avgEROASD730 > 0) {
                  newValues.push(['NETWORK', app.appName, weekRange, network.networkId, network.networkName, networkTotals.avgEROASD730, new Date()]);
                }
              });
            } else {
              // Обычные проекты
              const weekTotals = calculateWeekTotals(week.campaigns);
              
              // Week level
              const weekKey = this.createKey('WEEK', app.appName, weekRange, '', '');
              if (!this.memoryCache[weekKey] && weekTotals.avgEROASD730 > 0) {
                newValues.push(['WEEK', app.appName, weekRange, '', '', weekTotals.avgEROASD730, new Date()]);
              }
              
              // Campaign level
              if (week.campaigns) {
                week.campaigns.forEach(campaign => {
                  const campaignKey = this.createKey('CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp);
                  if (!this.memoryCache[campaignKey] && campaign.eRoasForecastD730 > 0) {
                    newValues.push(['CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp, campaign.eRoasForecastD730, new Date()]);
                  }
                });
              }
            }
          }
        });
      });
    }
    
    // Батч-запись новых значений
    if (newValues.length > 0) {
      try {
        const sheet = this.getOrCreateCacheSheet();
        const lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1, newValues.length, 7).setValues(newValues);
        
        // Обновляем память кеш
        newValues.forEach(row => {
          const [level, appName, weekRange, identifier, sourceApp, initialEROAS] = row;
          const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
          if (this.memoryCache) {
            this.memoryCache[key] = initialEROAS;
          }
        });
        
        console.log(`${this.projectName}: Recorded ${newValues.length} initial eROAS values for week ${lastWeekStart}`);
      } catch (e) {
        console.error(`Error batch saving initial eROAS values for ${this.projectName}:`, e);
      }
    }
  }

  /**
   * Форматирует значение eROAS с первоначальным значением
   */
  formatEROASWithInitial(level, appName, weekRange, currentEROAS, identifier = '', sourceApp = '') {
    const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
    const cache = this.loadAllInitialValues();
    
    const initialValue = cache[key];
    const currentValue = Math.round(currentEROAS);
    
    if (initialValue !== undefined) {
      const initialRounded = Math.round(initialValue);
      return `${initialRounded}% → ${currentValue}%`;
    } else {
      // Если первоначального значения нет, показываем текущее дважды
      return `${currentValue}% → ${currentValue}%`;
    }
  }

  /**
   * Очищает кеш в памяти (не трогает данные в таблице)
   */
  clearMemoryCache() {
    this.memoryCache = null;
  }
}

/**
 * Глобальная функция для очистки памяти всех кешей eROAS
 */
function clearAllInitialEROASMemoryCaches() {
  console.log('Clearing all initial eROAS memory caches...');
  // Эта функция будет вызываться между обновлениями проектов
}