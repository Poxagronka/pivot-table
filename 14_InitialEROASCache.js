/**
 * Initial Metrics Cache Management - ИСПРАВЛЕНО: единая логика для eROAS и eProfit + поддержка отрицательной прибыли
 */

const INITIAL_METRICS_CACHE_SPREADSHEET_ID = '1JBYtINHH7yLwdsfCPV3q3sj6NlP3WmftsPvbdfzTWdU';

class InitialMetricsCache {
  constructor(projectName = null) {
    this.projectName = projectName ? projectName.toUpperCase() : CURRENT_PROJECT;
    this.cacheSpreadsheet = null;
    this.cacheSheet = null;
    this.memoryCache = null;
    this.rowIndexCache = null;
    
    try {
      this.cacheSpreadsheet = SpreadsheetApp.openById(INITIAL_METRICS_CACHE_SPREADSHEET_ID);
    } catch (e) {
      console.error('Failed to open Initial Metrics cache spreadsheet:', e);
      throw new Error('Cannot access Initial Metrics cache spreadsheet. Check ID and permissions.');
    }
  }

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
    const lastCol = Math.max(8, sheet.getLastColumn());
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    
    data.forEach((row, index) => {
      const [level, appName, weekRange, identifier, sourceApp, initialEROAS, dateRecorded, initialProfit] = row;
      if (initialEROAS || initialProfit !== '') {
        const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
        cache[key] = {
          eROAS: (initialEROAS && initialEROAS !== '' && initialEROAS > 0) ? parseFloat(initialEROAS) : null,
          profit: (initialProfit !== '' && initialProfit !== null && initialProfit !== undefined) ? parseFloat(initialProfit) : null
        };
        this.rowIndexCache[key] = index + 2;
      }
    });
    
    // Добавляем legacy ключи для старых названий
    if (typeof APP_NAME_LEGACY !== 'undefined') {
      Object.keys(APP_NAME_LEGACY).forEach(newName => {
        const oldName = APP_NAME_LEGACY[newName];
        data.forEach((row, index) => {
          const [level, appName, weekRange, identifier, sourceApp, initialEROAS, dateRecorded, initialProfit] = row;
          if (appName === oldName && (initialEROAS || initialProfit !== '')) {
            const legacyKey = this.createKey(level, newName, weekRange, identifier, sourceApp);
            if (!cache[legacyKey]) {
              cache[legacyKey] = {
                eROAS: (initialEROAS && initialEROAS !== '' && initialEROAS > 0) ? parseFloat(initialEROAS) : null,
                profit: (initialProfit !== '' && initialProfit !== null && initialProfit !== undefined) ? parseFloat(initialProfit) : null
              };
              // Не добавляем в rowIndexCache, так как это виртуальная запись
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
    if ((!currentEROAS || currentEROAS <= 0) && (currentProfit === null || currentProfit === undefined || currentProfit === '')) return;
    
    const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
    const cache = this.loadAllInitialValues();
    
    if (cache[key] !== undefined) {
      return;
    }
    
    try {
      const sheet = this.getOrCreateCacheSheet();
      const lastRow = sheet.getLastRow();
      
      const validEROAS = (currentEROAS && currentEROAS > 0) ? currentEROAS : '';
      const validProfit = (currentProfit !== null && currentProfit !== undefined && currentProfit !== '') ? currentProfit : '';
      
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

  recordInitialValuesFromData(appData) {
    console.log(`${this.projectName}: Recording initial metrics values for all weeks in data`);
    
    this.loadAllInitialValues();
    
    const newValues = [];
    const updateRequests = [];
    
    if (this.projectName === 'INCENT_TRAFFIC') {
      Object.values(appData).forEach(network => {
        Object.values(network.countries).forEach(country => {
          Object.values(country.campaigns).forEach(campaign => {
            Object.values(campaign.weeks).forEach(week => {
              const weekRange = `${week.weekStart} - ${week.weekEnd}`;
              
              const allDataPoints = week.data || [];
              const weekTotals = calculateWeekTotals(allDataPoints.map(d => ({
                ...d,
                campaignId: campaign.campaignId,
                campaignName: campaign.campaignName
              })));
              
              const weekKey = this.createKey('WEEK', network.networkName, weekRange, '', '');
              const existingEntry = this.memoryCache[weekKey];
              
          if (!existingEntry) {
            if (weekTotals.avgEROASD730 > 0 || weekTotals.totalProfit !== 0) {
              newValues.push(['WEEK', network.networkName, weekRange, '', '', 
                weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : '', 
                new Date(), 
                weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : ''
              ]);
            }
          } else {
            const needsUpdate = (weekTotals.avgEROASD730 > 0 && !existingEntry.eROAS) || 
                               (weekTotals.totalProfit !== 0 && !existingEntry.profit);
            if (needsUpdate) {
              updateRequests.push({
                key: weekKey,
                rowIndex: this.rowIndexCache[weekKey],
                eROAS: existingEntry.eROAS || (weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : null),
                profit: existingEntry.profit || (weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : null)
              });
            }
          }
            });
          });
        });
      });
    } else if (this.projectName === 'OVERALL') {
      Object.values(appData).forEach(app => {
        Object.values(app.weeks).forEach(week => {
          const weekRange = `${week.weekStart} - ${week.weekEnd}`;
          
          const allCampaigns = [];
          Object.values(week.networks).forEach(network => {
            allCampaigns.push(...network.campaigns);
          });
          const weekTotals = calculateWeekTotals(allCampaigns);
          
          const weekKey = this.createKey('WEEK', app.appName, weekRange, '', '');
          const existingEntry = this.memoryCache[weekKey];
          
          if (!existingEntry) {
            if (weekTotals.avgEROASD730 > 0 || weekTotals.totalProfit !== 0) {
              newValues.push(['WEEK', app.appName, weekRange, '', '', 
                weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : '', 
                new Date(), 
                weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : ''
              ]);
            }
          } else {
            const needsUpdate = (weekTotals.avgEROASD730 > 0 && !existingEntry.eROAS) || 
                               (weekTotals.totalProfit !== 0 && !existingEntry.profit);
            if (needsUpdate) {
              updateRequests.push({
                key: weekKey,
                rowIndex: this.rowIndexCache[weekKey],
                eROAS: existingEntry.eROAS || (weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : null),
                profit: existingEntry.profit || (weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : null)
              });
            }
          }
          
          Object.values(week.networks).forEach(network => {
            const networkTotals = calculateWeekTotals(network.campaigns);
            const networkKey = this.createKey('NETWORK', app.appName, weekRange, network.networkId, network.networkName);
            const existingEntry = this.memoryCache[networkKey];
            
            if (!existingEntry) {
              if (networkTotals.avgEROASD730 > 0 || networkTotals.totalProfit !== 0) {
                newValues.push(['NETWORK', app.appName, weekRange, network.networkId, network.networkName, 
                  networkTotals.avgEROASD730 > 0 ? networkTotals.avgEROASD730 : '', 
                  new Date(), 
                  networkTotals.totalProfit !== 0 ? networkTotals.totalProfit : ''
                ]);
              }
            } else {
              const needsUpdate = (networkTotals.avgEROASD730 > 0 && !existingEntry.eROAS) || 
                                 (networkTotals.totalProfit !== 0 && !existingEntry.profit);
              if (needsUpdate) {
                updateRequests.push({
                  key: networkKey,
                  rowIndex: this.rowIndexCache[networkKey],
                  eROAS: existingEntry.eROAS || (networkTotals.avgEROASD730 > 0 ? networkTotals.avgEROASD730 : null),
                  profit: existingEntry.profit || (networkTotals.totalProfit !== 0 ? networkTotals.totalProfit : null)
                });
              }
            }
          });
        });
      });
    } else if (this.projectName === 'APPLOVIN_TEST') {
      // Обработка APPLOVIN_TEST с группировкой по странам
      Object.values(appData).forEach(app => {
        if (app.campaignGroups) {
          Object.values(app.campaignGroups).forEach(campaignGroup => {
            Object.entries(campaignGroup.weeks).forEach(([weekKey, week]) => {
              const weekRange = `${week.weekStart} - ${week.weekEnd}`;
              
              // Уровень недели (агрегация всех стран)
              if (week.countries) {
                const allCampaigns = [];
                Object.values(week.countries).forEach(country => {
                  allCampaigns.push(...country.campaigns);
                });
                const weekTotals = calculateWeekTotals(allCampaigns);
                
                const weekKey = this.createKey('WEEK', app.appName, weekRange, campaignGroup.campaignId, campaignGroup.campaignName);
                const existingEntry = this.memoryCache[weekKey];
                
                if (!existingEntry) {
                  if (weekTotals.avgEROASD730 > 0 || weekTotals.totalProfit !== 0) {
                    newValues.push(['WEEK', app.appName, weekRange, campaignGroup.campaignId, campaignGroup.campaignName, 
                      weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : '', 
                      new Date(), 
                      weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : ''
                    ]);
                  }
                } else {
                  const needsUpdate = (weekTotals.avgEROASD730 > 0 && !existingEntry.eROAS) || 
                                     (weekTotals.totalProfit !== 0 && !existingEntry.profit);
                  if (needsUpdate) {
                    updateRequests.push({
                      key: weekKey,
                      rowIndex: this.rowIndexCache[weekKey],
                      eROAS: existingEntry.eROAS || (weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : null),
                      profit: existingEntry.profit || (weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : null)
                    });
                  }
                }
                
                // Уровень стран
                Object.entries(week.countries).forEach(([countryCode, country]) => {
                  const countryCampaigns = country.campaigns || [];
                  const countryTotals = calculateWeekTotals(countryCampaigns);
                  
                  const countryKey = this.createKey('COUNTRY', app.appName, weekRange, 
                    `${campaignGroup.campaignId}_${countryCode}`, country.countryName);
                  const existingEntry = this.memoryCache[countryKey];
                  
                  if (!existingEntry) {
                    if (countryTotals.avgEROASD730 > 0 || countryTotals.totalProfit !== 0) {
                      newValues.push(['COUNTRY', app.appName, weekRange, 
                        `${campaignGroup.campaignId}_${countryCode}`, country.countryName, 
                        countryTotals.avgEROASD730 > 0 ? countryTotals.avgEROASD730 : '', 
                        new Date(), 
                        countryTotals.totalProfit !== 0 ? countryTotals.totalProfit : ''
                      ]);
                    }
                  } else {
                    const needsUpdate = (countryTotals.avgEROASD730 > 0 && !existingEntry.eROAS) || 
                                       (countryTotals.totalProfit !== 0 && !existingEntry.profit);
                    if (needsUpdate) {
                      updateRequests.push({
                        key: countryKey,
                        rowIndex: this.rowIndexCache[countryKey],
                        eROAS: existingEntry.eROAS || (countryTotals.avgEROASD730 > 0 ? countryTotals.avgEROASD730 : null),
                        profit: existingEntry.profit || (countryTotals.totalProfit !== 0 ? countryTotals.totalProfit : null)
                      });
                    }
                  }
                });
              }
            });
          });
        }
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
            const existingEntry = this.memoryCache[weekKey];
            
            if (!existingEntry) {
              if (weekTotals.avgEROASD730 > 0 || weekTotals.totalProfit !== 0) {
                newValues.push(['WEEK', app.appName, weekRange, '', '', 
                  weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : '', 
                  new Date(), 
                  weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : ''
                ]);
              }
            } else {
              const needsUpdate = (weekTotals.avgEROASD730 > 0 && !existingEntry.eROAS) || 
                                 (weekTotals.totalProfit !== 0 && !existingEntry.profit);
              if (needsUpdate) {
                updateRequests.push({
                  key: weekKey,
                  rowIndex: this.rowIndexCache[weekKey],
                  eROAS: existingEntry.eROAS || (weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : null),
                  profit: existingEntry.profit || (weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : null)
                });
              }
            }
            
            Object.values(week.sourceApps).forEach(sourceApp => {
              const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
              const sourceAppKey = this.createKey('SOURCE_APP', app.appName, weekRange, sourceApp.sourceAppId, sourceApp.sourceAppName);
              const existingEntry = this.memoryCache[sourceAppKey];
              
              if (!existingEntry) {
                if (sourceAppTotals.avgEROASD730 > 0 || sourceAppTotals.totalProfit !== 0) {
                  newValues.push(['SOURCE_APP', app.appName, weekRange, sourceApp.sourceAppId, sourceApp.sourceAppName, 
                    sourceAppTotals.avgEROASD730 > 0 ? sourceAppTotals.avgEROASD730 : '', 
                    new Date(), 
                    sourceAppTotals.totalProfit !== 0 ? sourceAppTotals.totalProfit : ''
                  ]);
                }
              } else {
                const needsUpdate = (sourceAppTotals.avgEROASD730 > 0 && !existingEntry.eROAS) || 
                                   (sourceAppTotals.totalProfit !== 0 && !existingEntry.profit);
                if (needsUpdate) {
                  updateRequests.push({
                    key: sourceAppKey,
                    rowIndex: this.rowIndexCache[sourceAppKey],
                    eROAS: existingEntry.eROAS || (sourceAppTotals.avgEROASD730 > 0 ? sourceAppTotals.avgEROASD730 : null),
                    profit: existingEntry.profit || (sourceAppTotals.totalProfit !== 0 ? sourceAppTotals.totalProfit : null)
                  });
                }
              }
              
              sourceApp.campaigns.forEach(campaign => {
                const campaignKey = this.createKey('CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp);
                const existingEntry = this.memoryCache[campaignKey];
                
                if (!existingEntry) {
                  if (campaign.eRoasForecastD730 > 0 || campaign.eProfitForecast !== 0) {
                    newValues.push(['CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp, 
                      campaign.eRoasForecastD730 > 0 ? campaign.eRoasForecastD730 : '', 
                      new Date(), 
                      campaign.eProfitForecast !== 0 ? campaign.eProfitForecast : ''
                    ]);
                  }
                } else {
                  const needsUpdate = (campaign.eRoasForecastD730 > 0 && !existingEntry.eROAS) || 
                                     (campaign.eProfitForecast !== 0 && !existingEntry.profit);
                  if (needsUpdate) {
                    updateRequests.push({
                      key: campaignKey,
                      rowIndex: this.rowIndexCache[campaignKey],
                      eROAS: existingEntry.eROAS || (campaign.eRoasForecastD730 > 0 ? campaign.eRoasForecastD730 : null),
                      profit: existingEntry.profit || (campaign.eProfitForecast !== 0 ? campaign.eProfitForecast : null)
                    });
                  }
                }
              });
            });
          } else {
            const weekTotals = calculateWeekTotals(week.campaigns || []);
            
            const weekKey = this.createKey('WEEK', app.appName, weekRange, '', '');
            const existingEntry = this.memoryCache[weekKey];
            
            if (!existingEntry) {
              if (weekTotals.avgEROASD730 > 0 || weekTotals.totalProfit !== 0) {
                newValues.push(['WEEK', app.appName, weekRange, '', '', 
                  weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : '', 
                  new Date(), 
                  weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : ''
                ]);
              }
            } else {
              const needsUpdate = (weekTotals.avgEROASD730 > 0 && !existingEntry.eROAS) || 
                                 (weekTotals.totalProfit !== 0 && !existingEntry.profit);
              if (needsUpdate) {
                updateRequests.push({
                  key: weekKey,
                  rowIndex: this.rowIndexCache[weekKey],
                  eROAS: existingEntry.eROAS || (weekTotals.avgEROASD730 > 0 ? weekTotals.avgEROASD730 : null),
                  profit: existingEntry.profit || (weekTotals.totalProfit !== 0 ? weekTotals.totalProfit : null)
                });
              }
            }
            
            if (week.campaigns) {
              week.campaigns.forEach(campaign => {
                const campaignKey = this.createKey('CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp);
                const existingEntry = this.memoryCache[campaignKey];
                
                if (!existingEntry) {
                  if (campaign.eRoasForecastD730 > 0 || campaign.eProfitForecast !== 0) {
                    newValues.push(['CAMPAIGN', app.appName, weekRange, campaign.campaignId, campaign.sourceApp, 
                      campaign.eRoasForecastD730 > 0 ? campaign.eRoasForecastD730 : '', 
                      new Date(), 
                      campaign.eProfitForecast !== 0 ? campaign.eProfitForecast : ''
                    ]);
                  }
                } else {
                  const needsUpdate = (campaign.eRoasForecastD730 > 0 && !existingEntry.eROAS) || 
                                     (campaign.eProfitForecast !== 0 && !existingEntry.profit);
                  if (needsUpdate) {
                    updateRequests.push({
                      key: campaignKey,
                      rowIndex: this.rowIndexCache[campaignKey],
                      eROAS: existingEntry.eROAS || (campaign.eRoasForecastD730 > 0 ? campaign.eRoasForecastD730 : null),
                      profit: existingEntry.profit || (campaign.eProfitForecast !== 0 ? campaign.eProfitForecast : null)
                    });
                  }
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
        
        newValues.forEach((row, index) => {
          const [level, appName, weekRange, identifier, sourceApp, initialEROAS, dateRecorded, initialProfit] = row;
          const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
          if (this.memoryCache) {
            this.memoryCache[key] = {
              eROAS: (initialEROAS && initialEROAS !== '' && initialEROAS > 0) ? parseFloat(initialEROAS) : null,
              profit: (initialProfit !== '' && initialProfit !== null && initialProfit !== undefined) ? parseFloat(initialProfit) : null
            };
            this.rowIndexCache[key] = lastRow + 1 + index;
          }
        });
        
        console.log(`${this.projectName}: Added ${newValues.length} new initial metrics values`);
      } catch (e) {
        console.error(`Error batch saving initial metrics values for ${this.projectName}:`, e);
      }
    }
    
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
    if (updateRequests.length === 0) return;
    
    const spreadsheetId = this.cacheSpreadsheet.getId();
    const batchUpdateData = [];
    
    updateRequests.forEach(update => {
      if (update.rowIndex && update.rowIndex > 1) {
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
      }
    });
    
    if (batchUpdateData.length > 0) {
      const batchUpdateRequest = {
        valueInputOption: 'RAW',
        data: batchUpdateData
      };
      
      Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest, spreadsheetId);
      
      updateRequests.forEach(request => {
        if (this.memoryCache && this.memoryCache[request.key]) {
          if (request.eROAS !== null) this.memoryCache[request.key].eROAS = request.eROAS;
          if (request.profit !== null) this.memoryCache[request.key].profit = request.profit;
        }
      });
    }
  }

  formatEROASWithInitial(level, appName, weekRange, currentEROAS, identifier = '', sourceApp = '') {
    const key = this.createKey(level, appName, weekRange, identifier, sourceApp);
    const cache = this.loadAllInitialValues();
    
    const metrics = cache[key];
    const currentValue = Math.round(currentEROAS);
    
    if (metrics && metrics.eROAS !== null && metrics.eROAS !== undefined) {
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
    
    if (metrics && metrics.profit !== null && metrics.profit !== undefined) {
      const initialRounded = Math.round(metrics.profit);
      return `${initialRounded}$ → ${currentValue}$`;
    } else {
      return `${currentValue}$ → ${currentValue}$`;
    }
  }

  clearMemoryCache() {
    this.memoryCache = null;
    this.rowIndexCache = null;
  }
}

function clearAllInitialMetricsMemoryCaches() {
  console.log('Clearing all initial metrics memory caches...');
}