const WEEK_TOTALS_CACHE = new Map();
const WOW_METRICS_CACHE = new Map();
const PRECOMPUTED_TOTALS = new Map();
const PRECOMPUTED_ROWS = new Map();
const WOW_KEYS_CACHE = new Map();

function buildUnifiedTable(data, tableData, formatData, wow, initialMetricsCache) {
  const startTime = Date.now();

  clearTableBuilderCaches();
  
  precomputeAllTotals(data);
  precomputeWoWCache(wow);
  
  let appsDbCache = null;
  if (CURRENT_PROJECT === 'TRICKY') {
    try {
      const appsDb = new AppsDatabase('TRICKY');
      appsDbCache = appsDb.loadFromCache();
    } catch (e) {
      console.error('Error loading AppsDatabase:', e);
      appsDbCache = {};
    }
  }
  
  if (CURRENT_PROJECT === 'APPLOVIN_TEST') {
    const appKeys = Object.keys(data).sort((a, b) => data[a].appName.localeCompare(data[b].appName));
    
    appKeys.forEach(appKey => {
      const app = data[appKey];
      
      formatData.push({ row: tableData.length + 1, type: 'APP' });
      const emptyRow = new Array(getUnifiedHeaders().length).fill('');
      emptyRow[COLUMN_CONFIG.COLUMNS.LEVEL - 1] = 'APP';
      emptyRow[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = app.appName;
      tableData.push(emptyRow);
      
      const campaignKeys = Object.keys(app.campaignGroups || {}).sort((a, b) => {
        const spendA = Object.values(app.campaignGroups[a].weeks).reduce((sum, w) => {
          return sum + Object.values(w.countries || {}).reduce((s, country) => 
            s + country.campaigns.reduce((cs, c) => cs + c.spend, 0), 0);
        }, 0);
        const spendB = Object.values(app.campaignGroups[b].weeks).reduce((sum, w) => {
          return sum + Object.values(w.countries || {}).reduce((s, country) => 
            s + country.campaigns.reduce((cs, c) => cs + c.spend, 0), 0);
        }, 0);
        return spendB - spendA;
      });
      
      campaignKeys.forEach(campaignKey => {
        const campaignGroup = app.campaignGroups[campaignKey];
        
        formatData.push({ row: tableData.length + 1, type: 'CAMPAIGN' });
        const campaignRow = new Array(getUnifiedHeaders().length).fill('');
        campaignRow[COLUMN_CONFIG.COLUMNS.LEVEL - 1] = 'CAMPAIGN';
        campaignRow[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = campaignGroup.campaignName;
        campaignRow[COLUMN_CONFIG.COLUMNS.ID - 1] = campaignGroup.campaignId;
        campaignRow[COLUMN_CONFIG.COLUMNS.GEO - 1] = ''; // Не показываем GEO на уровне кампании для APPLOVIN_TEST
        tableData.push(campaignRow);
        
        const weekKeys = Object.keys(campaignGroup.weeks).sort();
        weekKeys.forEach(weekKey => {
          const week = campaignGroup.weeks[weekKey];
          
          // Агрегируем данные всех стран для недели
          const allCountryCampaigns = [];
          Object.values(week.countries || {}).forEach(country => {
            allCountryCampaigns.push(...country.campaigns);
          });
          
          const weekTotals = calculateWeekTotals(allCountryCampaigns);
          const weekWoWKey = `${campaignGroup.campaignId}_${weekKey}`;
          const weekWoW = getOptimizedWoW(weekWoWKey, 'campaignWoW');
          const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
          const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
          const status = weekWoW.growthStatus || '';
          
          formatData.push({ row: tableData.length + 1, type: 'WEEK' });
          const weekRow = new Array(getUnifiedHeaders().length).fill('');
          weekRow[COLUMN_CONFIG.COLUMNS.LEVEL - 1] = 'WEEK';
          weekRow[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = `${week.weekStart} - ${week.weekEnd}`;
          weekRow[COLUMN_CONFIG.COLUMNS.ID - 1] = '';
          weekRow[COLUMN_CONFIG.COLUMNS.GEO - 1] = '';
          weekRow[COLUMN_CONFIG.COLUMNS.SPEND - 1] = formatSmartCurrency(weekTotals.totalSpend);
          weekRow[COLUMN_CONFIG.COLUMNS.SPEND_WOW - 1] = spendWoW;
          weekRow[COLUMN_CONFIG.COLUMNS.INSTALLS - 1] = weekTotals.totalInstalls;
          weekRow[COLUMN_CONFIG.COLUMNS.CPI - 1] = weekTotals.avgCpi.toFixed(3);
          
          const combinedRoas = `${weekTotals.avgRoasD1.toFixed(0)}% → ${weekTotals.avgRoasD3.toFixed(0)}% → ${weekTotals.avgRoasD7.toFixed(0)}% → ${weekTotals.avgRoasD14.toFixed(0)}% → ${weekTotals.avgRoasD30.toFixed(0)}%`;
          weekRow[COLUMN_CONFIG.COLUMNS.ROAS_COMBINED - 1] = combinedRoas;
          
          weekRow[COLUMN_CONFIG.COLUMNS.IPM - 1] = weekTotals.avgIpm.toFixed(1);
          weekRow[COLUMN_CONFIG.COLUMNS.RR_COMBINED - 1] = `${weekTotals.avgRrD1.toFixed(0)}% → ${weekTotals.avgRrD7.toFixed(0)}%`;
          weekRow[COLUMN_CONFIG.COLUMNS.EARPU - 1] = weekTotals.avgArpu.toFixed(3);
          weekRow[COLUMN_CONFIG.COLUMNS.EROAS_365 - 1] = `${weekTotals.avgERoas.toFixed(0)}%`;
          // Вместо простого форматирования, используем initialMetricsCache
          let weekEROAS730Display = `${weekTotals.avgEROASD730.toFixed(0)}%`;
          let weekEProfit730Display = formatSmartCurrency(weekTotals.totalProfit);

          if (initialMetricsCache && app.appName) {
            const weekRange = `${week.weekStart} - ${week.weekEnd}`;
            weekEROAS730Display = initialMetricsCache.formatEROASWithInitial(
              'WEEK', app.appName, weekRange, weekTotals.avgEROASD730, 
              campaignGroup.campaignId, campaignGroup.campaignName
            );
            weekEProfit730Display = initialMetricsCache.formatProfitWithInitial(
              'WEEK', app.appName, weekRange, weekTotals.totalProfit,
              campaignGroup.campaignId, campaignGroup.campaignName
            );
          }

          weekRow[COLUMN_CONFIG.COLUMNS.EROAS_730 - 1] = weekEROAS730Display;
          weekRow[COLUMN_CONFIG.COLUMNS.EPROFIT_730 - 1] = weekEProfit730Display;
          weekRow[COLUMN_CONFIG.COLUMNS.EPROFIT_WOW - 1] = profitWoW;
          weekRow[COLUMN_CONFIG.COLUMNS.GROWTH_STATUS - 1] = status;
          weekRow[COLUMN_CONFIG.COLUMNS.COMMENTS - 1] = '';
          
          tableData.push(weekRow);
          
          // Страны уже отсортированы в restructureToCampaignFirst
          Object.values(week.countries || {}).forEach(country => {
            if (!country.campaigns || country.campaigns.length === 0) return;
            const campaign = country.campaigns[0];
            
            const countryWoWKey = `${campaignGroup.campaignId}_${weekKey}_${country.countryCode}`;
            const countryWoW = getOptimizedWoW(countryWoWKey, 'countryWoW');
            const spendWoW = countryWoW.spendChangePercent !== undefined ? `${countryWoW.spendChangePercent.toFixed(0)}%` : '';
            const profitWoW = countryWoW.eProfitChangePercent !== undefined ? `${countryWoW.eProfitChangePercent.toFixed(0)}%` : '';
            const status = countryWoW.growthStatus || '';
            
            formatData.push({ row: tableData.length + 1, type: 'COUNTRY' });
            
            const countryRow = new Array(getUnifiedHeaders().length).fill('');
            countryRow[COLUMN_CONFIG.COLUMNS.LEVEL - 1] = 'COUNTRY';
            countryRow[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = country.countryName;
            countryRow[COLUMN_CONFIG.COLUMNS.ID - 1] = '';
            countryRow[COLUMN_CONFIG.COLUMNS.GEO - 1] = country.countryCode || 'OTHER';
            countryRow[COLUMN_CONFIG.COLUMNS.SPEND - 1] = formatSmartCurrency(campaign.spend || 0);
            countryRow[COLUMN_CONFIG.COLUMNS.SPEND_WOW - 1] = spendWoW;
            countryRow[COLUMN_CONFIG.COLUMNS.INSTALLS - 1] = campaign.installs || 0;
            countryRow[COLUMN_CONFIG.COLUMNS.CPI - 1] = campaign.cpi ? campaign.cpi.toFixed(3) : '0.000';
            
            const combinedRoas = `${(campaign.roasD1 || 0).toFixed(0)}% → ${(campaign.roasD3 || 0).toFixed(0)}% → ${(campaign.roasD7 || 0).toFixed(0)}% → ${(campaign.roasD14 || 0).toFixed(0)}% → ${(campaign.roasD30 || 0).toFixed(0)}%`;
            countryRow[COLUMN_CONFIG.COLUMNS.ROAS_COMBINED - 1] = combinedRoas;
            
            countryRow[COLUMN_CONFIG.COLUMNS.IPM - 1] = (campaign.ipm || 0).toFixed(1);
            countryRow[COLUMN_CONFIG.COLUMNS.RR_COMBINED - 1] = `${(campaign.rrD1 || 0).toFixed(0)}% → ${(campaign.rrD7 || 0).toFixed(0)}%`;
            countryRow[COLUMN_CONFIG.COLUMNS.EARPU - 1] = (campaign.eArpuForecast || 0).toFixed(3);
            countryRow[COLUMN_CONFIG.COLUMNS.EROAS_365 - 1] = `${(campaign.eRoasForecast || 0).toFixed(0)}%`;
            // Для каждой страны добавить обработку initial метрик
            let countryEROAS730Display = `${(campaign.eRoasForecastD730 || 0).toFixed(0)}%`;
            let countryEProfit730Display = formatSmartCurrency(campaign.eProfitForecast || 0);

            if (initialMetricsCache && app.appName) {
              const weekRange = `${week.weekStart} - ${week.weekEnd}`;
              countryEROAS730Display = initialMetricsCache.formatEROASWithInitial(
                'COUNTRY', app.appName, weekRange, campaign.eRoasForecastD730 || 0,
                `${campaignGroup.campaignId}_${country.countryCode}`, country.countryName
              );
              countryEProfit730Display = initialMetricsCache.formatProfitWithInitial(
                'COUNTRY', app.appName, weekRange, campaign.eProfitForecast || 0,
                `${campaignGroup.campaignId}_${country.countryCode}`, country.countryName
              );
            }

            countryRow[COLUMN_CONFIG.COLUMNS.EROAS_730 - 1] = countryEROAS730Display;
            countryRow[COLUMN_CONFIG.COLUMNS.EPROFIT_730 - 1] = countryEProfit730Display;
            countryRow[COLUMN_CONFIG.COLUMNS.EPROFIT_WOW - 1] = profitWoW;
            countryRow[COLUMN_CONFIG.COLUMNS.GROWTH_STATUS - 1] = status;
            countryRow[COLUMN_CONFIG.COLUMNS.COMMENTS - 1] = '';
            
            tableData.push(countryRow);
          });
        });
      });
    });
    
    console.log(`buildUnifiedTable completed: ${tableData.length} rows in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
    return;
  }
  
  // Далее идет оригинальный код для INCENT_TRAFFIC
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    const networkKeys = Object.keys(data).sort((a, b) => 
      data[a].networkName.localeCompare(data[b].networkName)
    );
    
    networkKeys.forEach(networkKey => {
      const network = data[networkKey];
      
      // Уровень NETWORK
      formatData.push({ row: tableData.length + 1, type: 'NETWORK' });
      const networkRow = new Array(getUnifiedHeaders().length).fill('');
      networkRow[COLUMN_CONFIG.COLUMNS.LEVEL - 1] = 'NETWORK';
      networkRow[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = network.networkName;
      tableData.push(networkRow);
      
      // Сортируем страны по spend (от большего к меньшему)
      const countryKeys = Object.keys(network.countries).sort((a, b) => {
        // Считаем общий spend для каждой страны
        const getCountrySpend = (country) => {
          let totalSpend = 0;
          Object.values(country.campaigns).forEach(campaign => {
            Object.values(campaign.weeks).forEach(week => {
              totalSpend += week.data.reduce((s, d) => s + d.spend, 0);
            });
          });
          return totalSpend;
        };
        
        const spendA = getCountrySpend(network.countries[a]);
        const spendB = getCountrySpend(network.countries[b]);
        return spendB - spendA; // От большего к меньшему
      });
      
      countryKeys.forEach(countryCode => {
        const country = network.countries[countryCode];
        
        // Уровень COUNTRY
        formatData.push({ row: tableData.length + 1, type: 'COUNTRY' });
        const countryRow = new Array(getUnifiedHeaders().length).fill('');
        countryRow[COLUMN_CONFIG.COLUMNS.LEVEL - 1] = 'COUNTRY';
        countryRow[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = country.countryName;
        countryRow[COLUMN_CONFIG.COLUMNS.GEO - 1] = country.countryCode;  // GEO колонка
        tableData.push(countryRow);
        
        // Сортируем кампании по общему spend
        const campaignKeys = Object.keys(country.campaigns).sort((a, b) => {
          const spendA = Object.values(country.campaigns[a].weeks).reduce((sum, w) => 
            sum + w.data.reduce((s, d) => s + d.spend, 0), 0);
          const spendB = Object.values(country.campaigns[b].weeks).reduce((sum, w) => 
            sum + w.data.reduce((s, d) => s + d.spend, 0), 0);
          return spendB - spendA;
        });
        
        campaignKeys.forEach(campaignId => {
          const campaign = country.campaigns[campaignId];
          
          // Уровень CAMPAIGN
          formatData.push({ row: tableData.length + 1, type: 'CAMPAIGN' });
          const campaignRow = new Array(getUnifiedHeaders().length).fill('');
          campaignRow[COLUMN_CONFIG.COLUMNS.LEVEL - 1] = 'CAMPAIGN';
          campaignRow[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = campaign.campaignName;
          campaignRow[COLUMN_CONFIG.COLUMNS.ID - 1] = campaign.campaignId;
          campaignRow[COLUMN_CONFIG.COLUMNS.GEO - 1] = campaign.geo || '';
          tableData.push(campaignRow);
          
          // Сортируем недели
          const weekKeys = Object.keys(campaign.weeks).sort();
          
          weekKeys.forEach(weekKey => {
            const week = campaign.weeks[weekKey];
            const weekData = week.data || [];
            
            // Получаем precomputed totals
            const totalsKey = `incent_week_${networkKey}_${countryCode}_${campaignId}_${weekKey}`;
            const weekTotals = getPrecomputedTotals(weekData, totalsKey);
            
            // Получаем WoW данные
            const wowKey = `${networkKey}_${countryCode}_${campaignId}_${weekKey}`;
            const weekWoW = getOptimizedWoW(wowKey, 'weekWoW');
            
            const spendWoW = weekWoW.spendChangePercent !== undefined ? 
              `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
            const profitWoW = weekWoW.eProfitChangePercent !== undefined ? 
              `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
            const status = weekWoW.growthStatus || '';
            
            // Уровень WEEK - только тут показываем метрики
            formatData.push({ row: tableData.length + 1, type: 'WEEK' });
            const weekRow = createUnifiedRow('WEEK', week, weekTotals, spendWoW, profitWoW, status, 
              network.networkName, initialMetricsCache);
            tableData.push(weekRow);
          });
        });
      });
    });
    
  } else {
    // Оригинальный код для остальных проектов
    const appKeys = Object.keys(data).sort((a, b) => data[a].appName.localeCompare(data[b].appName));
    
    appKeys.forEach((appKey, appIndex) => {
      const app = data[appKey];
      
      formatData.push({ row: tableData.length + 1, type: 'APP' });
      const emptyRow = new Array(getUnifiedHeaders().length).fill('');
      emptyRow[COLUMN_CONFIG.COLUMNS.LEVEL - 1] = 'APP';
      emptyRow[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = app.appName;
      tableData.push(emptyRow);

      // Для APPLOVIN_TEST структура другая: app.campaignGroups вместо app.weeks
      if (CURRENT_PROJECT === 'APPLOVIN_TEST' && app.campaignGroups) {
        // Сортируем кампании по общему spend
        const campaignKeys = Object.keys(app.campaignGroups).sort((a, b) => {
          const getTotalSpend = (campaign) => {
            return Object.values(campaign.weeks).reduce((sum, week) => {
              return sum + Object.values(week.countries || {}).reduce((wSum, country) => {
                return wSum + (country.campaigns || []).reduce((cSum, c) => cSum + c.spend, 0);
              }, 0);
            }, 0);
          };
          return getTotalSpend(app.campaignGroups[b]) - getTotalSpend(app.campaignGroups[a]);
        });
        
        campaignKeys.forEach(campaignKey => {
          const campaignGroup = app.campaignGroups[campaignKey];
          
          // Добавляем строку кампании
          formatData.push({ row: tableData.length + 1, type: 'CAMPAIGN' });
          const campaignEmptyRow = new Array(getUnifiedHeaders().length).fill('');
          campaignEmptyRow[COLUMN_CONFIG.COLUMNS.LEVEL - 1] = 'CAMPAIGN';
          campaignEmptyRow[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = campaignGroup.campaignName;
          tableData.push(campaignEmptyRow);
          
          // Обрабатываем недели для каждой кампании
          const weekKeys = Object.keys(campaignGroup.weeks).sort();
          
          weekKeys.forEach(weekKey => {
            const week = campaignGroup.weeks[weekKey];
            
            // Собираем все данные недели для totals
            const allWeekCampaigns = [];
            Object.values(week.countries || {}).forEach(country => {
              allWeekCampaigns.push(...(country.campaigns || []));
            });
            
            const weekTotals = getPrecomputedTotals(allWeekCampaigns, `applovin_test_week_${campaignKey}_${weekKey}`);
            const campaignWoWKey = `${campaignGroup.campaignId}_${weekKey}`;
            const weekWoW = getOptimizedWoW(campaignWoWKey, 'campaignWoW');
            
            const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
            const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
            const status = weekWoW.growthStatus || '';
            
            formatData.push({ row: tableData.length + 1, type: 'WEEK' });
            const weekRow = createUnifiedRow('WEEK', week, weekTotals, spendWoW, profitWoW, status, app.appName, initialMetricsCache);
            tableData.push(weekRow);
            
            // Добавляем страны для недели
            if (week.countries) {
              const countryKeys = Object.keys(week.countries).sort((a, b) => {
                const getCountrySpend = (country) => (country.campaigns || []).reduce((sum, c) => sum + c.spend, 0);
                return getCountrySpend(week.countries[b]) - getCountrySpend(week.countries[a]);
              });
              
              countryKeys.forEach(countryCode => {
                const country = week.countries[countryCode];
                const countryTotals = getPrecomputedTotals(country.campaigns || [], `applovin_test_country_${campaignKey}_${weekKey}_${countryCode}`);
                const countryWoWKey = `${campaignGroup.campaignId}_${weekKey}_${countryCode}`;
                const countryWoW = getOptimizedWoW(countryWoWKey, 'countryWoW');
                
                const countrySpendWoW = countryWoW.spendChangePercent !== undefined ? `${countryWoW.spendChangePercent.toFixed(0)}%` : '';
                const countryProfitWoW = countryWoW.eProfitChangePercent !== undefined ? `${countryWoW.eProfitChangePercent.toFixed(0)}%` : '';
                const countryStatus = countryWoW.growthStatus || '';
                
                formatData.push({ row: tableData.length + 1, type: 'COUNTRY' });
                const countryRow = createUnifiedRow('COUNTRY', week, countryTotals, countrySpendWoW, countryProfitWoW, countryStatus, app.appName, initialMetricsCache, countryCode, country.countryName);
                tableData.push(countryRow);
              });
            }
          });
        });
      } else {
        // Стандартная обработка для других проектов
        const weekKeys = Object.keys(app.weeks).sort();
        
        weekKeys.forEach((weekKey, weekIndex) => {
          const week = app.weeks[weekKey];
          
          formatData.push({ row: tableData.length + 1, type: 'WEEK' });
          
          let allCampaigns = [];
          if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
            Object.values(week.sourceApps).forEach(sourceApp => {
              allCampaigns.push(...sourceApp.campaigns);
            });
          } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
            Object.values(week.networks).forEach(network => {
              allCampaigns.push(...network.campaigns);
            });
          } else {
            allCampaigns = week.campaigns || [];
          }
          
          const weekTotals = getPrecomputedTotals(allCampaigns, `app_${appKey}_${weekKey}`);
          const appWeekKey = `${app.appName}_${weekKey}`;
          const weekWoW = getOptimizedWoW(appWeekKey, 'appWeekWoW');
          
          const spendWoW = weekWoW.spendChangePercent !== undefined ? `${weekWoW.spendChangePercent.toFixed(0)}%` : '';
          const profitWoW = weekWoW.eProfitChangePercent !== undefined ? `${weekWoW.eProfitChangePercent.toFixed(0)}%` : '';
          const status = weekWoW.growthStatus || '';
          
          const weekRow = createUnifiedRow('WEEK', week, weekTotals, spendWoW, profitWoW, status, app.appName, initialMetricsCache);
          tableData.push(weekRow);
          
          addOptimizedSubRows(tableData, week, weekKey, formatData, app.appName, initialMetricsCache, appsDbCache);
        });
      }
    });
  }
  
  console.log(`buildUnifiedTable completed: ${tableData.length} rows in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
}

function precomputeAllTotals(data) {
  const startTime = Date.now();
  let computedCount = 0;
  
  if (CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    Object.keys(data).forEach(networkKey => {
      const network = data[networkKey];
      Object.values(network.countries).forEach(country => {
        Object.values(country.campaigns).forEach(campaign => {
          Object.keys(campaign.weeks).forEach(weekKey => {
            const week = campaign.weeks[weekKey];
            const weekData = week.data || [];
            
            // Создаем кэш для каждой недели кампании
            const cacheKey = `incent_week_${networkKey}_${country.countryCode}_${campaign.campaignId}_${weekKey}`;
            const totals = calculateWeekTotals(weekData.map(d => ({
              ...d,
              campaignId: campaign.campaignId,
              campaignName: campaign.campaignName
            })));
            PRECOMPUTED_TOTALS.set(cacheKey, totals);
            computedCount++;
          });
        });
      });
    });
  } else {
    Object.keys(data).forEach(appKey => {
      const app = data[appKey];
      Object.keys(app.weeks).forEach(weekKey => {
        const week = app.weeks[weekKey];
        
        let allCampaigns = [];
        if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
          Object.values(week.sourceApps).forEach(sourceApp => {
            allCampaigns.push(...sourceApp.campaigns);
            
            const sourceAppCacheKey = `sourceapp_${appKey}_${weekKey}_${sourceApp.sourceAppId}`;
            const sourceAppTotals = calculateWeekTotals(sourceApp.campaigns);
            PRECOMPUTED_TOTALS.set(sourceAppCacheKey, sourceAppTotals);
            computedCount++;
            
            WOW_KEYS_CACHE.set(`sourceApp_${sourceApp.sourceAppId}_${weekKey}`, `sourceAppWoW_${sourceApp.sourceAppId}_${weekKey}`);
            
            sourceApp.campaigns.forEach(campaign => {
              const campaignKey = `campaign_${campaign.campaignId}_${weekKey}`;
              WOW_KEYS_CACHE.set(campaignKey, `campaignWoW_${campaign.campaignId}_${weekKey}`);
            });
          });
        } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
          Object.values(week.networks).forEach(network => {
            allCampaigns.push(...network.campaigns);
            
            const networkCacheKey = `overall_network_${appKey}_${weekKey}_${network.networkId}`;
            const networkTotals = calculateWeekTotals(network.campaigns);
            PRECOMPUTED_TOTALS.set(networkCacheKey, networkTotals);
            computedCount++;
            
            WOW_KEYS_CACHE.set(`network_${network.networkId}_${weekKey}`, `campaignWoW_${network.networkId}_${weekKey}`);
          });
        } else {
          allCampaigns = week.campaigns || [];
          if (allCampaigns.length > 0) {
            allCampaigns.forEach(campaign => {
              const campaignKey = `campaign_${campaign.campaignId}_${weekKey}`;
              WOW_KEYS_CACHE.set(campaignKey, `campaignWoW_${campaign.campaignId}_${weekKey}`);
            });
          }
        }
        
        const cacheKey = `app_${appKey}_${weekKey}`;
        const totals = calculateWeekTotals(allCampaigns);
        PRECOMPUTED_TOTALS.set(cacheKey, totals);
        computedCount++;
        
        WOW_KEYS_CACHE.set(`appWeek_${app.appName}_${weekKey}`, `appWeekWoW_${app.appName}_${weekKey}`);
      });
    });
  }
  
  console.log(`Precomputed ${computedCount} totals in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
}

function precomputeWoWCache(wow) {
  const startTime = Date.now();
  let cacheCount = 0;
  
  if (wow.campaignWoW) {
    Object.keys(wow.campaignWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`campaignWoW_${key}`, wow.campaignWoW[key]);
      cacheCount++;
    });
  }
  
  if (wow.appWeekWoW) {
    Object.keys(wow.appWeekWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`appWeekWoW_${key}`, wow.appWeekWoW[key]);
      cacheCount++;
    });
  }
  
  if (wow.sourceAppWoW) {
    Object.keys(wow.sourceAppWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`sourceAppWoW_${key}`, wow.sourceAppWoW[key]);
      cacheCount++;
    });
  }
  
  if (wow.weekWoW) {
    Object.keys(wow.weekWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`weekWoW_${key}`, wow.weekWoW[key]);
      cacheCount++;
    });
  }
  
  if (wow.appWoW) {
    Object.keys(wow.appWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`appWoW_${key}`, wow.appWoW[key]);
      cacheCount++;
    });
  }
  
  if (wow.networkWoW) {
    Object.keys(wow.networkWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`networkWoW_${key}`, wow.networkWoW[key]);
      cacheCount++;
    });
  }
  
  if (wow.countryWoW) {
    Object.keys(wow.countryWoW).forEach(key => {
      WOW_METRICS_CACHE.set(`countryWoW_${key}`, wow.countryWoW[key]);
      cacheCount++;
    });
  }
  
  console.log(`Precomputed ${cacheCount} WoW entries in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
}

function getOptimizedWoW(key, type) {
  const cacheKey = `${type}_${key}`;
  const cached = WOW_METRICS_CACHE.get(cacheKey);
  if (cached) {
    return cached;
  }
  
  return { spendChangePercent: 0, eProfitChangePercent: 0, growthStatus: 'First Week' };
}

function getCachedWoW(key, type, fallbackWow) {
  return getOptimizedWoW(key, type);
}

function getPrecomputedTotals(campaigns, cacheKey) {
  const cached = PRECOMPUTED_TOTALS.get(cacheKey);
  if (cached) {
    return cached;
  }
  
  return getCachedWeekTotals(campaigns);
}

function addOptimizedSubRows(tableData, week, weekKey, formatData, appName = '', initialMetricsCache = null, appsDbCache = null) {
  if (CURRENT_PROJECT === 'TRICKY' && week.sourceApps) {
    const sourceAppKeys = Object.keys(week.sourceApps).sort((a, b) => {
      const totalSpendA = week.sourceApps[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
      const totalSpendB = week.sourceApps[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
      return totalSpendB - totalSpendA;
    });
    
    sourceAppKeys.forEach(sourceAppKey => {
      const sourceApp = week.sourceApps[sourceAppKey];
      const sourceAppTotals = getPrecomputedTotals(sourceApp.campaigns, `sourceapp_${appName}_${weekKey}_${sourceApp.sourceAppId}`);
      
      const sourceAppWoW = getOptimizedWoW(`${sourceApp.sourceAppId}_${weekKey}`, 'sourceAppWoW');
      
      const spendWoW = sourceAppWoW.spendChangePercent !== undefined ? `${sourceAppWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = sourceAppWoW.eProfitChangePercent !== undefined ? `${sourceAppWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = sourceAppWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'SOURCE_APP' });
      
      let sourceAppDisplayName = sourceApp.sourceAppName;
      if (CURRENT_PROJECT === 'TRICKY' && appsDbCache) {
      const appInfo = appsDbCache[sourceApp.sourceAppId];
      if (appInfo && appInfo.linkApp) {
        sourceAppDisplayName = `=HYPERLINK("${appInfo.linkApp}", "${sourceApp.sourceAppName}")`;
        formatData.push({ row: tableData.length + 1, type: 'HYPERLINK' });
      }
}
      
      const sourceAppRow = createUnifiedRow('SOURCE_APP', week, sourceAppTotals, spendWoW, profitWoW, status, appName, initialMetricsCache, sourceApp.sourceAppId, sourceAppDisplayName);
      tableData.push(sourceAppRow);
      
      addOptimizedCampaignRows(tableData, sourceApp.campaigns, { weekStart: weekKey.split('-').join('/'), weekEnd: '' }, weekKey, formatData, appName, initialMetricsCache);
    });
  } else if (CURRENT_PROJECT === 'OVERALL' && week.networks) {
    const networkKeys = Object.keys(week.networks).sort((a, b) => {
      const totalSpendA = week.networks[a].campaigns.reduce((sum, c) => sum + c.spend, 0);
      const totalSpendB = week.networks[b].campaigns.reduce((sum, c) => sum + c.spend, 0);
      return totalSpendB - totalSpendA;
    });
    
    networkKeys.forEach(networkKey => {
      const network = week.networks[networkKey];
      const networkTotals = getPrecomputedTotals(network.campaigns, `overall_network_${appName}_${weekKey}_${network.networkId}`);
      
      const networkWoW = getOptimizedWoW(`${networkKey}_${weekKey}`, 'campaignWoW');
      
      const spendWoW = networkWoW.spendChangePercent !== undefined ? `${networkWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitWoW = networkWoW.eProfitChangePercent !== undefined ? `${networkWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const status = networkWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'NETWORK' });
      
      const networkRow = createUnifiedRow('NETWORK', week, networkTotals, spendWoW, profitWoW, status, appName, initialMetricsCache, network.networkId, network.networkName);
      tableData.push(networkRow);
    });
  } else if (CURRENT_PROJECT !== 'OVERALL' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
    addOptimizedCampaignRows(tableData, week.campaigns, week, weekKey, formatData, appName, initialMetricsCache);
  }
}

function addOptimizedCampaignRows(tableData, campaigns, week, weekKey, formatData, appName = '', initialMetricsCache = null) {
  if (CURRENT_PROJECT === 'OVERALL' || CURRENT_PROJECT === 'INCENT_TRAFFIC') {
    return;
  }
  
  const sortedCampaigns = campaigns.sort((a, b) => b.spend - a.spend);
  const batchSize = 1000;
  
  for (let i = 0; i < sortedCampaigns.length; i += batchSize) {
    const batch = sortedCampaigns.slice(i, i + batchSize);
    
    batch.forEach(campaign => {
      let campaignIdValue;
      if (CURRENT_PROJECT === 'TRICKY' || CURRENT_PROJECT === 'REGULAR') {
        campaignIdValue = `=HYPERLINK("https://app.appgrowth.com/campaigns/${campaign.campaignId}", "${campaign.campaignId}")`;
      } else {
        campaignIdValue = campaign.campaignId;
      }
      
      const campaignWoW = getOptimizedWoW(`${campaign.campaignId}_${weekKey}`, 'campaignWoW');
      
      const spendPct = campaignWoW.spendChangePercent !== undefined ? `${campaignWoW.spendChangePercent.toFixed(0)}%` : '';
      const profitPct = campaignWoW.eProfitChangePercent !== undefined ? `${campaignWoW.eProfitChangePercent.toFixed(0)}%` : '';
      const growthStatus = campaignWoW.growthStatus || '';
      
      formatData.push({ row: tableData.length + 1, type: 'CAMPAIGN' });
      
      const campaignRow = createUnifiedRow('CAMPAIGN', week, campaign, spendPct, profitPct, growthStatus, appName, initialMetricsCache, campaign.campaignId, campaign.sourceApp, campaignIdValue);
      tableData.push(campaignRow);
    });
  }
}

function addUnifiedSubRows(tableData, week, weekKey, wow, formatData, appName = '', initialMetricsCache = null, appsDbCache = null) {
  return addOptimizedSubRows(tableData, week, weekKey, formatData, appName, initialMetricsCache, appsDbCache);
}

function addCampaignRowsBatched(tableData, campaigns, week, weekKey, wow, formatData, appName = '', initialMetricsCache = null) {
  return addOptimizedCampaignRows(tableData, campaigns, week, weekKey, formatData, appName, initialMetricsCache);
}

function createUnifiedRow(level, week, data, spendWoW, profitWoW, status, appName = '', initialMetricsCache = null, identifier = '', displayName = '', campaignIdValue = '') {
  const headers = getUnifiedHeaders();
  const row = new Array(headers.length).fill('');
  
  row[COLUMN_CONFIG.COLUMNS.LEVEL - 1] = level;
  
  if (level === 'APP' && CURRENT_PROJECT !== 'INCENT_TRAFFIC') {
    row[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = displayName || identifier;
    return row;
  } else if (level === 'WEEK') {
    row[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = `${week.weekStart} - ${week.weekEnd}`;
    const combinedRoas = `${data.avgRoasD1.toFixed(0)}% → ${data.avgRoasD3.toFixed(0)}% → ${data.avgRoasD7.toFixed(0)}% → ${data.avgRoasD14.toFixed(0)}% → ${data.avgRoasD30.toFixed(0)}%`;
    
    let eROAS730Display = `${data.avgEROASD730.toFixed(0)}%`;
    let eProfit730Display = formatSmartCurrency(data.totalProfit);
    
    if (initialMetricsCache && appName) {
      const weekRange = `${week.weekStart} - ${week.weekEnd}`;
      eROAS730Display = initialMetricsCache.formatEROASWithInitial('WEEK', appName, weekRange, data.avgEROASD730);
      eProfit730Display = initialMetricsCache.formatProfitWithInitial('WEEK', appName, weekRange, data.totalProfit);
    }
    
    row[COLUMN_CONFIG.COLUMNS.SPEND - 1] = formatSmartCurrency(data.totalSpend); row[COLUMN_CONFIG.COLUMNS.SPEND_WOW - 1] = spendWoW; row[COLUMN_CONFIG.COLUMNS.INSTALLS - 1] = data.totalInstalls; row[COLUMN_CONFIG.COLUMNS.CPI - 1] = data.avgCpi.toFixed(3);
    row[COLUMN_CONFIG.COLUMNS.ROAS_COMBINED - 1] = combinedRoas; row[COLUMN_CONFIG.COLUMNS.IPM - 1] = data.avgIpm.toFixed(1); row[COLUMN_CONFIG.COLUMNS.RR_COMBINED - 1] = `${data.avgRrD1.toFixed(0)}% → ${data.avgRrD7.toFixed(0)}%`;
    row[COLUMN_CONFIG.COLUMNS.EARPU - 1] = data.avgArpu.toFixed(3); row[COLUMN_CONFIG.COLUMNS.EROAS_365 - 1] = `${data.avgERoas.toFixed(0)}%`; row[COLUMN_CONFIG.COLUMNS.EROAS_730 - 1] = eROAS730Display;
    row[COLUMN_CONFIG.COLUMNS.EPROFIT_730 - 1] = eProfit730Display; row[COLUMN_CONFIG.COLUMNS.EPROFIT_WOW - 1] = profitWoW; row[COLUMN_CONFIG.COLUMNS.GROWTH_STATUS - 1] = status; row[COLUMN_CONFIG.COLUMNS.COMMENTS - 1] = '';
  } else if (level === 'CAMPAIGN') {
    row[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = data.sourceApp; row[COLUMN_CONFIG.COLUMNS.ID - 1] = campaignIdValue; row[COLUMN_CONFIG.COLUMNS.GEO - 1] = data.geo;
    const combinedRoas = `${data.roasD1.toFixed(0)}% → ${data.roasD3.toFixed(0)}% → ${data.roasD7.toFixed(0)}% → ${data.roasD14.toFixed(0)}% → ${data.roasD30.toFixed(0)}%`;
    
    let eROAS730Display = `${data.eRoasForecastD730.toFixed(0)}%`;
    let eProfit730Display = formatSmartCurrency(data.eProfitForecast);
    
    if (initialMetricsCache && appName) {
      const weekRange = `${week.weekStart} - ${week.weekEnd}`;
      eROAS730Display = initialMetricsCache.formatEROASWithInitial('CAMPAIGN', appName, weekRange, data.eRoasForecastD730, data.campaignId, data.sourceApp);
      eProfit730Display = initialMetricsCache.formatProfitWithInitial('CAMPAIGN', appName, weekRange, data.eProfitForecast, data.campaignId, data.sourceApp);
    }
    
    row[COLUMN_CONFIG.COLUMNS.SPEND - 1] = formatSmartCurrency(data.spend); row[COLUMN_CONFIG.COLUMNS.SPEND_WOW - 1] = spendWoW; row[COLUMN_CONFIG.COLUMNS.INSTALLS - 1] = data.installs; row[COLUMN_CONFIG.COLUMNS.CPI - 1] = data.cpi ? data.cpi.toFixed(3) : '0.000';
    row[COLUMN_CONFIG.COLUMNS.ROAS_COMBINED - 1] = combinedRoas; row[COLUMN_CONFIG.COLUMNS.IPM - 1] = data.ipm.toFixed(1); row[COLUMN_CONFIG.COLUMNS.RR_COMBINED - 1] = `${data.rrD1.toFixed(0)}% → ${data.rrD7.toFixed(0)}%`;
    row[COLUMN_CONFIG.COLUMNS.EARPU - 1] = data.eArpuForecast.toFixed(3); row[COLUMN_CONFIG.COLUMNS.EROAS_365 - 1] = `${data.eRoasForecast.toFixed(0)}%`; row[COLUMN_CONFIG.COLUMNS.EROAS_730 - 1] = eROAS730Display;
    row[COLUMN_CONFIG.COLUMNS.EPROFIT_730 - 1] = eProfit730Display; row[COLUMN_CONFIG.COLUMNS.EPROFIT_WOW - 1] = profitWoW; row[COLUMN_CONFIG.COLUMNS.GROWTH_STATUS - 1] = status; row[COLUMN_CONFIG.COLUMNS.COMMENTS - 1] = '';
  } else {
    row[COLUMN_CONFIG.COLUMNS.WEEK_RANGE - 1] = displayName || identifier;
    const combinedRoas = `${data.avgRoasD1.toFixed(0)}% → ${data.avgRoasD3.toFixed(0)}% → ${data.avgRoasD7.toFixed(0)}% → ${data.avgRoasD14.toFixed(0)}% → ${data.avgRoasD30.toFixed(0)}%`;
    
    let eROAS730Display = `${data.avgEROASD730.toFixed(0)}%`;
    let eProfit730Display = formatSmartCurrency(data.totalProfit);
    
    if (initialMetricsCache && appName) {
      const weekRange = `${week.weekStart} - ${week.weekEnd}`;
      eROAS730Display = initialMetricsCache.formatEROASWithInitial(level, appName, weekRange, data.avgEROASD730, identifier, displayName);
      eProfit730Display = initialMetricsCache.formatProfitWithInitial(level, appName, weekRange, data.totalProfit, identifier, displayName);
    }
    
    row[COLUMN_CONFIG.COLUMNS.SPEND - 1] = formatSmartCurrency(data.totalSpend); row[COLUMN_CONFIG.COLUMNS.SPEND_WOW - 1] = spendWoW; row[COLUMN_CONFIG.COLUMNS.INSTALLS - 1] = data.totalInstalls; row[COLUMN_CONFIG.COLUMNS.CPI - 1] = data.avgCpi.toFixed(3);
    row[COLUMN_CONFIG.COLUMNS.ROAS_COMBINED - 1] = combinedRoas; row[COLUMN_CONFIG.COLUMNS.IPM - 1] = data.avgIpm.toFixed(1); row[COLUMN_CONFIG.COLUMNS.RR_COMBINED - 1] = `${data.avgRrD1.toFixed(0)}% → ${data.avgRrD7.toFixed(0)}%`;
    row[COLUMN_CONFIG.COLUMNS.EARPU - 1] = data.avgArpu.toFixed(3); row[COLUMN_CONFIG.COLUMNS.EROAS_365 - 1] = `${data.avgERoas.toFixed(0)}%`; row[COLUMN_CONFIG.COLUMNS.EROAS_730 - 1] = eROAS730Display;
    row[COLUMN_CONFIG.COLUMNS.EPROFIT_730 - 1] = eProfit730Display; row[COLUMN_CONFIG.COLUMNS.EPROFIT_WOW - 1] = profitWoW; row[COLUMN_CONFIG.COLUMNS.GROWTH_STATUS - 1] = status; row[COLUMN_CONFIG.COLUMNS.COMMENTS - 1] = '';
  }
  
  return row;
}

function getCachedWeekTotals(campaigns) {
  if (!campaigns || campaigns.length === 0) {
    return {
      totalSpend: 0, totalInstalls: 0, avgCpi: 0, avgRoasD1: 0, avgRoasD3: 0, avgRoasD7: 0, avgRoasD14: 0, avgRoasD30: 0,
      avgIpm: 0, avgRrD1: 0, avgRrD7: 0, avgArpu: 0, avgERoas: 0, avgEROASD730: 0, totalProfit: 0
    };
  }
  
  const cacheKey = campaigns.map(c => `${c.campaignId}_${c.spend}_${c.installs}`).join('|');
  
  if (WEEK_TOTALS_CACHE.has(cacheKey)) {
    return WEEK_TOTALS_CACHE.get(cacheKey);
  }
  
  const result = calculateWeekTotals(campaigns);
  WEEK_TOTALS_CACHE.set(cacheKey, result);
  return result;
}

function calculateWeekTotals(campaigns) {
  const totalSpend = campaigns.reduce((s, c) => s + c.spend, 0);
  const totalInstalls = campaigns.reduce((s, c) => s + c.installs, 0);
  const avgCpi = totalInstalls ? totalSpend / totalInstalls : 0;
  
  const avgRoasD1 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD1, 0) / campaigns.length : 0;
  const avgRoasD3 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD3, 0) / campaigns.length : 0;
  const avgRoasD7 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD7, 0) / campaigns.length : 0;
  const avgRoasD14 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD14, 0) / campaigns.length : 0;
  const avgRoasD30 = campaigns.length ? campaigns.reduce((s, c) => s + c.roasD30, 0) / campaigns.length : 0;
  
  const avgIpm = campaigns.length ? campaigns.reduce((s, c) => s + c.ipm, 0) / campaigns.length : 0;
  const avgRrD1 = campaigns.length ? campaigns.reduce((s, c) => s + c.rrD1, 0) / campaigns.length : 0;
  const avgRrD7 = campaigns.length ? campaigns.reduce((s, c) => s + c.rrD7, 0) / campaigns.length : 0;
  const avgArpu = campaigns.length ? campaigns.reduce((s, c) => s + c.eArpuForecast, 0) / campaigns.length : 0;
  
  const validForEROAS = campaigns.filter(c => 
    c.eRoasForecast >= 1 && 
    c.eRoasForecast <= 1000 && 
    c.spend > 0
  );
  
  let avgERoas = 0;
  if (validForEROAS.length > 0) {
    const totalWeightedEROAS = validForEROAS.reduce((sum, c) => sum + (c.eRoasForecast * c.spend), 0);
    const totalSpendForEROAS = validForEROAS.reduce((sum, c) => sum + c.spend, 0);
    avgERoas = totalSpendForEROAS > 0 ? totalWeightedEROAS / totalSpendForEROAS : 0;
  }
  
  const validForEROASD730 = campaigns.filter(c => 
    c.eRoasForecastD730 >= 1 && 
    c.eRoasForecastD730 <= 1000 && 
    c.spend > 0
  );
  
  let avgEROASD730 = 0;
  if (validForEROASD730.length > 0) {
    const totalWeightedEROASD730 = validForEROASD730.reduce((sum, c) => sum + (c.eRoasForecastD730 * c.spend), 0);
    const totalSpendForEROASD730 = validForEROASD730.reduce((sum, c) => sum + c.spend, 0);
    avgEROASD730 = totalSpendForEROASD730 > 0 ? totalWeightedEROASD730 / totalSpendForEROASD730 : 0;
  }
  
  const totalProfit = campaigns.reduce((s, c) => s + c.eProfitForecast, 0);

  return {
    totalSpend, totalInstalls, avgCpi, avgRoasD1, avgRoasD3, avgRoasD7, avgRoasD14, avgRoasD30, avgIpm, avgRrD1, avgRrD7,
    avgArpu, avgERoas, avgEROASD730, totalProfit
  };
}

function clearTableBuilderCaches() {
  WEEK_TOTALS_CACHE.clear();
  WOW_METRICS_CACHE.clear();
  PRECOMPUTED_TOTALS.clear();
  PRECOMPUTED_ROWS.clear();
  WOW_KEYS_CACHE.clear();
}

function getUnifiedHeaders() {
  return TABLE_CONFIG.HEADERS;
}