
function debugSingleProject() {
  const project = showChoice('Select Project:', MENU_PROJECTS);
  if (!project) return;
  
  try {
    setCurrentProject(MENU_PROJECTS[project-1].toUpperCase());
    const dateRange = getDateRange(7);
    const raw = fetchCampaignData(dateRange);
    const count = raw.data?.analytics?.richStats?.stats?.length || 0;
    
    SpreadsheetApp.getUi().alert(
      `✅ ${CURRENT_PROJECT} API works: ${count} records`
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${e.toString()}`);
  }
}

function debugIncentTraffic() {
  console.log('=== DEBUG INCENT_TRAFFIC START ===');
  
  try {
    // Шаг 1: Установка проекта
    console.log('Step 1: Setting project to INCENT_TRAFFIC');
    setCurrentProject('INCENT_TRAFFIC');
    console.log('Current project:', CURRENT_PROJECT);
    
    const config = getCurrentConfig();
    const apiConfig = getCurrentApiConfig();
    console.log('Config loaded. Sheet name:', config.SHEET_NAME);
    console.log('API networks:', apiConfig.FILTERS?.ATTRIBUTION_NETWORK_HID?.length || 0);
    console.log('Group by dimensions:', apiConfig.GROUP_BY?.length || 0);
    
    // Шаг 2: Получение данных за 6 недель
    console.log('Step 2: Fetching 6 weeks of data...');
    const dateRange = getDateRange(42); // 6 недель = 42 дня
    console.log('Date range:', dateRange);
    console.log('Date dimension:', apiConfig.DATE_DIMENSION);
    
    const raw = fetchCampaignData(dateRange);
    const statsCount = raw.data?.analytics?.richStats?.stats?.length || 0;
    console.log('Raw data received, stats count:', statsCount);
    
    if (statsCount === 0) {
      console.warn('WARNING: No data received from API');
      return;
    }
    
    // Анализируем первые несколько записей
    console.log('Step 3: Analyzing raw data structure...');
    const stats = raw.data.analytics.richStats.stats;
    const sampleSize = Math.min(3, stats.length);
    
    for (let i = 0; i < sampleSize; i++) {
      const row = stats[i];
      console.log(`Sample row ${i + 1}:`);
      console.log(`  [0] Date: ${row[0]?.value} (${row[0]?.__typename})`);
      console.log(`  [1] Country: ${row[1]?.value} (${row[1]?.code}) (${row[1]?.__typename})`);
      console.log(`  [2] Network: ${row[2]?.value} (ID: ${row[2]?.id}) (${row[2]?.__typename})`);
      console.log(`  [3] Campaign: ${row[3]?.value} (ID: ${row[3]?.id}) (${row[3]?.__typename})`);
      console.log(`  [4] App: ${row[4]?.name} (ID: ${row[4]?.id}) (${row[4]?.__typename})`);
      console.log(`  [5+] Metrics: ${row.length - 5} values`);
    }
    
    // Шаг 4: Обработка данных
    console.log('Step 4: Processing data with new structure...');
    const processed = processApiData(raw);
    console.log('Data processed successfully');
    
    // Анализируем структуру обработанных данных
    console.log('Step 5: Analyzing processed data structure...');
    const networkKeys = Object.keys(processed);
    console.log('Networks count:', networkKeys.length);
    console.log('Network keys:', networkKeys);
    
    if (networkKeys.length > 0) {
      const firstNetworkKey = networkKeys[0];
      const firstNetwork = processed[firstNetworkKey];
      console.log('First network:', firstNetwork.networkName);
      console.log('Countries in first network:', Object.keys(firstNetwork.countries).length);
      
      const countryKeys = Object.keys(firstNetwork.countries);
      if (countryKeys.length > 0) {
        const firstCountryKey = countryKeys[0];
        const firstCountry = firstNetwork.countries[firstCountryKey];
        console.log('First country:', firstCountry.countryName, `(${firstCountry.countryCode})`);
        console.log('Campaigns in first country:', Object.keys(firstCountry.campaigns).length);
        
        const campaignKeys = Object.keys(firstCountry.campaigns);
        if (campaignKeys.length > 0) {
          const firstCampaignKey = campaignKeys[0];
          const firstCampaign = firstCountry.campaigns[firstCampaignKey];
          console.log('First campaign:', firstCampaign.campaignName);
          console.log('Weeks in first campaign:', Object.keys(firstCampaign.weeks).length);
          
          const weekKeys = Object.keys(firstCampaign.weeks);
          if (weekKeys.length > 0) {
            const firstWeekKey = weekKeys[0];
            const firstWeek = firstCampaign.weeks[firstWeekKey];
            console.log('First week:', `${firstWeek.weekStart} - ${firstWeek.weekEnd}`);
            console.log('Data points in first week:', firstWeek.data?.length || 0);
            
            if (firstWeek.data && firstWeek.data.length > 0) {
              const firstDataPoint = firstWeek.data[0];
              console.log('First data point metrics:', {
                spend: firstDataPoint.spend,
                installs: firstDataPoint.installs,
                eProfitForecast: firstDataPoint.eProfitForecast,
                appName: firstDataPoint.appName
              });
            }
          }
        }
      }
    }
    
    // Шаг 6: Расчет WoW метрик
    console.log('Step 6: Calculating WoW metrics...');
    const wow = calculateIncentTrafficWoWMetrics(processed);
    const wowKeys = Object.keys(wow.weekWoW);
    console.log('WoW metrics calculated:', wowKeys.length, 'entries');
    
    if (wowKeys.length > 0) {
      const sampleWoWKey = wowKeys[0];
      const sampleWoW = wow.weekWoW[sampleWoWKey];
      console.log('Sample WoW entry:', sampleWoWKey);
      console.log('Sample WoW data:', {
        spendChangePercent: sampleWoW.spendChangePercent,
        eProfitChangePercent: sampleWoW.eProfitChangePercent,
        growthStatus: sampleWoW.growthStatus
      });
    }
    
    // Шаг 7: Очистка листа
    console.log('Step 7: Clearing sheet...');
    clearAllDataSilent();
    console.log('Sheet cleared');
    
    // Шаг 8: Создание таблицы
    console.log('Step 8: Creating INCENT_TRAFFIC table...');
    try {
      createIncentTrafficPivotTable(processed);
      console.log('✅ INCENT_TRAFFIC table created successfully!');
    } catch (tableError) {
      console.error('❌ ERROR in createIncentTrafficPivotTable:', tableError);
      console.error('Table error stack trace:', tableError.stack);
      throw tableError;
    }
    
    console.log('=== DEBUG INCENT_TRAFFIC COMPLETE ===');
    
  } catch (e) {
    console.error('=== DEBUG INCENT_TRAFFIC ERROR ===');
    console.error('Error message:', e.toString());
    console.error('Error stack trace:', e.stack);
    
    // Дополнительная диагностика
    try {
      console.log('Additional diagnostics:');
      console.log('Current project:', CURRENT_PROJECT);
      const config = getCurrentConfig();
      console.log('Config sheet name:', config?.SHEET_NAME);
      console.log('Bearer token present:', !!config?.BEARER_TOKEN);
    } catch (diagError) {
      console.error('Failed to get diagnostics:', diagError.toString());
    }
    
    throw e; // Re-throw для отображения в UI
  }
}