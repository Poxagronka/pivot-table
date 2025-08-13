
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

function debugApplovinTest() {
  console.log('=== DEBUG APPLOVIN_TEST START ===');
  
  try {
    // Шаг 1: Установка проекта
    console.log('Step 1: Setting project to APPLOVIN_TEST');
    setCurrentProject('APPLOVIN_TEST');
    console.log('Current project:', CURRENT_PROJECT);
    
    // Шаг 2: Получение данных
    console.log('Step 2: Fetching data...');
    const dateRange = getDateRange(24);
    console.log('Date range:', dateRange);
    
    const raw = fetchCampaignData(dateRange);
    console.log('Raw data received, stats count:', raw.data?.analytics?.richStats?.stats?.length || 0);
    
    // Шаг 3: Обработка данных
    console.log('Step 3: Processing data...');
    const processed = processApiData(raw);
    console.log('Processed data structure:');
    
    // Проверяем структуру
    const appKeys = Object.keys(processed);
    console.log('Apps count:', appKeys.length);
    
    if (appKeys.length > 0) {
      const firstApp = processed[appKeys[0]];
      console.log('First app name:', firstApp.appName);
      console.log('Has campaignGroups?', !!firstApp.campaignGroups);
      console.log('Has weeks?', !!firstApp.weeks);
      
      if (firstApp.campaignGroups) {
        const campaignCount = Object.keys(firstApp.campaignGroups).length;
        console.log('Campaign groups count:', campaignCount);
        
        if (campaignCount > 0) {
          const firstCampaignKey = Object.keys(firstApp.campaignGroups)[0];
          const firstCampaign = firstApp.campaignGroups[firstCampaignKey];
          console.log('First campaign name:', firstCampaign.campaignName);
          console.log('First campaign weeks count:', Object.keys(firstCampaign.weeks).length);
        }
      }
    }
    
    // Шаг 4: Очистка листа
    console.log('Step 4: Clearing sheet...');
    clearAllDataSilent();
    console.log('Sheet cleared');
    
    // Шаг 5: Создание таблицы
    console.log('Step 5: Creating table...');
    try {
      createEnhancedPivotTable(processed);
      console.log('Table created successfully!');
    } catch (tableError) {
      console.error('ERROR in createEnhancedPivotTable:', tableError);
      console.error('Stack trace:', tableError.stack);
      throw tableError;
    }
    
    console.log('=== DEBUG COMPLETE ===');
    
  } catch (e) {
    console.error('=== DEBUG ERROR ===');
    console.error('Error message:', e.toString());
    console.error('Stack trace:', e.stack);
  }
}