

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

function debugRunAllProjectsCaching() {
  try {
    console.log('=== ЗАПУСК КОМПЛЕКСНОГО ТЕСТА КЕШИРОВАНИЯ ===');
    const result = debugTestAllProjectsCaching();
    console.log(`=== ТЕСТ ЗАВЕРШЕН: ${result} ===`);
  } catch (e) {
    console.error('=== ОШИБКА В ТЕСТЕ ===', e);
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

/**
 * Debug функция для очистки плохих записей из кеш-таблиц И основных таблиц
 * Удаляет записи с Growth Status эмодзи и "First Week" в колонке Comments
 */
function debugCleanBadCacheEntries() {
  try {
    console.log('=== DEBUG: Cleaning bad cache entries START ===');
    
    const projects = ['INCENT_TRAFFIC', 'TRICKY', 'APPLOVIN_TEST', 'OVERALL', 'REGULAR', 'MOLOCO', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL'];
    let totalCleaned = 0;
    
    projects.forEach(projectName => {
      console.log(`Checking project: ${projectName}`);
      let cacheCleaned = 0;
      let mainCleaned = 0;
      
      try {
        // Устанавливаем текущий проект для правильной конфигурации
        setCurrentProject(projectName);
        const cache = new CommentCache(projectName);
        const config = getCurrentConfig();
        
        // 1. ОЧИЩАЕМ КЕШИ-ТАБЛИЦЫ
        console.log(`${projectName}: Cleaning cache table...`);
        cache.getOrCreateCacheSheet();
        
        const cacheResponse = cache.getSheetData(cache.cacheSpreadsheetId, `${cache.cacheSheetName}!A:I`);
        
        if (cacheResponse.values && cacheResponse.values.length > 1) {
          const badCacheRows = [];
          cacheResponse.values.slice(1).forEach((row, i) => {
            if (row.length >= 7 && row[6]) {
              const comment = row[6];
              // Проверяем на Growth Status эмодзи или "First Week"
              if (comment && typeof comment === 'string' && 
                  (comment.includes('🟢') || comment.includes('🔴') || comment.includes('🟠') || 
                   comment.includes('🔵') || comment.includes('🟡') || comment.includes('⚪') ||
                   comment.toLowerCase().includes('first week'))) {
                badCacheRows.push({
                  rowNum: i + 2, // +2 так как slice(1) и нумерация с 1
                  comment: comment,
                  key: `${row[0]}|||${row[1]}|||${row[2]}`
                });
              }
            }
          });
          
          if (badCacheRows.length > 0) {
            console.log(`${projectName} CACHE: Found ${badCacheRows.length} bad entries`);
            badCacheRows.forEach(entry => {
              console.log(`  Cache Row ${entry.rowNum}: "${entry.comment}" for ${entry.key}`);
            });
            
            // Удаляем строки из кеша BATCH запросом
            badCacheRows.sort((a, b) => b.rowNum - a.rowNum); // Сортируем по убыванию
            const deleteCacheRequests = badCacheRows.map(entry => ({
              deleteDimension: {
                range: {
                  sheetId: cache.cacheSheetId,
                  dimension: 'ROWS',
                  startIndex: entry.rowNum - 1,
                  endIndex: entry.rowNum
                }
              }
            }));
            
            Sheets.Spreadsheets.batchUpdate({
              requests: deleteCacheRequests
            }, cache.cacheSpreadsheetId);
            
            cacheCleaned = badCacheRows.length;
            console.log(`${projectName} CACHE: Cleaned ${cacheCleaned} bad entries`);
          }
        }
        
        // 2. ОЧИЩАЕМ ОСНОВНУЮ ТАБЛИЦУ
        console.log(`${projectName}: Cleaning main table Comments column...`);
        
        try {
          const mainData = cache.getSheetData(config.SHEET_ID, `${config.SHEET_NAME}!A:Z`);
          if (mainData.values && mainData.values.length > 1) {
            const cols = cache.getColumns();
            const updates = [];
            
            mainData.values.slice(1).forEach((row, i) => {
              if (row.length >= cols.comment && row[cols.comment - 1]) {
                const comment = row[cols.comment - 1];
                // Проверяем на Growth Status эмодзи или "First Week"
                if (comment && typeof comment === 'string' && 
                    (comment.includes('🟢') || comment.includes('🔴') || comment.includes('🟠') || 
                     comment.includes('🔵') || comment.includes('🟡') || comment.includes('⚪') ||
                     comment.toLowerCase().includes('first week'))) {
                  
                  const rowNum = i + 2; // +2 так как slice(1) и нумерация с 1
                  console.log(`  Main Row ${rowNum}: "${comment}" in col ${cols.comment}`);
                  
                  updates.push({
                    range: `${config.SHEET_NAME}!${String.fromCharCode(64 + cols.comment)}${rowNum}`,
                    values: [['']] // Очищаем ячейку
                  });
                }
              }
            });
            
            if (updates.length > 0) {
              Sheets.Spreadsheets.Values.batchUpdate({
                valueInputOption: 'RAW',
                data: updates
              }, config.SHEET_ID);
              
              mainCleaned = updates.length;
              console.log(`${projectName} MAIN: Cleaned ${mainCleaned} bad comments`);
            }
          }
        } catch (mainError) {
          console.error(`${projectName} MAIN: Error cleaning main table:`, mainError);
        }
        
        const projectTotal = cacheCleaned + mainCleaned;
        if (projectTotal > 0) {
          console.log(`${projectName}: TOTAL cleaned ${projectTotal} entries (cache: ${cacheCleaned}, main: ${mainCleaned})`);
        } else {
          console.log(`${projectName}: No bad entries found`);
        }
        totalCleaned += projectTotal;
        
      } catch (e) {
        console.error(`${projectName}: Error cleaning cache:`, e);
      }
      
      // Задержка для избежания quota exceeded
      if (cacheCleaned + mainCleaned > 0) {
        Utilities.sleep(2000); // 2 секунды между проектами если что-то чистили
      }
    });
    
    console.log(`=== DEBUG: Cleaning complete. Total cleaned: ${totalCleaned} entries ===`);
    return `✅ Cleaning complete. Removed ${totalCleaned} bad cache entries.`;
    
  } catch (e) {
    console.error('=== DEBUG: Error in debugCleanBadCacheEntries ===', e);
    return `❌ Error: ${e.toString()}`;
  }
}

/**
 * Debug функция для тестирования кеширования и восстановления комментариев
 * Проверяет правильность индексов колонок Comments vs Growth Status
 */
function debugTestCommentCaching() {
  try {
    console.log('=== DEBUG: Testing comment caching START ===');
    
    // Используем текущий проект
    const projectName = CURRENT_PROJECT;
    console.log(`Testing project: ${projectName}`);
    
    const cache = new CommentCache(projectName);
    const cols = cache.getColumns();
    
    console.log('Column mapping:');
    console.log(`  Comments column: ${cols.comment}`);
    console.log(`  Level column: ${cols.level}`);
    console.log(`  Name column: ${cols.name}`);
    console.log(`  ID column: ${cols.id}`);
    
    // Проверяем заголовки текущего листа
    const config = getCurrentConfig();
    const headers = cache.getSheetData(config.SHEET_ID, `${config.SHEET_NAME}!1:1`).values?.[0] || [];
    
    console.log('Sheet headers analysis:');
    headers.forEach((header, i) => {
      const colNum = i + 1;
      if (colNum === 17) console.log(`  Col ${colNum}: "${header}" <- Growth Status`);
      else if (colNum === 18) console.log(`  Col ${colNum}: "${header}" <- Comments`);
      else if (colNum >= 15) console.log(`  Col ${colNum}: "${header}"`);
    });
    
    // Тестируем сохранение тестового комментария
    const testComment = `TEST COMMENT ${new Date().toISOString()}`;
    const testKey = {
      appName: 'DEBUG_TEST',
      weekRange: '2024-01-01 - 2024-01-07',
      level: 'WEEK',
      comment: testComment,
      identifier: 'TEST_ID',
      sourceApp: 'TEST_APP',
      campaign: 'TEST_CAMPAIGN'
    };
    
    console.log('Saving test comment:', testComment);
    cache.batchSaveComments([testKey]);
    
    // Проверяем загрузку
    console.log('Loading all comments...');
    const loadedComments = cache.loadAllComments();
    const testCommentKey = cache.getCommentKey(
      testKey.appName, testKey.weekRange, testKey.level, 
      testKey.identifier, testKey.sourceApp, testKey.campaign
    );
    
    const retrievedComment = loadedComments[testCommentKey];
    console.log(`Retrieved comment: "${retrievedComment}"`);
    
    if (retrievedComment === testComment) {
      console.log('✅ Comment caching test PASSED');
      return '✅ Comment caching test PASSED';
    } else {
      console.log('❌ Comment caching test FAILED');
      console.log(`Expected: "${testComment}"`);
      console.log(`Got: "${retrievedComment}"`);
      return `❌ Comment caching test FAILED\nExpected: ${testComment}\nGot: ${retrievedComment}`;
    }
    
    // Показываем статистику кеша
    const totalComments = Object.keys(loadedComments).length;
    console.log(`Total comments in cache: ${totalComments}`);
    
    // Проверяем на подозрительные комментарии с эмодзи
    const suspiciousComments = Object.entries(loadedComments).filter(([key, comment]) => {
      return comment && typeof comment === 'string' && 
             (comment.includes('🟢') || comment.includes('🔴') || comment.includes('🟠') || 
              comment.includes('🔵') || comment.includes('🟡') || comment.includes('⚪'));
    });
    
    if (suspiciousComments.length > 0) {
      console.log(`⚠️ Found ${suspiciousComments.length} suspicious comments with Growth Status emojis:`);
      suspiciousComments.forEach(([key, comment]) => {
        console.log(`  ${key} -> "${comment}"`);
      });
    }
    
    console.log('=== DEBUG: Testing comment caching COMPLETE ===');
    return 'Debug test completed - see console for details';
    
  } catch (e) {
    console.error('=== DEBUG: Error in debugTestCommentCaching ===', e);
    return `❌ Error: ${e.toString()}`;
  }
}

/**
 * Комплексный тест кеширования для всех проектов и структур данных
 * Проверяет правильность колонок и кеширования для каждого типа проекта
 */
function debugTestAllProjectsCaching() {
  try {
    console.log('=== DEBUG: Testing ALL projects caching START ===');
    
    const projects = [
      { name: 'INCENT_TRAFFIC', structure: 'NETWORK → COUNTRY → CAMPAIGN → WEEK' },
      { name: 'APPLOVIN_TEST', structure: 'APP → CAMPAIGN → WEEK → COUNTRY' },
      { name: 'TRICKY', structure: 'APP → WEEK → SOURCE_APP → CAMPAIGN' },
      { name: 'OVERALL', structure: 'APP → WEEK → NETWORK' },
      { name: 'REGULAR', structure: 'APP → WEEK → CAMPAIGN' },
      { name: 'MOLOCO', structure: 'APP → WEEK → CAMPAIGN' },
      { name: 'GOOGLE_ADS', structure: 'APP → WEEK → CAMPAIGN' },
      { name: 'APPLOVIN', structure: 'APP → WEEK → CAMPAIGN' },
      { name: 'MINTEGRAL', structure: 'APP → WEEK → CAMPAIGN' }
    ];
    
    const results = [];
    let totalPassed = 0;
    let totalFailed = 0;
    
    projects.forEach(project => {
      console.log(`\n--- Testing ${project.name} (${project.structure}) ---`);
      
      try {
        // Переключаемся на проект
        setCurrentProject(project.name);
        const cache = new CommentCache(project.name);
        const config = getCurrentConfig();
        
        // Проверяем структуру колонок
        const cols = cache.getColumns();
        console.log(`${project.name} Column mapping:`);
        console.log(`  Comments: ${cols.comment}, Level: ${cols.level}, Name: ${cols.name}, ID: ${cols.id}`);
        
        // Проверяем заголовки листа
        let headers = [];
        try {
          headers = cache.getSheetData(config.SHEET_ID, `${config.SHEET_NAME}!1:1`).values?.[0] || [];
        } catch (e) {
          console.warn(`${project.name}: Could not read sheet headers: ${e.message}`);
          headers = ['N/A']; // Dummy для продолжения теста
        }
        
        const commentsHeader = headers[cols.comment - 1] || 'MISSING';
        const growthHeader = headers[16] || 'MISSING'; // Growth Status всегда колонка 17 (индекс 16)
        
        console.log(`${project.name} Headers:`);
        console.log(`  Col ${cols.comment}: "${commentsHeader}" <- Should be Comments`);
        console.log(`  Col 17: "${growthHeader}" <- Should be Growth Status`);
        
        // Тестируем кеширование с разными структурами
        const testData = generateTestDataForProject(project.name);
        
        console.log(`${project.name}: Testing comment save/load with structure: ${project.structure}`);
        
        let testsPassed = 0;
        let testsFailed = 0;
        
        testData.forEach((test, i) => {
          try {
            const testComment = `TEST_${project.name}_${i}_${new Date().getTime()}`;
            const testKey = { ...test, comment: testComment };
            
            // Сохраняем
            cache.batchSaveComments([testKey]);
            
            // Загружаем обратно
            const loadedComments = cache.loadAllComments();
            const generatedKey = cache.getCommentKey(
              testKey.appName, testKey.weekRange, testKey.level,
              testKey.identifier, testKey.sourceApp, testKey.campaign, testKey.country
            );
            
            const retrievedComment = loadedComments[generatedKey];
            
            if (retrievedComment === testComment) {
              console.log(`  ✅ Test ${i + 1}: ${test.level} - PASSED`);
              testsPassed++;
            } else {
              console.log(`  ❌ Test ${i + 1}: ${test.level} - FAILED`);
              console.log(`    Expected: "${testComment}"`);
              console.log(`    Got: "${retrievedComment}"`);
              console.log(`    Key: ${generatedKey}`);
              testsFailed++;
            }
          } catch (testError) {
            console.error(`  ❌ Test ${i + 1}: ${test.level} - ERROR: ${testError.message}`);
            testsFailed++;
          }
        });
        
        const projectResult = {
          project: project.name,
          structure: project.structure,
          commentsCol: cols.comment,
          commentsHeader: commentsHeader,
          growthHeader: growthHeader,
          testsPassed: testsPassed,
          testsFailed: testsFailed,
          status: testsFailed === 0 ? 'PASSED' : 'FAILED'
        };
        
        results.push(projectResult);
        
        if (testsFailed === 0) {
          totalPassed++;
          console.log(`${project.name}: ✅ ALL TESTS PASSED (${testsPassed}/${testsPassed + testsFailed})`);
        } else {
          totalFailed++;
          console.log(`${project.name}: ❌ SOME TESTS FAILED (${testsPassed}/${testsPassed + testsFailed})`);
        }
        
      } catch (projectError) {
        console.error(`${project.name}: ❌ PROJECT ERROR: ${projectError.message}`);
        results.push({
          project: project.name,
          structure: project.structure,
          status: 'ERROR',
          error: projectError.message
        });
        totalFailed++;
      }
      
      // Небольшая задержка между проектами
      Utilities.sleep(500);
    });
    
    // Итоговый отчет
    console.log('\n=== FINAL RESULTS ===');
    console.log(`Projects PASSED: ${totalPassed}/${projects.length}`);
    console.log(`Projects FAILED: ${totalFailed}/${projects.length}`);
    
    console.log('\nDetailed Results:');
    results.forEach(result => {
      const status = result.status === 'PASSED' ? '✅' : '❌';
      console.log(`${status} ${result.project}: Col=${result.commentsCol}, Header="${result.commentsHeader}", Tests=${result.testsPassed || 0}/${(result.testsPassed || 0) + (result.testsFailed || 0)}`);
      if (result.error) {
        console.log(`    Error: ${result.error}`);
      }
    });
    
    const summary = `Testing complete: ${totalPassed}/${projects.length} projects passed`;
    console.log(`\n=== DEBUG: ${summary} ===`);
    return summary;
    
  } catch (e) {
    console.error('=== DEBUG: Error in debugTestAllProjectsCaching ===', e);
    return `❌ Error: ${e.toString()}`;
  }
}

/**
 * Генерирует тестовые данные специфичные для структуры каждого проекта
 */
function generateTestDataForProject(projectName) {
  const baseData = {
    identifier: 'TEST_ID',
    sourceApp: 'TEST_APP',
    campaign: 'TEST_CAMPAIGN',
    country: 'N/A'
  };
  
  switch (projectName) {
    case 'INCENT_TRAFFIC':
      return [
        { ...baseData, appName: 'TestNetwork', weekRange: '', level: 'NETWORK' },
        { ...baseData, appName: 'TestNetwork', weekRange: '', level: 'COUNTRY', country: 'US' },
        { ...baseData, appName: 'TestNetwork', weekRange: '', level: 'CAMPAIGN', country: 'US' },
        { ...baseData, appName: 'TestNetwork', weekRange: '2024-01-01 - 2024-01-07', level: 'WEEK', country: 'US' }
      ];
      
    case 'APPLOVIN_TEST':
      return [
        { ...baseData, appName: 'TestApp', weekRange: '', level: 'CAMPAIGN' },
        { ...baseData, appName: 'TestApp', weekRange: '2024-01-01 - 2024-01-07', level: 'WEEK' },
        { ...baseData, appName: 'TestApp', weekRange: '2024-01-01 - 2024-01-07', level: 'COUNTRY', country: 'US' }
      ];
      
    case 'TRICKY':
      return [
        { ...baseData, appName: 'TestApp', weekRange: '2024-01-01 - 2024-01-07', level: 'WEEK' },
        { ...baseData, appName: 'TestApp', weekRange: '2024-01-01 - 2024-01-07', level: 'SOURCE_APP' },
        { ...baseData, appName: 'TestApp', weekRange: '2024-01-01 - 2024-01-07', level: 'CAMPAIGN' }
      ];
      
    case 'OVERALL':
      return [
        { ...baseData, appName: 'TestApp', weekRange: '2024-01-01 - 2024-01-07', level: 'WEEK' },
        { ...baseData, appName: 'TestApp', weekRange: '2024-01-01 - 2024-01-07', level: 'NETWORK' }
      ];
      
    default: // REGULAR, MOLOCO, GOOGLE_ADS, APPLOVIN, MINTEGRAL
      return [
        { ...baseData, appName: 'TestApp', weekRange: '2024-01-01 - 2024-01-07', level: 'WEEK' },
        { ...baseData, appName: 'TestApp', weekRange: '2024-01-01 - 2024-01-07', level: 'CAMPAIGN' }
      ];
  }
}