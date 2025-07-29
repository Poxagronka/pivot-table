/**
 * Debug Functions - ОБНОВЛЕНО: поддержка унифицированных метрик
 */

function debugReportGeneration() {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = getCurrentConfig();
    const projectName = CURRENT_PROJECT;
    const debugSheet = createDebugSheet(projectName);
    
    logDebug(debugSheet, `=== НАЧАЛО ДИАГНОСТИКИ ${projectName} ===`, 'HEADER');
    
    logDebug(debugSheet, '1. ПРОВЕРКА КОНФИГУРАЦИИ', 'SECTION');
    debugConfiguration(debugSheet, projectName);
    
    logDebug(debugSheet, '2. ТЕСТИРОВАНИЕ API ЗАПРОСА', 'SECTION');
    const apiResult = debugAPIRequest(debugSheet, projectName);
    
    if (!apiResult.success) {
      logDebug(debugSheet, 'ОШИБКА: API запрос не прошел. Дальнейшая диагностика невозможна.', 'ERROR');
      ui.alert('Дебаг завершен', `Проверьте лист "Debug_Log_${projectName}" для подробностей ошибки API.`, ui.ButtonSet.OK);
      return;
    }
    
    logDebug(debugSheet, '3. АНАЛИЗ СТРУКТУРЫ ДАННЫХ', 'SECTION');
    debugDataStructure(debugSheet, apiResult.data);
    
    logDebug(debugSheet, '4. ПРОВЕРКА ОБРАБОТКИ ДАННЫХ', 'SECTION');
    debugDataProcessing(debugSheet, apiResult.data);
    
    logDebug(debugSheet, '5. ПРОВЕРКА ФИЛЬТРОВ', 'SECTION');
    debugFilters(debugSheet, apiResult.data, projectName);
    
    logDebug(debugSheet, `=== ДИАГНОСТИКА ${projectName} ЗАВЕРШЕНА ===`, 'HEADER');
    ui.alert('Дебаг завершен', `Проверьте лист "Debug_Log_${projectName}" для подробного анализа проблемы.`, ui.ButtonSet.OK);
  } catch (e) {
    console.error('Ошибка в дебаг функции:', e);
    ui.alert('Ошибка дебага', 'Ошибка: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function createDebugSheet(projectName) {
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheetName = `Debug_Log_${projectName}`;
  let debugSheet = spreadsheet.getSheetByName(sheetName);
  
  if (debugSheet) {
    debugSheet.clear();
  } else {
    debugSheet = spreadsheet.insertSheet(sheetName);
  }
  
  debugSheet.getRange(1, 1, 1, 4).setValues([['Время', 'Тип', 'Сообщение', 'Детали']]);
  debugSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  debugSheet.setColumnWidth(1, 120);
  debugSheet.setColumnWidth(2, 100);
  debugSheet.setColumnWidth(3, 400);
  debugSheet.setColumnWidth(4, 500);
  
  return debugSheet;
}

function logDebug(sheet, message, type = 'INFO', details = '') {
  const timestamp = new Date().toLocaleString();
  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;
  
  sheet.getRange(newRow, 1, 1, 4).setValues([[timestamp, type, message, details]]);
  
  const colors = {
    'HEADER': { background: '#1c4587', fontColor: 'white' },
    'SECTION': { background: '#6fa8dc', fontColor: 'white' },
    'ERROR': { background: '#cc0000', fontColor: 'white' },
    'WARNING': { background: '#ff9900', fontColor: 'white' },
    'SUCCESS': { background: '#00aa00', fontColor: 'white' },
    'INFO': { background: '#f3f3f3', fontColor: 'black' }
  };
  
  if (colors[type]) {
    sheet.getRange(newRow, 1, 1, 4).setBackground(colors[type].background).setFontColor(colors[type].fontColor);
  }
  SpreadsheetApp.flush();
}

function debugConfiguration(debugSheet, projectName) {
  try {
    const config = getCurrentConfig();
    const apiConfig = getCurrentApiConfig();
    
    logDebug(debugSheet, `Проект: ${projectName}`, 'INFO');
    logDebug(debugSheet, 'Sheet ID: ' + config.SHEET_ID, 'INFO');
    logDebug(debugSheet, 'Sheet Name: ' + config.SHEET_NAME, 'INFO');
    logDebug(debugSheet, 'API URL: ' + config.API_URL, 'INFO');
    logDebug(debugSheet, 'Target eROAS D730: ' + config.TARGET_EROAS + '%', 'INFO');
    
    if (config.BEARER_TOKEN && config.BEARER_TOKEN.length > 50) {
      logDebug(debugSheet, 'Bearer Token: Найден (длина: ' + config.BEARER_TOKEN.length + ')', 'SUCCESS');
    } else {
      logDebug(debugSheet, 'Bearer Token: Отсутствует или слишком короткий!', 'ERROR');
    }
    
    logDebug(debugSheet, 'API Конфигурация:', 'INFO');
    logDebug(debugSheet, '- Users: ' + apiConfig.FILTERS.USER.length + ' элементов', 'INFO', JSON.stringify(apiConfig.FILTERS.USER));
    logDebug(debugSheet, '- Attribution Partner: ' + apiConfig.FILTERS.ATTRIBUTION_PARTNER.join(', '), 'INFO');
    
    if (apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID && apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.length > 0) {
      logDebug(debugSheet, '- Attribution Network HID: ' + apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.join(', '), 'INFO');
    } else {
      logDebug(debugSheet, '- Attribution Network HID: ALL NETWORKS (пустой массив)', 'INFO');
    }
    
    // ОБНОВЛЕНО: унифицированные метрики
    logDebug(debugSheet, 'УНИФИЦИРОВАННЫЕ МЕТРИКИ:', 'INFO');
    logDebug(debugSheet, '✅ CPI, Installs, IPM, Spend', 'SUCCESS');
    logDebug(debugSheet, '✅ RR D-1, ROAS D-1, ROAS D-3, RR D-7, ROAS D-7, ROAS D-30', 'SUCCESS');
    logDebug(debugSheet, '✅ eARPU D365, eROAS D365, eROAS D730, eProfit D730', 'SUCCESS');
    logDebug(debugSheet, '✅ Фильтрация: spend > 0 на уровне API (havingFilters)', 'SUCCESS');
    
    // GROUP_BY analysis
    logDebug(debugSheet, 'GROUP_BY структура:', 'INFO');
    apiConfig.GROUP_BY.forEach((group, index) => {
      const groupInfo = `[${index}] ${group.dimension}${group.timeBucket ? ` (${group.timeBucket})` : ''}`;
      logDebug(debugSheet, `- ${groupInfo}`, 'INFO');
    });
    
    if (projectName === 'OVERALL') {
      logDebug(debugSheet, '✅ OVERALL: Использует упрощенную группировку (app + week)', 'SUCCESS');
      logDebug(debugSheet, '✅ OVERALL: Нет разбивки по кампаниям', 'INFO');
    } else if (projectName === 'TRICKY') {
      logDebug(debugSheet, '✅ TRICKY: Использует локальную группировку через Apps Database', 'SUCCESS');
      
      try {
        const appsDb = new AppsDatabase('TRICKY');
        const cache = appsDb.loadFromCache();
        const appCount = Object.keys(cache).length;
        logDebug(debugSheet, `✅ Apps Database: ${appCount} приложений в кеше`, 'SUCCESS');
      } catch (e) {
        logDebug(debugSheet, '❌ Apps Database: Ошибка загрузки', 'ERROR', e.toString());
      }
    }
    
    if (apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH) {
      logDebug(debugSheet, '- Campaign Search: ' + apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH, 'INFO');
    } else {
      logDebug(debugSheet, '- Campaign Search: НЕТ ФИЛЬТРА (берем все кампании)', 'INFO');
    }
  } catch (e) {
    logDebug(debugSheet, 'Ошибка проверки конфигурации: ' + e.toString(), 'ERROR');
  }
}

function debugAPIRequest(debugSheet, projectName) {
  try {
    const config = getCurrentConfig();
    const apiConfig = getCurrentApiConfig();
    
    const dateRange = getDateRange(30);
    logDebug(debugSheet, 'Период запроса: ' + dateRange.from + ' до ' + dateRange.to, 'INFO');
    
    const filters = [
      { dimension: "USER", values: apiConfig.FILTERS.USER, include: true },
      { dimension: "ATTRIBUTION_PARTNER", values: apiConfig.FILTERS.ATTRIBUTION_PARTNER, include: true }
    ];
    
    if (apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID && apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.length > 0) {
      filters.push({ dimension: "ATTRIBUTION_NETWORK_HID", values: apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID, include: true });
    }
    
    if (apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH) {
      const searchPattern = apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH;
      if (searchPattern.startsWith('!')) {
        filters.push({
          dimension: "ATTRIBUTION_CAMPAIGN_HID", 
          values: [], 
          include: false,
          searchByString: searchPattern.substring(1)
        });
      } else {
        filters.push({
          dimension: "ATTRIBUTION_CAMPAIGN_HID", 
          values: [], 
          include: true, 
          searchByString: searchPattern
        });
      }
    }
    
    const dateDimension = (projectName === 'GOOGLE_ADS' || projectName === 'APPLOVIN' || projectName === 'INCENT' || projectName === 'OVERALL') ? 'DATE' : 'INSTALL_DATE';
    
    const payload = {
      operationName: apiConfig.OPERATION_NAME,
      variables: {
        dateFilters: [{
          dimension: dateDimension,
          from: dateRange.from,
          to: dateRange.to,
          include: true
        }],
        filters: filters,
        groupBy: apiConfig.GROUP_BY,
        measures: apiConfig.MEASURES,
        havingFilters: [{ measure: { id: "spend", day: null }, operator: "MORE", value: 0 }],
        anonymizationMode: "OFF",
        topFilter: null,
        revenuePredictionVersion: "",
        isMultiMediation: true
      },
      query: getGraphQLQuery()
    };
    
    logDebug(debugSheet, 'Payload создан', 'SUCCESS', 'Размер: ' + JSON.stringify(payload).length + ' символов');
    logDebug(debugSheet, '✅ УНИФИЦИРОВАННЫЕ МЕТРИКИ: ' + apiConfig.MEASURES.length + ' метрик', 'SUCCESS');
    logDebug(debugSheet, '✅ ФИЛЬТРАЦИЯ: spend > 0 включена в havingFilters', 'SUCCESS');
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Accept: 'application/json, text/plain, */*',
        'Accept-Language': 'en-US,en;q=0.9',
        Authorization: `Bearer ${config.BEARER_TOKEN}`,
        Connection: 'keep-alive',
        DNT: '1',
        Origin: 'https://app.appodeal.com',
        Referer: 'https://app.appodeal.com/analytics/reports?reloadTime=' + Date.now(),
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36',
        'x-requested-with': 'XMLHttpRequest',
        'Trace-Id': Utilities.getUuid()
      },
      payload: JSON.stringify(payload)
    };
    
    logDebug(debugSheet, 'Отправляем API запрос...', 'INFO');
    
    const response = UrlFetchApp.fetch(config.API_URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    logDebug(debugSheet, 'HTTP код ответа: ' + responseCode, responseCode === 200 ? 'SUCCESS' : 'ERROR');
    logDebug(debugSheet, 'Размер ответа: ' + responseText.length + ' символов', 'INFO');
    
    if (responseCode !== 200) {
      logDebug(debugSheet, 'Ошибка API запроса', 'ERROR', responseText.substring(0, 1000));
      return { success: false, error: 'HTTP ' + responseCode };
    }
    
    let parsedResponse;
    try {
      parsedResponse = JSON.parse(responseText);
      logDebug(debugSheet, 'JSON ответ успешно распарсен', 'SUCCESS');
    } catch (parseError) {
      logDebug(debugSheet, 'Ошибка парсинга JSON ответа', 'ERROR', parseError.toString());
      return { success: false, error: 'JSON parse error' };
    }
    
    if (parsedResponse.errors) {
      logDebug(debugSheet, 'API вернул ошибки', 'ERROR', JSON.stringify(parsedResponse.errors));
      return { success: false, error: 'API errors', data: parsedResponse };
    }
    
    logDebug(debugSheet, 'API запрос выполнен успешно', 'SUCCESS');
    return { success: true, data: parsedResponse };
  } catch (e) {
    logDebug(debugSheet, 'Исключение при API запросе: ' + e.toString(), 'ERROR');
    return { success: false, error: e.toString() };
  }
}

function debugDataStructure(debugSheet, apiResponse) {
  try {
    if (!apiResponse?.data?.analytics?.richStats?.stats) {
      logDebug(debugSheet, 'Отсутствует структура richStats!', 'ERROR');
      return;
    }
    
    const stats = apiResponse.data.analytics.richStats.stats;
    logDebug(debugSheet, 'Количество записей в stats: ' + stats.length, stats.length > 0 ? 'SUCCESS' : 'WARNING');
    
    if (stats.length === 0) {
      logDebug(debugSheet, 'Массив stats пуст - нет данных для обработки!', 'WARNING');
      return;
    }
    
    const firstRecord = stats[0];
    logDebug(debugSheet, 'Структура первой записи:', 'INFO', JSON.stringify(firstRecord, null, 2));
    
    if (Array.isArray(firstRecord)) {
      logDebug(debugSheet, 'Первая запись - массив с ' + firstRecord.length + ' элементами', 'INFO');
      
      firstRecord.forEach((item, index) => {
        if (item && typeof item === 'object') {
          const typename = item.__typename || 'неизвестный тип';
          let description = `Элемент [${index}]: ${typename}`;
          
          if (typename === 'StatsValue') {
            description += ` (value: ${item.value})`;
          } else if (typename === 'UaCampaign') {
            description += ` (campaignName: ${item.campaignName}, id: ${item.campaignId})`;
          } else if (typename === 'AppInfo') {
            description += ` (name: ${item.name}, platform: ${item.platform})`;
          } else if (typename === 'ForecastStatsItem') {
            description += ` (value: ${item.value})`;
          } else if (typename === 'RetentionStatsValue') {
            description += ` (value: ${item.value}, cohortSize: ${item.cohortSize})`;
          }
          
          logDebug(debugSheet, description, 'INFO');
        }
      });
      
      // ОБНОВЛЕНО: проверка унифицированной структуры
      logDebug(debugSheet, 'Проверка УНИФИЦИРОВАННОЙ структуры данных:', 'SECTION');
      if (CURRENT_PROJECT === 'OVERALL') {
        if (firstRecord.length >= 16) { // date + app + 14 metrics
          const hasDate = firstRecord[0] && firstRecord[0].__typename === 'StatsValue';
          const hasApp = firstRecord[1] && firstRecord[1].__typename === 'AppInfo';
          
          if (hasDate && hasApp) {
            logDebug(debugSheet, '✅ OVERALL: Корректная структура [date, app, 12 metrics] - без кампаний', 'SUCCESS');
            logDebug(debugSheet, `✅ Всего элементов: ${firstRecord.length} (ожидается: 16)`, firstRecord.length === 16 ? 'SUCCESS' : 'WARNING');
          } else {
            logDebug(debugSheet, '❌ OVERALL: Неожиданная структура данных!', 'ERROR');
            logDebug(debugSheet, `- [0] date: ${firstRecord[0]?.__typename || 'отсутствует'}`, 'INFO');
            logDebug(debugSheet, `- [1] app: ${firstRecord[1]?.__typename || 'отсутствует'}`, 'INFO');
          }
        } else {
          logDebug(debugSheet, '❌ OVERALL: Недостаточно элементов в записи!', 'ERROR');
        }
      } else {
        if (firstRecord.length >= 17) { // date + campaign + app + 14 metrics
          const hasDate = firstRecord[0] && firstRecord[0].__typename === 'StatsValue';
          const hasCampaign = firstRecord[1] && firstRecord[1].__typename === 'UaCampaign';
          const hasApp = firstRecord[2] && firstRecord[2].__typename === 'AppInfo';
          
          if (hasDate && hasCampaign && hasApp) {
            logDebug(debugSheet, '✅ Корректная структура [date, campaign, app, 12 metrics]', 'SUCCESS');
            logDebug(debugSheet, `✅ Всего элементов: ${firstRecord.length} (ожидается: 17)`, firstRecord.length === 17 ? 'SUCCESS' : 'WARNING');
            
            if (CURRENT_PROJECT === 'TRICKY') {
              logDebug(debugSheet, '✅ TRICKY: Будет группироваться локально через Apps Database', 'SUCCESS');
            }
          } else {
            logDebug(debugSheet, '❌ Неожиданная структура данных!', 'ERROR');
            logDebug(debugSheet, `- [0] date: ${firstRecord[0]?.__typename || 'отсутствует'}`, 'INFO');
            logDebug(debugSheet, `- [1] campaign: ${firstRecord[1]?.__typename || 'отсутствует'}`, 'INFO');
            logDebug(debugSheet, `- [2] app: ${firstRecord[2]?.__typename || 'отсутствует'}`, 'INFO');
          }
        } else {
          logDebug(debugSheet, '❌ Недостаточно элементов в записи!', 'ERROR');
        }
      }
    }
    
    if (apiResponse.data.analytics.richStats.totals) {
      logDebug(debugSheet, 'Поле totals найдено, записей: ' + apiResponse.data.analytics.richStats.totals.length, 'INFO');
    } else {
      logDebug(debugSheet, 'Поле totals отсутствует', 'WARNING');
    }
  } catch (e) {
    logDebug(debugSheet, 'Ошибка анализа структуры данных: ' + e.toString(), 'ERROR');
  }
}

function debugDataProcessing(debugSheet, apiResponse) {
  try {
    if (!apiResponse.data?.analytics?.richStats?.stats) {
      logDebug(debugSheet, 'Нет данных для обработки', 'ERROR');
      return;
    }
    
    const stats = apiResponse.data.analytics.richStats.stats;
    logDebug(debugSheet, 'Начинаем обработку ' + stats.length + ' записей', 'INFO');
    
    const today = new Date();
    const currentWeekStart = formatDateForAPI(getMondayOfWeek(today));
    logDebug(debugSheet, 'Текущая неделя (исключается): ' + currentWeekStart, 'INFO');
    
    const processedData = {};
    let totalProcessed = 0;
    let skippedCurrentWeek = 0;
    let errorCount = 0;
    let bundleIdCount = 0;
    
    stats.forEach((row, index) => {
      try {
        if (!Array.isArray(row)) {
          logDebug(debugSheet, `Запись ${index}: не является массивом`, 'WARNING');
          errorCount++;
          return;
        }
        
        const date = row[0]?.value;
        if (!date) {
          logDebug(debugSheet, `Запись ${index}: отсутствует дата`, 'WARNING');
          errorCount++;
          return;
        }
        
        const monday = getMondayOfWeek(new Date(date));
        const weekKey = formatDateForAPI(monday);
        
        if (weekKey >= currentWeekStart) {
          skippedCurrentWeek++;
          return;
        }
        
        let campaign, app, metricsStartIndex;
        
        if (CURRENT_PROJECT === 'OVERALL') {
          campaign = null;
          app = row[1];
          metricsStartIndex = 2;
        } else {
          campaign = row[1];
          app = row[2];
          metricsStartIndex = 3;
        }
        
        if (!app) {
          logDebug(debugSheet, `Запись ${index}: отсутствует app`, 'WARNING');
          errorCount++;
          return;
        }
        
        // ОБНОВЛЕНО: проверяем унифицированные метрики
        if (row.length < metricsStartIndex + 14) {
          logDebug(debugSheet, `Запись ${index}: недостаточно метрик (${row.length - metricsStartIndex}/14)`, 'WARNING');
          errorCount++;
          return;
        }
        
        const spendIndex = metricsStartIndex + 3; // spend всегда на 4-й позиции среди метрик
        const spendValue = parseFloat(row[spendIndex]?.value) || 0;
        
        // Spend > 0 фильтрация теперь на уровне API, но проверяем для отладки
        if (spendValue <= 0) {
          logDebug(debugSheet, `Запись ${index}: spend = ${spendValue} (должно быть отфильтровано API)`, 'WARNING');
          return;
        }
        
        const appKey = app.id;
        if (!processedData[appKey]) {
          processedData[appKey] = {
            appId: app.id,
            appName: app.name,
            platform: app.platform,
            bundleId: app.bundleId,
            weeks: {}
          };
        }
        
        if (!processedData[appKey].weeks[weekKey]) {
          processedData[appKey].weeks[weekKey] = {
            weekStart: weekKey,
            sourceApps: CURRENT_PROJECT === 'TRICKY' ? {} : null,
            campaigns: CURRENT_PROJECT === 'TRICKY' ? [] : []
          };
        }
        
        totalProcessed++;
        
        if (CURRENT_PROJECT === 'TRICKY' && campaign) {
          let campaignName = 'Unknown';
          if (campaign.campaignName) {
            campaignName = campaign.campaignName;
          } else if (campaign.value) {
            campaignName = campaign.value;
          }
          
          const bundleId = extractBundleIdFromCampaign(campaignName);
          if (bundleId) {
            bundleIdCount++;
          }
        }
        
        if (index < 3) {
          let campaignName = 'Unknown';
          if (CURRENT_PROJECT !== 'OVERALL' && campaign) {
            if (campaign.campaignName) {
              campaignName = campaign.campaignName;
            } else if (campaign.value) {
              campaignName = campaign.value;
            }
          } else if (CURRENT_PROJECT === 'OVERALL') {
            campaignName = 'N/A (app-level data)';
          }
          
          const shortCampaignName = campaignName.length > 50 ? campaignName.substring(0, 50) + '...' : campaignName;
          let recordInfo = `Запись ${index}: ${app.name}, ${shortCampaignName}, ${date}, spend=${spendValue}`;
          
          if (CURRENT_PROJECT === 'TRICKY' && campaign) {
            const bundleId = extractBundleIdFromCampaign(campaignName);
            recordInfo += `, Bundle ID: ${bundleId || 'не найден'}`;
          }
          
          logDebug(debugSheet, recordInfo, 'INFO');
        }
      } catch (e) {
        logDebug(debugSheet, `Ошибка обработки записи ${index}: ${e.toString()}`, 'ERROR');
        errorCount++;
      }
    });
    
    logDebug(debugSheet, 'Обработка завершена:', 'SUCCESS');
    logDebug(debugSheet, '- Всего записей: ' + stats.length, 'INFO');
    logDebug(debugSheet, '- Успешно обработано: ' + totalProcessed, 'INFO');
    logDebug(debugSheet, '- Пропущено (текущая неделя): ' + skippedCurrentWeek, 'INFO');
    logDebug(debugSheet, '- Ошибок обработки: ' + errorCount, errorCount > 0 ? 'WARNING' : 'INFO');
    logDebug(debugSheet, '- Уникальных приложений: ' + Object.keys(processedData).length, 'INFO');
    
    if (CURRENT_PROJECT === 'TRICKY') {
      logDebug(debugSheet, '- Bundle ID найдено: ' + bundleIdCount + ' из ' + totalProcessed, 'INFO');
    }
    
    logDebug(debugSheet, '✅ УНИФИЦИРОВАННАЯ ОБРАБОТКА: Все метрики обрабатываются одинаково', 'SUCCESS');
    logDebug(debugSheet, '✅ API ФИЛЬТРАЦИЯ: spend > 0 фильтруется на уровне API', 'SUCCESS');
    
    if (Object.keys(processedData).length === 0) {
      logDebug(debugSheet, 'ПРОБЛЕМА: После обработки не осталось данных!', 'ERROR');
      if (skippedCurrentWeek === stats.length) {
        logDebug(debugSheet, 'Все записи относятся к текущей неделе и были отфильтрованы', 'WARNING');
        logDebug(debugSheet, 'РЕШЕНИЕ: Попробуйте изменить период запроса или отключить фильтрацию текущей недели', 'INFO');
      }
    } else {
      logDebug(debugSheet, 'Данные успешно обработаны', 'SUCCESS');
      Object.values(processedData).forEach(app => {
        const weekCount = Object.keys(app.weeks).length;
        logDebug(debugSheet, `Приложение: ${app.appName} (${weekCount} недель)`, 'INFO');
      });
    }
  } catch (e) {
    logDebug(debugSheet, 'Ошибка проверки обработки данных: ' + e.toString(), 'ERROR');
  }
}

function debugFilters(debugSheet, apiResponse, projectName) {
  try {
    if (!apiResponse.data?.analytics?.richStats?.stats) {
      logDebug(debugSheet, 'Нет данных для проверки фильтров', 'ERROR');
      return;
    }
    
    const stats = apiResponse.data.analytics.richStats.stats;
    const apiConfig = getCurrentApiConfig();
    
    const uniqueApps = new Set();
    const uniqueCampaigns = new Set();
    const uniqueDates = new Set();
    const uniqueBundleIds = new Set();
    const campaignPatterns = new Set();
    let spendFilteredCount = 0;
    
    stats.forEach(row => {
      if (Array.isArray(row)) {
        const date = row[0]?.value;
        let campaign, app;
        
        if (projectName === 'OVERALL') {
          campaign = null;
          app = row[1];
        } else {
          campaign = row[1];
          app = row[2];
        }
        
        if (date) uniqueDates.add(date);
        if (app?.name) uniqueApps.add(app.name);
        
        // Проверяем фильтрацию spend > 0
        const metricsStartIndex = projectName === 'OVERALL' ? 2 : 3;
        const spendIndex = metricsStartIndex + 3;
        const spendValue = parseFloat(row[spendIndex]?.value) || 0;
        if (spendValue > 0) spendFilteredCount++;
        
        let campaignName = null;
        if (projectName !== 'OVERALL' && campaign) {
          if (campaign.campaignName) {
            campaignName = campaign.campaignName;
          } else if (campaign.value) {
            campaignName = campaign.value;
          }
        }
        
        if (campaignName) {
          uniqueCampaigns.add(campaignName);
          
          if (projectName === 'TRICKY') {
            const bundleId = extractBundleIdFromCampaign(campaignName);
            if (bundleId) {
              uniqueBundleIds.add(bundleId);
            }
          }
          
          if (projectName === 'MOLOCO' || projectName === 'MINTEGRAL') {
            if (campaignName.startsWith('APD_')) {
              campaignPatterns.add(campaignName);
            }
          } else if (projectName === 'REGULAR' || projectName === 'GOOGLE_ADS' || projectName === 'APPLOVIN') {
            campaignPatterns.add(campaignName);
          } else if (apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH) {
            const searchPattern = apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH;
            const regex = new RegExp(searchPattern.slice(1, -2), 'i');
            if (regex.test(campaignName)) {
              campaignPatterns.add(campaignName);
            }
          }
        }
      }
    });
    
    logDebug(debugSheet, `Анализ уникальных значений для ${projectName}:`, 'INFO');
    logDebug(debugSheet, '- Уникальных приложений: ' + uniqueApps.size, 'INFO');
    
    if (projectName === 'OVERALL') {
      logDebug(debugSheet, '- Кампаний: N/A (app-level data)', 'INFO');
      logDebug(debugSheet, '- Данные агрегированы на уровне приложений', 'INFO');
    } else {
      logDebug(debugSheet, '- Уникальных кампаний: ' + uniqueCampaigns.size, 'INFO');
    }
    
    logDebug(debugSheet, '- Уникальных дат: ' + uniqueDates.size, 'INFO');
    logDebug(debugSheet, '✅ SPEND ФИЛЬТРАЦИЯ: ' + spendFilteredCount + '/' + stats.length + ' записей прошли фильтр spend > 0', 'SUCCESS');
    
    if (projectName === 'TRICKY') {
      logDebug(debugSheet, '- Уникальных Bundle ID: ' + uniqueBundleIds.size, 'INFO');
      if (uniqueBundleIds.size > 0) {
        const bundleIdsList = Array.from(uniqueBundleIds).slice(0, 5);
        logDebug(debugSheet, 'Примеры Bundle ID:', 'INFO', bundleIdsList.join('\n'));
      }
    }
    
    const appsList = Array.from(uniqueApps).slice(0, 5);
    logDebug(debugSheet, 'Примеры приложений: ' + appsList.join(', '), 'INFO');
    
    if (projectName !== 'OVERALL') {
      const campaignsList = Array.from(uniqueCampaigns).slice(0, 3);
      logDebug(debugSheet, 'Примеры кампаний:', 'INFO', campaignsList.join('\n'));
    }
    
    const sortedDates = Array.from(uniqueDates).sort();
    if (sortedDates.length > 0) {
      logDebug(debugSheet, `Диапазон дат: ${sortedDates[0]} - ${sortedDates[sortedDates.length - 1]}`, 'INFO');
    }
    
    if (projectName === 'OVERALL') {
      logDebug(debugSheet, `${projectName}: Агрегированные данные по приложениям`, 'INFO');
      logDebug(debugSheet, `Network HID: ${apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.length > 0 ? apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.join(', ') : 'ALL NETWORKS'}`, 'INFO');
      logDebug(debugSheet, '✅ OVERALL данные корректны для app-level отчета!', 'SUCCESS');
    } else if (projectName === 'TRICKY') {
      logDebug(debugSheet, `${projectName}: Фильтр кампаний + локальная группировка`, 'INFO');
      logDebug(debugSheet, `Campaign Search: ${apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH}`, 'INFO');
      logDebug(debugSheet, `Network HID: ${apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.join(', ')}`, 'INFO');
      logDebug(debugSheet, '- Всего кампаний найдено: ' + campaignPatterns.size, 'INFO');
      logDebug(debugSheet, '- Bundle ID найдено: ' + uniqueBundleIds.size, 'INFO');
      
      if (campaignPatterns.size > 0 && uniqueBundleIds.size > 0) {
        const examples = Array.from(campaignPatterns).slice(0, 3);
        logDebug(debugSheet, 'Примеры кампаний:', 'INFO', examples.join('\n'));
        logDebug(debugSheet, `✅ ${projectName} данные корректны для локальной группировки!`, 'SUCCESS');
      } else {
        logDebug(debugSheet, 'ПРОБЛЕМА: Недостаточно данных для группировки!', 'ERROR');
      }
    } else if (projectName === 'MOLOCO' || projectName === 'MINTEGRAL') {
      logDebug(debugSheet, `${projectName}: Фильтр кампаний ОТКЛЮЧЕН (берем все кампании)`, 'INFO');
      logDebug(debugSheet, '- APD_ кампаний найдено: ' + campaignPatterns.size, 'INFO');
      logDebug(debugSheet, `Network HID: ${apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.join(', ')}`, 'INFO');
      
      if (campaignPatterns.size > 0) {
        const examples = Array.from(campaignPatterns).slice(0, 3);
        logDebug(debugSheet, 'Примеры APD_ кампаний:', 'INFO', examples.join('\n'));
        logDebug(debugSheet, `✅ ${projectName} данные корректны!`, 'SUCCESS');
      } else {
        logDebug(debugSheet, 'ВНИМАНИЕ: Нет APD_ кампаний в данных', 'WARNING');
      }
    } else {
      logDebug(debugSheet, `${projectName}: Фильтр кампаний работает`, 'INFO');
      logDebug(debugSheet, `Network HID: ${apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.join(', ')}`, 'INFO');
      logDebug(debugSheet, '- Всего кампаний найдено: ' + campaignPatterns.size, 'INFO');
      
      if (campaignPatterns.size > 0) {
        const examples = Array.from(campaignPatterns).slice(0, 5);
        logDebug(debugSheet, 'Примеры кампаний:', 'INFO', examples.join('\n'));
        logDebug(debugSheet, `✅ ${projectName} данные корректны!`, 'SUCCESS');
      } else {
        logDebug(debugSheet, 'ПРОБЛЕМА: Кампании не найдены!', 'ERROR');
      }
    }
  } catch (e) {
    logDebug(debugSheet, 'Ошибка проверки фильтров: ' + e.toString(), 'ERROR');
  }
}

function quickAPICheck() {
  const ui = SpreadsheetApp.getUi();
  const projectName = CURRENT_PROJECT;
  
  try {
    const dateRange = getDateRange(7);
    const raw = fetchCampaignData(dateRange);
    
    if (!raw.data?.analytics?.richStats?.stats?.length) {
      ui.alert(`${projectName} API Проверка`, `API не возвращает данные за последние 7 дней.\nЗапустите полную диагностику для детального анализа.`, ui.ButtonSet.OK);
    } else {
      const count = raw.data.analytics.richStats.stats.length;
      let message = `API работает: получено ${count} записей за последние 7 дней.`;
      
      if (projectName === 'OVERALL') {
        const firstRecord = raw.data.analytics.richStats.stats[0];
        if (firstRecord && Array.isArray(firstRecord) && firstRecord.length >= 14) {
          message += `\n✅ Структура данных корректна для app-level отчета.`;
          message += `\n✅ Унифицированные метрики: ${firstRecord.length - 2} штук.`;
        } else {
          message += `\n⚠️ Структура данных может быть некорректной.`;
        }
      } else if (projectName === 'TRICKY') {
        const firstRecord = raw.data.analytics.richStats.stats[0];
        if (firstRecord && Array.isArray(firstRecord) && firstRecord.length >= 15) {
          message += `\n✅ Структура данных корректна для локальной группировки.`;
          message += `\n✅ Унифицированные метрики: ${firstRecord.length - 3} штук.`;
        } else {
          message += `\n⚠️ Структура данных может быть некорректной.`;
        }
      } else {
        const firstRecord = raw.data.analytics.richStats.stats[0];
        if (firstRecord && Array.isArray(firstRecord) && firstRecord.length >= 15) {
          message += `\n✅ Унифицированные метрики: ${firstRecord.length - 3} штук.`;
        }
      }
      
      ui.alert(`${projectName} API Проверка`, message, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert(`${projectName} API Проверка`, 'Ошибка API: ' + e.toString(), ui.ButtonSet.OK);
  }
}

function debugCommentHashes() {
  const ui = SpreadsheetApp.getUi();
  const projectName = CURRENT_PROJECT;
  
  console.log(`\n=== Debugging comment hashes for ${projectName} ===`);
  
  try {
    const config = getCurrentConfig();
    const cache = new CommentCache(projectName);
    
    // Load cached comments
    const { comments, commentsByHash } = cache.loadAllComments();
    console.log(`Loaded ${Object.keys(commentsByHash).length} comments from cache`);
    
    // Get sheet data
    const range = `${config.SHEET_NAME}!A:T`;
    const response = Sheets.Spreadsheets.Values.get(config.SHEET_ID, range);
    
    if (!response.values || response.values.length < 2) {
      console.log('No data in sheet');
      return;
    }
    
    const data = response.values;
    const headers = data[0];
    const hashCol = headers.findIndex(h => h === 'RowHash');
    const idCol = headers.findIndex(h => h === 'ID');
    
    if (hashCol === -1) {
      console.log('RowHash column not found');
      return;
    }
    
    let currentApp = '';
    let currentWeek = '';
    let totalRows = 0;
    let rowsWithHash = 0;
    let matchingHashes = 0;
    let mismatchDetails = [];
    
    // Проходим по всем строкам и собираем хеши из кеша для анализа
    const cacheHashesByLevel = {
      WEEK: [],
      SOURCE_APP: [],
      CAMPAIGN: [],
      NETWORK: []
    };
    
    // Анализируем структуру кеша
    Object.keys(commentsByHash).forEach(hash => {
      if (hash.startsWith('TRI_W_')) cacheHashesByLevel.WEEK.push(hash);
      else if (hash.startsWith('TRI_S_')) cacheHashesByLevel.SOURCE_APP.push(hash);
      else if (hash.startsWith('TRI_C_')) cacheHashesByLevel.CAMPAIGN.push(hash);
      else if (hash.startsWith('TRI_N_')) cacheHashesByLevel.NETWORK.push(hash);
    });
    
    console.log('\nCache analysis:');
    console.log(`WEEK hashes: ${cacheHashesByLevel.WEEK.length}`);
    console.log(`SOURCE_APP hashes: ${cacheHashesByLevel.SOURCE_APP.length}`);
    console.log(`CAMPAIGN hashes: ${cacheHashesByLevel.CAMPAIGN.length}`);
    console.log(`NETWORK hashes: ${cacheHashesByLevel.NETWORK.length}`);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const level = row[0];
      const nameOrRange = row[1];
      const hash = row[hashCol];
      
      if (level === 'APP') {
        currentApp = nameOrRange;
      } else if (level === 'WEEK') {
        currentWeek = nameOrRange;
        if (hash) {
          totalRows++;
          rowsWithHash++;
          
          if (commentsByHash[hash]) {
            matchingHashes++;
          } else {
            const expectedHash = generateCommentHash('WEEK', currentApp, currentWeek, projectName);
            mismatchDetails.push({
              row: i + 1,
              level: level,
              app: currentApp,
              week: currentWeek,
              hash: hash,
              expectedHash: expectedHash,
              inCache: false
            });
          }
        }
      } else if ((level === 'CAMPAIGN' || level === 'SOURCE_APP' || level === 'NETWORK') && hash) {
        totalRows++;
        rowsWithHash++;
        
        if (commentsByHash[hash]) {
          matchingHashes++;
        } else {
          // Calculate expected hash based on level
          let expectedHash = '';
          let debugInfo = {};
          
          if (level === 'SOURCE_APP') {
            const sourceAppName = nameOrRange;
            expectedHash = generateDetailedCommentHash('SOURCE_APP', currentApp, currentWeek, 
              sourceAppName, sourceAppName, '', projectName);
            debugInfo = {
              sourceApp: sourceAppName,
              identifier: sourceAppName
            };
          } else if (level === 'CAMPAIGN') {
            const sourceApp = nameOrRange;
            const idValue = row[idCol] || '';
            
            let campaignId = 'N/A';
            if (projectName === 'TRICKY' && idValue.includes('HYPERLINK')) {
              const match = idValue.match(/campaigns\/([^"]+)/);
              campaignId = match ? match[1] : 'N/A';
            }
            
            const identifier = projectName === 'TRICKY' ? campaignId : 'N/A';
            const campaignName = projectName === 'TRICKY' ? 
              (campaignId !== 'N/A' ? campaignId : 'Unknown') : 
              sourceApp;
            
            expectedHash = generateDetailedCommentHash('CAMPAIGN', currentApp, currentWeek, 
              identifier, sourceApp, campaignName, projectName);
            debugInfo = {
              sourceApp: sourceApp,
              campaignId: campaignId,
              identifier: identifier,
              campaignName: campaignName
            };
          } else if (level === 'NETWORK') {
            const networkName = nameOrRange;
            const networkId = row[idCol] || 'N/A';
            expectedHash = generateDetailedCommentHash('NETWORK', currentApp, currentWeek, 
              networkId, 'N/A', networkName, projectName);
            debugInfo = {
              networkName: networkName,
              networkId: networkId
            };
          }
          
          mismatchDetails.push({
            row: i + 1,
            level: level,
            app: currentApp,
            week: currentWeek,
            hash: hash,
            expectedHash: expectedHash,
            inCache: false,
            debugInfo: debugInfo
          });
        }
      }
    }
    
    console.log(`\nSummary:`);
    console.log(`Total data rows: ${totalRows}`);
    console.log(`Rows with hash: ${rowsWithHash}`);
    console.log(`Matching hashes in cache: ${matchingHashes}`);
    console.log(`Mismatches: ${rowsWithHash - matchingHashes}`);
    
    if (mismatchDetails.length > 0) {
      console.log(`\nFirst 10 mismatches:`);
      mismatchDetails.slice(0, 10).forEach(detail => {
        console.log(`\n  Row ${detail.row}: ${detail.level}`);
        console.log(`    App: ${detail.app}`);
        console.log(`    Week: ${detail.week}`);
        console.log(`    Sheet hash: ${detail.hash}`);
        console.log(`    Expected: ${detail.expectedHash}`);
        console.log(`    Hash match: ${detail.hash === detail.expectedHash}`);
        if (detail.debugInfo) {
          console.log(`    Debug info:`, JSON.stringify(detail.debugInfo, null, 2));
        }
      });
      
      // Попробуем найти похожие хеши в кеше
      console.log(`\nSearching for similar hashes in cache...`);
      const firstMismatch = mismatchDetails[0];
      if (firstMismatch) {
        const prefix = firstMismatch.hash.substring(0, 5);
        const similarInCache = Object.keys(commentsByHash).filter(h => h.startsWith(prefix));
        console.log(`Hashes starting with ${prefix}: ${similarInCache.length}`);
        if (similarInCache.length > 0 && similarInCache.length < 10) {
          console.log(`Similar hashes:`, similarInCache);
        }
      }
    }
    
  } catch (e) {
    console.error('Error debugging hashes:', e);
  }
}

function analyzeCommentCache() {
  const projectName = CURRENT_PROJECT;
  console.log(`\n=== Analyzing comment cache for ${projectName} ===`);
  
  try {
    const cache = new CommentCache(projectName);
    // Сначала нужно получить или создать лист кеша
    cache.getOrCreateCacheSheet();
    
    if (!cache.cacheSheetName) {
      console.log('Error: Cache sheet name is null');
      return;
    }
    
    const cacheRange = `${cache.cacheSheetName}!A:I`;
    const response = cache.getCachedSheetData(cache.cacheSpreadsheetId, cacheRange);
    
    if (!response.values || response.values.length <= 1) {
      console.log('Cache is empty');
      return;
    }
    
    console.log(`Total rows in cache: ${response.values.length - 1}`);
    
    // Анализируем первые 10 строк
    console.log('\nFirst 10 entries in cache:');
    for (let i = 1; i <= Math.min(10, response.values.length - 1); i++) {
      const row = response.values[i];
      const [appName, weekRange, level, identifier, sourceApp, campaign, comment, lastUpdated, hash] = row;
      
      console.log(`\nRow ${i}:`);
      console.log(`  Level: ${level}`);
      console.log(`  App: ${appName}`);
      console.log(`  Week: ${weekRange}`);
      console.log(`  Identifier: ${identifier}`);
      console.log(`  SourceApp: ${sourceApp}`);
      console.log(`  Campaign: ${campaign}`);
      console.log(`  Hash: ${hash}`);
      console.log(`  Comment: ${comment ? comment.substring(0, 50) + '...' : 'No comment'}`);
    }
    
    // Считаем по уровням
    const levelCounts = {};
    for (let i = 1; i < response.values.length; i++) {
      const level = response.values[i][2];
      levelCounts[level] = (levelCounts[level] || 0) + 1;
    }
    
    console.log('\nComment counts by level:');
    Object.entries(levelCounts).forEach(([level, count]) => {
      console.log(`  ${level}: ${count}`);
    });
    
    // Анализируем хеши
    console.log('\nHash analysis:');
    const hashPrefixes = {};
    for (let i = 1; i < response.values.length; i++) {
      const hash = response.values[i][8];
      if (hash) {
        const prefix = hash.substring(0, 7);
        hashPrefixes[prefix] = (hashPrefixes[prefix] || 0) + 1;
      }
    }
    
    Object.entries(hashPrefixes).forEach(([prefix, count]) => {
      console.log(`  ${prefix}: ${count}`);
    });
    
  } catch (e) {
    console.error('Error analyzing cache:', e);
  }
}