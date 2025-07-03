/**
 * Debug Functions - Сокращенная версия для всех проектов - ОБНОВЛЕНО: добавлен Mintegral
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
    logDebug(debugSheet, 'Target eROAS: ' + config.TARGET_EROAS + '%', 'INFO');
    
    if (config.BEARER_TOKEN && config.BEARER_TOKEN.length > 50) {
      logDebug(debugSheet, 'Bearer Token: Найден (длина: ' + config.BEARER_TOKEN.length + ')', 'SUCCESS');
    } else {
      logDebug(debugSheet, 'Bearer Token: Отсутствует или слишком короткий!', 'ERROR');
    }
    
    logDebug(debugSheet, 'API Фильтры:', 'INFO');
    logDebug(debugSheet, '- Users: ' + apiConfig.FILTERS.USER.length + ' элементов', 'INFO', JSON.stringify(apiConfig.FILTERS.USER));
    logDebug(debugSheet, '- Attribution Partner: ' + apiConfig.FILTERS.ATTRIBUTION_PARTNER.join(', '), 'INFO');
    logDebug(debugSheet, '- Attribution Network HID: ' + apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.join(', '), 'INFO');
    
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
      { dimension: "ATTRIBUTION_PARTNER", values: apiConfig.FILTERS.ATTRIBUTION_PARTNER, include: true },
      { dimension: "ATTRIBUTION_NETWORK_HID", values: apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID, include: true }
    ];
    
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
    
    const dateDimension = (projectName === 'GOOGLE_ADS' || projectName === 'APPLOVIN') ? 'DATE' : 'INSTALL_DATE';
    
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
        havingFilters: [],
        anonymizationMode: "OFF",
        topFilter: null,
        revenuePredictionVersion: "",
        isMultiMediation: true
      },
      query: getGraphQLQuery()
    };
    
    logDebug(debugSheet, 'Payload создан', 'SUCCESS', 'Размер: ' + JSON.stringify(payload).length + ' символов');
    
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
          logDebug(debugSheet, `Элемент [${index}]: ${item.__typename || 'неизвестный тип'}`, 'INFO', 
                  JSON.stringify(item, null, 2).substring(0, 200));
        }
      });
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
    
    stats.forEach((row, index) => {
      try {
        if (!Array.isArray(row) || row.length < 8) {
          logDebug(debugSheet, `Запись ${index}: неверная структура (длина: ${row ? row.length : 'undefined'})`, 'WARNING');
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
        
        const campaign = row[1];
        const app = row[2];
        
        if (!campaign || !app) {
          logDebug(debugSheet, `Запись ${index}: отсутствует campaign или app`, 'WARNING');
          errorCount++;
          return;
        }
        
        let campaignName = 'Unknown';
        if (campaign.campaignName) {
          campaignName = campaign.campaignName;
        } else if (campaign.value) {
          campaignName = campaign.value;
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
            campaigns: []
          };
        }
        
        totalProcessed++;
        
        if (index < 3) {
          const shortCampaignName = campaignName.length > 50 ? campaignName.substring(0, 50) + '...' : campaignName;
          logDebug(debugSheet, `Запись ${index}: ${app.name}, ${shortCampaignName}, ${date}`, 'INFO');
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
    const campaignPatterns = new Set();
    
    stats.forEach(row => {
      if (Array.isArray(row) && row.length >= 3) {
        const date = row[0]?.value;
        const campaign = row[1];
        const app = row[2];
        
        if (date) uniqueDates.add(date);
        if (app?.name) uniqueApps.add(app.name);
        
        let campaignName = null;
        if (campaign) {
          if (campaign.campaignName) {
            campaignName = campaign.campaignName;
          } else if (campaign.value) {
            campaignName = campaign.value;
          }
        }
        
        if (campaignName) {
          uniqueCampaigns.add(campaignName);
          
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
    logDebug(debugSheet, '- Уникальных кампаний: ' + uniqueCampaigns.size, 'INFO');
    logDebug(debugSheet, '- Уникальных дат: ' + uniqueDates.size, 'INFO');
    
    const appsList = Array.from(uniqueApps).slice(0, 5);
    logDebug(debugSheet, 'Примеры приложений: ' + appsList.join(', '), 'INFO');
    
    const campaignsList = Array.from(uniqueCampaigns).slice(0, 3);
    logDebug(debugSheet, 'Примеры кампаний:', 'INFO', campaignsList.join('\n'));
    
    const sortedDates = Array.from(uniqueDates).sort();
    if (sortedDates.length > 0) {
      logDebug(debugSheet, `Диапазон дат: ${sortedDates[0]} - ${sortedDates[sortedDates.length - 1]}`, 'INFO');
    }
    
    // Project-specific analysis
    if (projectName === 'MOLOCO' || projectName === 'MINTEGRAL') {
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
    } else if (projectName === 'REGULAR' || projectName === 'GOOGLE_ADS' || projectName === 'APPLOVIN') {
      logDebug(debugSheet, `${projectName}: Фильтр кампаний использует исключение`, 'INFO');
      logDebug(debugSheet, `Network HID: ${apiConfig.FILTERS.ATTRIBUTION_NETWORK_HID.join(', ')}`, 'INFO');
      logDebug(debugSheet, '- Всего кампаний найдено: ' + campaignPatterns.size, 'INFO');
      
      if (campaignPatterns.size > 0) {
        const examples = Array.from(campaignPatterns).slice(0, 5);
        logDebug(debugSheet, 'Примеры кампаний:', 'INFO', examples.join('\n'));
        logDebug(debugSheet, `✅ ${projectName} данные корректны!`, 'SUCCESS');
      } else {
        logDebug(debugSheet, 'ПРОБЛЕМА: Кампании не найдены!', 'ERROR');
      }
    } else {
      // Tricky logic
      const searchPattern = apiConfig.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH;
      logDebug(debugSheet, `Паттерн поиска кампаний: ${searchPattern}`, 'INFO');
      logDebug(debugSheet, '- Кампаний, соответствующих паттерну: ' + campaignPatterns.size, 'INFO');
      
      if (campaignPatterns.size === 0) {
        logDebug(debugSheet, 'ПРОБЛЕМА: Ни одна кампания не соответствует паттерну поиска!', 'ERROR');
        logDebug(debugSheet, 'РЕШЕНИЕ: Проверьте правильность паттерна ATTRIBUTION_CAMPAIGN_SEARCH', 'INFO');
        
        if (uniqueCampaigns.size > 0) {
          const examples = Array.from(uniqueCampaigns).slice(0, 10);
          logDebug(debugSheet, 'Примеры найденных кампаний для анализа:', 'INFO', examples.join('\n'));
        }
      } else {
        logDebug(debugSheet, `Фильтр кампаний работает корректно для ${projectName}`, 'SUCCESS');
        const examples = Array.from(campaignPatterns).slice(0, 5);
        logDebug(debugSheet, 'Примеры найденных кампаний:', 'INFO', examples.join('\n'));
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
      ui.alert(`${projectName} API Проверка`, `API работает: получено ${count} записей за последние 7 дней.`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert(`${projectName} API Проверка`, 'Ошибка API: ' + e.toString(), ui.ButtonSet.OK);
  }
}