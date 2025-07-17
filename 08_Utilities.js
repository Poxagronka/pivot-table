function ensureSheetExists(spreadsheetId, sheetName, recreate = false) {
  console.log(`Обеспечение существования листа: ${sheetName} (recreate: ${recreate})`);
  
  try {
    console.log('Получение информации о таблице...');
    const spreadsheet = Sheets.Spreadsheets.get(spreadsheetId);
    const existingSheet = spreadsheet.sheets.find(s => s.properties.title === sheetName);
    
    const requests = [];
    
    if (existingSheet) {
      console.log(`Существующий лист найден: ${sheetName} (ID: ${existingSheet.properties.sheetId})`);
      if (recreate) {
        console.log('Добавляем запрос на удаление листа...');
        requests.push({
          deleteSheet: {
            sheetId: existingSheet.properties.sheetId
          }
        });
      }
    } else {
      console.log(`Лист не найден: ${sheetName}`);
    }
    
    if (!existingSheet || recreate) {
      console.log('Добавляем запрос на создание листа...');
      requests.push({
        addSheet: {
          properties: {
            title: sheetName,
            index: 0,
            gridProperties: {
              rowCount: 1000,
              columnCount: 20
            }
          }
        }
      });
    }
    
    if (requests.length > 0) {
      console.log(`Выполняем batch update с ${requests.length} запросами...`);
      const response = Sheets.Spreadsheets.batchUpdate({
        requests: requests
      }, spreadsheetId);
      console.log('✅ Batch update выполнен успешно');
      
      if (recreate || !existingSheet) {
        const addSheetResponse = response.replies.find(r => r.addSheet);
        if (addSheetResponse) {
          const newSheetId = addSheetResponse.addSheet.properties.sheetId;
          console.log(`✅ Новый лист создан с ID: ${newSheetId}`);
          return newSheetId;
        }
      }
    } else {
      console.log('Запросы не требуются');
    }
    
    const finalSheetId = existingSheet ? existingSheet.properties.sheetId : 0;
    console.log(`Возвращаем sheet ID: ${finalSheetId}`);
    return finalSheetId;
  } catch (e) {
    console.error('❌ Ошибка обеспечения существования листа:', e);
    return null;
  }
}

function clearAllDataSilent() {
  console.log('=== ОЧИСТКА ДАННЫХ ТЕКУЩЕГО ПРОЕКТА ===');
  const config = getCurrentConfig();
  console.log(`Проект: ${CURRENT_PROJECT}`);
  console.log(`Sheet ID: ${config.SHEET_ID}`);
  console.log(`Sheet Name: ${config.SHEET_NAME}`);
  
  try {
    console.log('Этап 1: Проверка существующего листа...');
    const existingSheet = getSheetByName(config.SHEET_ID, config.SHEET_NAME);
    
    if (existingSheet) {
      console.log('Существующий лист найден');
      console.log('Этап 2: Синхронизация комментариев...');
      try {
        new CommentCache().syncCommentsFromSheet();
        console.log('✅ Комментарии синхронизированы');
      } catch (e) {
        console.log('⚠️ Ошибка синхронизации комментариев:', e);
      }
    } else {
      console.log('Существующий лист не найден');
    }
    
    console.log('Этап 3: Пересоздание листа...');
    const newSheetId = ensureSheetExists(config.SHEET_ID, config.SHEET_NAME, true);
    console.log(`✅ Лист ${config.SHEET_NAME} пересоздан с ID: ${newSheetId}`);
    
    return newSheetId;
    
  } catch (e) {
    console.error('❌ Ошибка очистки данных:', e);
    throw e;
  }
  
  console.log('=== ОЧИСТКА ДАННЫХ ЗАВЕРШЕНА ===');
}

function clearProjectDataSilent(projectName) {
  console.log(`=== ОЧИСТКА ДАННЫХ ПРОЕКТА ${projectName} ===`);
  const config = getProjectConfig(projectName);
  console.log(`Sheet ID: ${config.SHEET_ID}`);
  console.log(`Sheet Name: ${config.SHEET_NAME}`);
  
  try {
    console.log('Этап 1: Проверка существующего листа...');
    const existingSheet = getSheetByName(config.SHEET_ID, config.SHEET_NAME);
    
    if (existingSheet) {
      console.log('Существующий лист найден');
      console.log('Этап 2: Синхронизация комментариев...');
      try {
        new CommentCache(projectName).syncCommentsFromSheet();
        console.log('✅ Комментарии синхронизированы');
      } catch (e) {
        console.log(`⚠️ Ошибка синхронизации комментариев для ${projectName}:`, e);
      }
    } else {
      console.log('Существующий лист не найден');
    }
    
    console.log('Этап 3: Пересоздание листа...');
    const newSheetId = ensureSheetExists(config.SHEET_ID, config.SHEET_NAME, true);
    console.log(`✅ Лист ${projectName} пересоздан с ID: ${newSheetId}`);
    
    return newSheetId;
    
  } catch (e) {
    console.error(`❌ Ошибка очистки данных ${projectName}:`, e);
    throw e;
  }
  
  console.log(`=== ОЧИСТКА ДАННЫХ ${projectName} ЗАВЕРШЕНА ===`);
}

function getSheetByName(spreadsheetId, sheetName) {
  console.log(`Получение листа: ${sheetName} из таблицы ${spreadsheetId}`);
  try {
    const spreadsheet = Sheets.Spreadsheets.get(spreadsheetId);
    console.log(`Таблица получена, листов: ${spreadsheet.sheets.length}`);
    
    const sheet = spreadsheet.sheets.find(s => s.properties.title === sheetName);
    if (sheet) {
      console.log(`✅ Лист найден: ${sheetName} (ID: ${sheet.properties.sheetId})`);
      return sheet;
    } else {
      console.log(`❌ Лист не найден: ${sheetName}`);
      const sheetNames = spreadsheet.sheets.map(s => s.properties.title).join(', ');
      console.log(`Доступные листы: ${sheetNames}`);
      return null;
    }
  } catch (e) {
    console.error('❌ Ошибка получения листа:', e);
    return null;
  }
}

function getOrCreateProjectSheet(projectName) {
  console.log(`Получение или создание листа для проекта: ${projectName}`);
  const config = getProjectConfig(projectName);
  console.log(`Конфигурация: Sheet ID = ${config.SHEET_ID}, Sheet Name = ${config.SHEET_NAME}`);
  
  const sheetId = ensureSheetExists(config.SHEET_ID, config.SHEET_NAME);
  return { sheetId, sheet: getSheetByName(config.SHEET_ID, config.SHEET_NAME) };
}

function sortProjectSheets() {
  console.log('=== СОРТИРОВКА ЛИСТОВ ПРОЕКТОВ ===');
  
  try {
    const projectOrder = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Overall', 'Settings', 'To do'];
    console.log(`Желаемый порядок: ${projectOrder.join(', ')}`);
    
    console.log('Получение информации о таблице...');
    const spreadsheet = Sheets.Spreadsheets.get(MAIN_SHEET_ID);
    console.log(`Найдено листов: ${spreadsheet.sheets.length}`);
    
    const sheets = spreadsheet.sheets.map(s => ({
      id: s.properties.sheetId,
      title: s.properties.title,
      hidden: s.properties.hidden || false,
      index: s.properties.index
    }));
    
    console.log('Текущие листы:');
    sheets.forEach(sheet => {
      console.log(`  ${sheet.index}: ${sheet.title} (ID: ${sheet.id}, hidden: ${sheet.hidden})`);
    });
    
    const projectSheets = [];
    const visibleOtherSheets = [];
    const hiddenSheets = [];
    
    sheets.forEach(sheet => {
      const projectIndex = projectOrder.indexOf(sheet.title);
      if (projectIndex !== -1) {
        projectSheets.push({ ...sheet, order: projectIndex });
      } else if (sheet.hidden) {
        hiddenSheets.push(sheet);
      } else {
        visibleOtherSheets.push(sheet);
      }
    });
    
    console.log(`Проектные листы: ${projectSheets.length}`);
    console.log(`Видимые другие листы: ${visibleOtherSheets.length}`);
    console.log(`Скрытые листы: ${hiddenSheets.length}`);
    
    projectSheets.sort((a, b) => a.order - b.order);
    visibleOtherSheets.sort((a, b) => a.title.localeCompare(b.title));
    hiddenSheets.sort((a, b) => a.title.localeCompare(b.title));
    
    const finalOrder = [...projectSheets, ...visibleOtherSheets, ...hiddenSheets];
    console.log('Финальный порядок:');
    finalOrder.forEach((sheet, index) => {
      console.log(`  ${index}: ${sheet.title}`);
    });
    
    const requests = [];
    finalOrder.forEach((sheet, index) => {
      if (sheet.index !== index) {
        console.log(`Изменение позиции: ${sheet.title} с ${sheet.index} на ${index}`);
        requests.push({
          updateSheetProperties: {
            properties: {
              sheetId: sheet.id,
              index: index
            },
            fields: 'index'
          }
        });
      }
    });
    
    if (requests.length > 0) {
      console.log(`Выполняем сортировку с ${requests.length} запросами...`);
      Sheets.Spreadsheets.batchUpdate({
        requests: requests
      }, MAIN_SHEET_ID);
      console.log('✅ Сортировка выполнена успешно');
    } else {
      console.log('Сортировка не требуется - все листы уже в правильном порядке');
    }
    
  } catch (e) {
    console.error('❌ Ошибка сортировки листов:', e);
    throw e;
  }
  
  console.log('=== СОРТИРОВКА ЛИСТОВ ЗАВЕРШЕНА ===');
}

function getDateRange(days) {
  console.log(`Создание диапазона дат на ${days} дней назад`);
  
  const today = new Date();
  const endDate = new Date(today);
  const startDate = new Date(today);
  startDate.setDate(today.getDate() - days);
  
  const result = {
    from: formatDateForAPI(startDate),
    to: formatDateForAPI(endDate)
  };
  
  console.log(`Диапазон дат: ${result.from} - ${result.to}`);
  return result;
}

function formatDateForAPI(date) {
  if (!date || !(date instanceof Date)) {
    console.error('Некорректная дата для форматирования:', date);
    return new Date().toISOString().split('T')[0];
  }
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

function getMondayOfWeek(date) {
  if (!date || !(date instanceof Date)) {
    console.error('Некорректная дата для получения понедельника:', date);
    return new Date();
  }
  
  const monday = new Date(date);
  const dayOfWeek = monday.getDay();
  const daysToSubtract = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  monday.setDate(monday.getDate() - daysToSubtract);
  monday.setHours(0, 0, 0, 0);
  
  return monday;
}

function getSundayOfWeek(date) {
  if (!date || !(date instanceof Date)) {
    console.error('Некорректная дата для получения воскресенья:', date);
    return new Date();
  }
  
  const sunday = new Date(date);
  const dayOfWeek = sunday.getDay();
  const daysToAdd = dayOfWeek === 0 ? 0 : 7 - dayOfWeek;
  sunday.setDate(sunday.getDate() + daysToAdd);
  sunday.setHours(23, 59, 59, 999);
  
  return sunday;
}

function getWeekRange(date) {
  const monday = getMondayOfWeek(date);
  const sunday = getSundayOfWeek(date);
  
  return {
    start: formatDateForAPI(monday),
    end: formatDateForAPI(sunday)
  };
}

function getCurrentWeekStart() {
  const today = new Date();
  return getMondayOfWeek(today);
}

function isPreviousWeek(weekStartDate) {
  const today = new Date();
  const currentWeekStart = getCurrentWeekStart();
  const previousWeekStart = new Date(currentWeekStart);
  previousWeekStart.setDate(currentWeekStart.getDate() - 7);
  
  const compareDate = new Date(weekStartDate);
  return compareDate.getTime() === previousWeekStart.getTime();
}

function shouldIncludeWeek(weekStartDate) {
  const today = new Date();
  const dayOfWeek = today.getDay();
  const currentWeekStart = getCurrentWeekStart();
  const compareDate = new Date(weekStartDate);
  
  if (compareDate.getTime() >= currentWeekStart.getTime()) {
    return false;
  }
  
  if (dayOfWeek >= 2 || dayOfWeek === 0) {
    return true;
  }
  
  return !isPreviousWeek(weekStartDate);
}

function addDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function subtractDays(date, days) {
  return addDays(date, -days);
}

function isValidDateString(dateString) {
  if (!dateString || typeof dateString !== 'string') return false;
  
  const regex = /^\d{4}-\d{2}-\d{2}$/;
  if (!regex.test(dateString)) return false;
  
  const date = new Date(dateString);
  return date instanceof Date && !isNaN(date.getTime());
}

function parseDate(dateString) {
  if (!isValidDateString(dateString)) {
    console.error('Некорректный формат даты:', dateString);
    return new Date();
  }
  
  return new Date(dateString);
}

function testApiConnection(projectName = null) {
  console.log('=== ТЕСТ API ПОДКЛЮЧЕНИЯ ===');
  
  if (projectName) {
    setCurrentProject(projectName);
    console.log(`Тестируем проект: ${projectName}`);
  } else {
    console.log(`Тестируем текущий проект: ${CURRENT_PROJECT}`);
  }
  
  try {
    console.log('Этап 1: Проверка конфигурации...');
    const config = getCurrentConfig();
    console.log(`✅ Конфигурация загружена: ${config.SHEET_NAME}`);
    
    console.log('Этап 2: Проверка Bearer Token...');
    if (!config.BEARER_TOKEN || config.BEARER_TOKEN.length < 50) {
      throw new Error('Bearer Token не настроен или слишком короткий');
    }
    console.log(`✅ Bearer Token найден (длина: ${config.BEARER_TOKEN.length})`);
    
    console.log('Этап 3: Создание тестового запроса...');
    const dateRange = getDateRange(7);
    console.log(`✅ Диапазон дат создан: ${dateRange.from} - ${dateRange.to}`);
    
    console.log('Этап 4: Выполнение API запроса...');
    const raw = fetchCampaignData(dateRange);
    console.log(`✅ API запрос выполнен успешно`);
    
    console.log('Этап 5: Проверка структуры ответа...');
    if (!raw.data?.analytics?.richStats?.stats) {
      throw new Error('Неожиданная структура ответа API');
    }
    
    const recordCount = raw.data.analytics.richStats.stats.length;
    console.log(`✅ Получено записей: ${recordCount}`);
    
    return {
      success: true,
      project: CURRENT_PROJECT,
      recordCount: recordCount,
      dateRange: dateRange
    };
    
  } catch (e) {
    console.error(`❌ Ошибка тестирования API для ${CURRENT_PROJECT}:`, e);
    return {
      success: false,
      project: CURRENT_PROJECT,
      error: e.toString()
    };
  }
}

function quickTestAllProjects() {
  console.log('=== БЫСТРЫЙ ТЕСТ ВСЕХ ПРОЕКТОВ ===');
  
  const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'OVERALL'];
  const results = [];
  
  projects.forEach(project => {
    console.log(`\n--- Тестирование ${project} ---`);
    const result = testApiConnection(project);
    results.push(result);
    
    if (result.success) {
      console.log(`✅ ${project}: ${result.recordCount} записей`);
    } else {
      console.log(`❌ ${project}: ${result.error}`);
    }
  });
  
  console.log('\n=== РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ ===');
  const successful = results.filter(r => r.success).length;
  console.log(`Успешно: ${successful}/${projects.length} проектов`);
  
  return results;
}