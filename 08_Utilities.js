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