function getMondayOfWeek(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(d.setDate(diff));
}

function getSundayOfWeek(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() + (day === 0 ? 0 : 7 - day);
  return new Date(d.setDate(diff));
}

function formatDateForAPI(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function getDateRange(days) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const endDate = new Date(today);
  const startDate = new Date(today);
  startDate.setDate(startDate.getDate() - days + 1);
  return {
    from: Utilities.formatDate(startDate, tz, 'yyyy-MM-dd'),
    to: Utilities.formatDate(endDate, tz, 'yyyy-MM-dd')
  };
}

function isValidDate(dateString) {
  const regex = /^\d{4}-\d{2}-\d{2}$/;
  if (!regex.test(dateString)) return false;
  const date = new Date(dateString);
  return date instanceof Date && !isNaN(date);
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
    ensureSheetExists(config.SHEET_ID, config.SHEET_NAME, true);
    console.log(`✅ Лист ${config.SHEET_NAME} пересоздан успешно`);
    
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
    ensureSheetExists(config.SHEET_ID, config.SHEET_NAME, true);
    console.log(`✅ Лист ${projectName} пересоздан успешно`);
    
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
      Sheets.Spreadsheets.batchUpdate({
        requests: requests
      }, spreadsheetId);
      console.log('✅ Batch update выполнен успешно');
    } else {
      console.log('Запросы не требуются');
    }
    
    return true;
  } catch (e) {
    console.error('❌ Ошибка обеспечения существования листа:', e);
    return false;
  }
}

function getOrCreateProjectSheet(projectName) {
  console.log(`Получение или создание листа для проекта: ${projectName}`);
  const config = getProjectConfig(projectName);
  console.log(`Конфигурация: Sheet ID = ${config.SHEET_ID}, Sheet Name = ${config.SHEET_NAME}`);
  
  ensureSheetExists(config.SHEET_ID, config.SHEET_NAME);
  return getSheetByName(config.SHEET_ID, config.SHEET_NAME);
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

function sanitizeString(str) {
  if (!str) return '';
  return str.toString().trim().replace(/[^\w\s-]/g, '').substring(0, 100);
}

function truncateString(str, maxLength = 50) {
  if (!str) return '';
  if (str.length <= maxLength) return str;
  return str.substring(0, maxLength - 3) + '...';
}

function removeDuplicates(arr) {
  return [...new Set(arr)];
}

function groupBy(arr, keyFn) {
  return arr.reduce((groups, item) => {
    const key = keyFn(item);
    if (!groups[key]) groups[key] = [];
    groups[key].push(item);
    return groups;
  }, {});
}

function sortByProperty(arr, property, ascending = true) {
  return arr.sort((a, b) => {
    const aVal = a[property];
    const bVal = b[property];
    if (aVal === bVal) return 0;
    const comparison = aVal < bVal ? -1 : 1;
    return ascending ? comparison : -comparison;
  });
}

function formatCurrency(amount, currency = 'USD') {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: currency,
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(amount);
}

function formatPercentage(value, decimals = 1) {
  return (value * 100).toFixed(decimals) + '%';
}

function roundToDecimals(num, decimals = 2) {
  return Math.round(num * Math.pow(10, decimals)) / Math.pow(10, decimals);
}

function isValidNumber(value) {
  return typeof value === 'number' && !isNaN(value) && isFinite(value);
}

function safeExecute(fn, fallbackValue = null, context = 'Unknown') {
  try {
    return fn();
  } catch (e) {
    return fallbackValue;
  }
}

function retryWithBackoff(fn, maxRetries = 3, baseDelay = 1000) {
  let attempts = 0;
  
  function attempt() {
    try {
      return fn();
    } catch (e) {
      attempts++;
      if (attempts >= maxRetries) throw e;
      Utilities.sleep(baseDelay * Math.pow(2, attempts - 1));
      return attempt();
    }
  }
  
  return attempt();
}

function measureExecutionTime(fn, label = 'Function') {
  const startTime = new Date().getTime();
  const result = fn();
  const endTime = new Date().getTime();
  return result;
}

function batchOperation(items, batchSize, operation) {
  const results = [];
  for (let i = 0; i < items.length; i += batchSize) {
    const batch = items.slice(i, i + batchSize);
    const batchResults = operation(batch, i);
    results.push(...batchResults);
    if (i + batchSize < items.length) Utilities.sleep(100);
  }
  return results;
}

function getConfiguredProjects() {
  const configured = [];
  Object.keys(PROJECTS).forEach(projectName => {
    const project = PROJECTS[projectName];
    if (project.BEARER_TOKEN && project.API_CONFIG.FILTERS.USER.length > 0) {
      configured.push(projectName);
    }
  });
  return configured;
}

function validateProjectConfig(projectName) {
  if (!PROJECTS[projectName]) {
    return { valid: false, error: `Project ${projectName} does not exist` };
  }
  
  const project = PROJECTS[projectName];
  
  if (!project.BEARER_TOKEN) {
    return { valid: false, error: `Project ${projectName} missing BEARER_TOKEN` };
  }
  
  if (!project.API_CONFIG.FILTERS.USER || project.API_CONFIG.FILTERS.USER.length === 0) {
    return { valid: false, error: `Project ${projectName} missing USER filters` };
  }
  
  if (!project.SHEET_NAME) {
    return { valid: false, error: `Project ${projectName} missing SHEET_NAME` };
  }
  
  return { valid: true };
}

function getProjectStatus(projectName) {
  const validation = validateProjectConfig(projectName);
  if (!validation.valid) {
    return { configured: false, error: validation.error };
  }
  
  const project = PROJECTS[projectName];
  const status = {
    configured: true,
    sheetName: project.SHEET_NAME,
    hasToken: !!project.BEARER_TOKEN,
    userCount: project.API_CONFIG.FILTERS.USER.length,
    campaignSearch: project.API_CONFIG.FILTERS.ATTRIBUTION_CAMPAIGN_SEARCH
  };
  
  if (projectName === 'OVERALL') {
    status.dataType = 'app-level aggregated';
    status.networkFilter = project.API_CONFIG.FILTERS.ATTRIBUTION_NETWORK_HID.length > 0 ? 
      project.API_CONFIG.FILTERS.ATTRIBUTION_NETWORK_HID.join(', ') : 'ALL NETWORKS';
  }
  
  return status;
}

function debugUtilities() {
  console.log('=== UTILITIES DEBUG START ===');
  let passed = 0;
  let failed = 0;
  
  const test = (name, fn) => {
    try {
      const result = fn();
      if (result) {
        console.log(`✅ ${name}`);
        passed++;
      } else {
        console.log(`❌ ${name}`);
        failed++;
      }
    } catch (e) {
      console.log(`❌ ${name}: ${e.toString()}`);
      failed++;
    }
  };
  
  test('getMondayOfWeek', () => {
    const monday = getMondayOfWeek(new Date('2024-01-15'));
    return monday.getDay() === 1;
  });
  
  test('getSundayOfWeek', () => {
    const sunday = getSundayOfWeek(new Date('2024-01-15'));
    return sunday.getDay() === 0;
  });
  
  test('formatDateForAPI', () => {
    const formatted = formatDateForAPI(new Date('2024-01-15'));
    return formatted === '2024-01-15';
  });
  
  test('getDateRange', () => {
    const range = getDateRange(7);
    return range.from && range.to && range.from < range.to;
  });
  
  test('isValidDate', () => {
    return isValidDate('2024-01-15') && !isValidDate('invalid') && !isValidDate('2024-13-45');
  });
  
  test('ensureSheetExists', () => {
    return typeof ensureSheetExists === 'function';
  });
  
  test('getSheetByName', () => {
    return typeof getSheetByName === 'function';
  });
  
  console.log('\n=== SUMMARY ===');
  console.log(`Total: ${passed + failed}`);
  console.log(`Passed: ${passed}`);
  console.log(`Failed: ${failed}`);
  console.log(`Success rate: ${Math.round(passed / (passed + failed) * 100)}%`);
  console.log('=== UTILITIES DEBUG END ===');
  
  return { passed, failed, total: passed + failed };
}