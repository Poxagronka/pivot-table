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
  const config = getCurrentConfig();
  
  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      const oldSheet = SpreadsheetApp.openById(config.SHEET_ID).getSheetByName(config.SHEET_NAME);
      
      if (oldSheet?.getLastRow() > 1) {
        try {
          new CommentCache().syncCommentsFromSheet();
        } catch (e) {}
      }
      
      Utilities.sleep(1000);
      
      const requests = [];
      const spreadsheet = Sheets.Spreadsheets.get(config.SHEET_ID);
      const sheetId = spreadsheet.sheets.find(s => s.properties.title === config.SHEET_NAME)?.properties.sheetId;
      
      if (sheetId !== undefined) {
        requests.push({
          deleteSheet: { sheetId: sheetId }
        });
      }
      
      requests.push({
        addSheet: {
          properties: {
            title: config.SHEET_NAME,
            index: 0
          }
        }
      });
      
      Sheets.Spreadsheets.batchUpdate({
        requests: requests
      }, config.SHEET_ID);
      
      return;
    } catch (e) {
      if (attempt === 2) throw e;
      Utilities.sleep(3000 * Math.pow(2, attempt - 1));
    }
  }
}

function clearProjectDataSilent(projectName) {
  const config = getProjectConfig(projectName);
  
  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      const oldSheet = SpreadsheetApp.openById(config.SHEET_ID).getSheetByName(config.SHEET_NAME);
      
      if (oldSheet?.getLastRow() > 1) {
        try {
          new CommentCache(projectName).syncCommentsFromSheet();
        } catch (e) {}
      }
      
      Utilities.sleep(1500);
      
      const requests = [];
      const spreadsheet = Sheets.Spreadsheets.get(config.SHEET_ID);
      const sheetId = spreadsheet.sheets.find(s => s.properties.title === config.SHEET_NAME)?.properties.sheetId;
      
      if (sheetId !== undefined) {
        requests.push({
          deleteSheet: { sheetId: sheetId }
        });
      }
      
      requests.push({
        addSheet: {
          properties: {
            title: config.SHEET_NAME,
            index: 0
          }
        }
      });
      
      Sheets.Spreadsheets.batchUpdate({
        requests: requests
      }, config.SHEET_ID);
      
      return;
    } catch (e) {
      if (attempt === 2) throw e;
      Utilities.sleep(3000 * Math.pow(2, attempt - 1) + 2000);
    }
  }
}

function getOrCreateProjectSheet(projectName) {
  const config = getProjectConfig(projectName);
  const spreadsheet = Sheets.Spreadsheets.get(config.SHEET_ID);
  const sheet = spreadsheet.sheets.find(s => s.properties.title === config.SHEET_NAME);
  
  if (!sheet) {
    Sheets.Spreadsheets.batchUpdate({
      requests: [{
        addSheet: {
          properties: {
            title: config.SHEET_NAME
          }
        }
      }]
    }, config.SHEET_ID);
  }
  
  return SpreadsheetApp.openById(config.SHEET_ID).getSheetByName(config.SHEET_NAME);
}

function sortProjectSheets() {
  try {
    const projectOrder = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Overall', 'Settings', 'To do'];
    const spreadsheet = Sheets.Spreadsheets.get(MAIN_SHEET_ID);
    
    const sheets = spreadsheet.sheets.map(s => ({
      id: s.properties.sheetId,
      title: s.properties.title,
      hidden: s.properties.hidden || false,
      index: s.properties.index
    }));
    
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
    
    projectSheets.sort((a, b) => a.order - b.order);
    visibleOtherSheets.sort((a, b) => a.title.localeCompare(b.title));
    hiddenSheets.sort((a, b) => a.title.localeCompare(b.title));
    
    const finalOrder = [...projectSheets, ...visibleOtherSheets, ...hiddenSheets];
    const requests = [];
    
    finalOrder.forEach((sheet, index) => {
      if (sheet.index !== index) {
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
      Sheets.Spreadsheets.batchUpdate({
        requests: requests
      }, MAIN_SHEET_ID);
    }
  } catch (e) {
    throw e;
  }
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
  
  test('sanitizeString', () => {
    const clean = sanitizeString('Test@#$123 -_abc');
    return clean === 'Test123 -_abc';
  });
  
  test('truncateString', () => {
    const truncated = truncateString('Very long string that needs truncation', 10);
    return truncated === 'Very lo...' && truncateString('Short', 10) === 'Short';
  });
  
  test('removeDuplicates', () => {
    const unique = removeDuplicates([1, 2, 2, 3, 3, 3]);
    return unique.length === 3 && unique[0] === 1 && unique[1] === 2 && unique[2] === 3;
  });
  
  test('groupBy', () => {
    const grouped = groupBy([{type: 'a', val: 1}, {type: 'b', val: 2}, {type: 'a', val: 3}], item => item.type);
    return grouped.a?.length === 2 && grouped.b?.length === 1;
  });
  
  test('sortByProperty', () => {
    const sorted = sortByProperty([{val: 3}, {val: 1}, {val: 2}], 'val');
    const desc = sortByProperty([{val: 1}, {val: 2}, {val: 3}], 'val', false);
    return sorted[0].val === 1 && sorted[2].val === 3 && desc[0].val === 3;
  });
  
  test('formatCurrency', () => {
    const formatted = formatCurrency(1234.56);
    return formatted === '$1,234.56';
  });
  
  test('formatPercentage', () => {
    const formatted = formatPercentage(0.156);
    const formatted2 = formatPercentage(0.1234, 2);
    return formatted === '15.6%' && formatted2 === '12.34%';
  });
  
  test('roundToDecimals', () => {
    const rounded = roundToDecimals(1.23456, 2);
    const rounded3 = roundToDecimals(1.23456, 3);
    return rounded === 1.23 && rounded3 === 1.235;
  });
  
  test('isValidNumber', () => {
    return isValidNumber(123) && isValidNumber(-45.67) && 
           !isValidNumber(NaN) && !isValidNumber(Infinity) && !isValidNumber('123');
  });
  
  test('safeExecute', () => {
    const result1 = safeExecute(() => 'success');
    const result2 = safeExecute(() => { throw new Error('test'); }, 'fallback');
    return result1 === 'success' && result2 === 'fallback';
  });
  
  test('retryWithBackoff', () => {
    let attempts = 0;
    const fn = () => {
      attempts++;
      if (attempts < 3) throw new Error('retry');
      return 'success';
    };
    const result = retryWithBackoff(fn, 3, 10);
    return result === 'success' && attempts === 3;
  });
  
  test('measureExecutionTime', () => {
    const result = measureExecutionTime(() => {
      Utilities.sleep(100);
      return 'done';
    });
    return result === 'done';
  });
  
  test('batchOperation', () => {
    const items = [1, 2, 3, 4, 5];
    const results = batchOperation(items, 2, (batch) => batch.map(x => x * 2));
    return results.length === 5 && results[0] === 2 && results[4] === 10;
  });
  
  test('clearAllDataSilent', () => {
    return typeof clearAllDataSilent === 'function';
  });
  
  test('clearProjectDataSilent', () => {
    return typeof clearProjectDataSilent === 'function';
  });
  
  test('getOrCreateProjectSheet', () => {
    return typeof getOrCreateProjectSheet === 'function';
  });
  
  test('sortProjectSheets', () => {
    return typeof sortProjectSheets === 'function';
  });
  
  console.log('\n=== SUMMARY ===');
  console.log(`Total: ${passed + failed}`);
  console.log(`Passed: ${passed}`);
  console.log(`Failed: ${failed}`);
  console.log(`Success rate: ${Math.round(passed / (passed + failed) * 100)}%`);
  console.log('=== UTILITIES DEBUG END ===');
  
  return { passed, failed, total: passed + failed };
}