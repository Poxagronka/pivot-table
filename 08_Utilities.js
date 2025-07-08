/**
 * Utility Functions - ОБНОВЛЕНО: добавлена поддержка Overall
 */

// Date Utils
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

// Sheet Utils
function expandAllGroups(sheet) {
  try {
    const maxRows = sheet.getMaxRows();
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        sheet.getRange(1, 1, maxRows, 1).expandGroups();
      } catch (e) {
        break;
      }
    }
  } catch (e) {
    console.log('No groups to expand or error expanding groups:', e);
  }
}

function clearAllGroups(sheet) {
  try {
    const maxRows = sheet.getMaxRows();
    let hasGroups = true;
    let attempts = 0;
    
    while (hasGroups && attempts < 10) {
      try {
        sheet.getRange(1, 1, maxRows, 1).shiftRowGroupDepth(-1);
        attempts++;
      } catch (e) {
        hasGroups = false;
      }
    }
  } catch (e) {
    console.log('Error clearing groups:', e);
  }
}

function recreateGrouping(sheet) {
  expandAllGroups(sheet);
  clearAllGroups(sheet);
  const data = sheet.getDataRange().getValues();
  createRowGrouping(sheet, data, null);
}

function clearAllDataSilent() {
  try {
    const config = getCurrentConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
    const oldSheet = spreadsheet.getSheetByName(config.SHEET_NAME);
    
    // ВАЖНО: Кешируем комментарии перед удалением листа
    if (oldSheet && oldSheet.getLastRow() > 1) {
      const cache = new CommentCache();
      cache.syncCommentsFromSheet(); // БЕЗ раскрытия групп
    }
    
    const tempSheetName = config.SHEET_NAME + '_temp_' + Date.now();
    const newSheet = spreadsheet.insertSheet(tempSheetName);
    if (oldSheet) spreadsheet.deleteSheet(oldSheet);
    newSheet.setName(config.SHEET_NAME);
  } catch (e) {
    console.error('Error during sheet recreation:', e);
    throw e;
  }
}

function clearProjectDataSilent(projectName) {
  try {
    const config = getProjectConfig(projectName);
    const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
    const oldSheet = spreadsheet.getSheetByName(config.SHEET_NAME);
    
    // ВАЖНО: Кешируем комментарии перед удалением листа
    if (oldSheet && oldSheet.getLastRow() > 1) {
      const cache = new CommentCache(projectName);
      cache.syncCommentsFromSheet(); // БЕЗ раскрытия групп
    }
    
    const tempSheetName = config.SHEET_NAME + '_temp_' + Date.now();
    const newSheet = spreadsheet.insertSheet(tempSheetName);
    if (oldSheet) spreadsheet.deleteSheet(oldSheet);
    newSheet.setName(config.SHEET_NAME);
  } catch (e) {
    console.error(`Error during ${projectName} sheet recreation:`, e);
    throw e;
  }
}

function getOrCreateProjectSheet(projectName) {
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(config.SHEET_NAME);
  }
  return sheet;
}

// String Utils
function sanitizeString(str) {
  if (!str) return '';
  return str.toString().trim().replace(/[^\w\s-]/g, '').substring(0, 100);
}

function truncateString(str, maxLength = 50) {
  if (!str) return '';
  if (str.length <= maxLength) return str;
  return str.substring(0, maxLength - 3) + '...';
}

// Array Utils
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

// Number Utils
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

// Error Handling
function safeExecute(fn, fallbackValue = null, context = 'Unknown') {
  try {
    return fn();
  } catch (e) {
    console.error(`Error in ${context}:`, e);
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
      if (attempts >= maxRetries) {
        throw e;
      }
      
      const delay = baseDelay * Math.pow(2, attempts - 1);
      console.log(`Attempt ${attempts} failed, retrying in ${delay}ms...`);
      Utilities.sleep(delay);
      return attempt();
    }
  }
  
  return attempt();
}

// Performance Utils
function measureExecutionTime(fn, label = 'Function') {
  const startTime = new Date().getTime();
  const result = fn();
  const endTime = new Date().getTime();
  console.log(`${label} execution time: ${endTime - startTime}ms`);
  return result;
}

function batchOperation(items, batchSize, operation) {
  const results = [];
  for (let i = 0; i < items.length; i += batchSize) {
    const batch = items.slice(i, i + batchSize);
    const batchResults = operation(batch, i);
    results.push(...batchResults);
    
    if (i + batchSize < items.length) {
      Utilities.sleep(100);
    }
  }
  return results;
}

// Project Utils
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
  
  // Для OVERALL добавляем информацию о типе данных
  if (projectName === 'OVERALL') {
    status.dataType = 'app-level aggregated';
    status.networkFilter = project.API_CONFIG.FILTERS.ATTRIBUTION_NETWORK_HID.length > 0 ? 
      project.API_CONFIG.FILTERS.ATTRIBUTION_NETWORK_HID.join(', ') : 'ALL NETWORKS';
  }
  
  return status;
}