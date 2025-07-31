// Date Utils
function getCurrentDateString() {
  return new Date().toISOString().split('T')[0];
}

function formatDateForAPI(date) {
  return date.toISOString().split('T')[0];
}

function getDateRange(days) {
  const today = new Date();
  const fromDate = new Date(today);
  fromDate.setDate(today.getDate() - days + 1);
  
  return {
    from: formatDateForAPI(fromDate),
    to: formatDateForAPI(today)
  };
}

function parseCustomDateRange(customRange) {
  if (!customRange || typeof customRange !== 'string') {
    throw new Error('Custom range must be a string');
  }
  
  const cleanRange = customRange.trim();
  const rangeParts = cleanRange.split(/\s*-\s*/);
  
  if (rangeParts.length !== 2) {
    throw new Error('Range must be in format "YYYY-MM-DD - YYYY-MM-DD"');
  }
  
  const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
  
  if (!dateRegex.test(rangeParts[0]) || !dateRegex.test(rangeParts[1])) {
    throw new Error('Dates must be in YYYY-MM-DD format');
  }
  
  const fromDate = new Date(rangeParts[0] + 'T00:00:00');
  const toDate = new Date(rangeParts[1] + 'T00:00:00');
  
  if (isNaN(fromDate.getTime()) || isNaN(toDate.getTime())) {
    throw new Error('Invalid dates provided');
  }
  
  if (fromDate > toDate) {
    throw new Error('From date must be before or equal to To date');
  }
  
  return {
    from: rangeParts[0],
    to: rangeParts[1]
  };
}

function getMondayOfWeek(date) {
  const d = new Date(date);
  const dayOfWeek = d.getDay();
  const daysToMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  d.setDate(d.getDate() - daysToMonday);
  return d;
}

function getFormattedDateRange(startDate, endDate) {
  const options = { month: 'short', day: 'numeric' };
  const start = new Date(startDate).toLocaleDateString('en-US', options);
  const end = new Date(endDate).toLocaleDateString('en-US', options);
  return `${start} - ${end}`;
}

function isValidDateString(dateStr) {
  const regex = /^\d{4}-\d{2}-\d{2}$/;
  if (!regex.test(dateStr)) return false;
  
  const date = new Date(dateStr + 'T00:00:00');
  return !isNaN(date.getTime()) && date.toISOString().startsWith(dateStr);
}

// Logging Utils
function getFormattedTimestamp() {
  return new Date().toLocaleString('en-US', {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    second: '2-digit',
    hour12: true
  });
}

function logInfo(projectName, recordCount, rowCount, totalTimeMs) {
  const timestamp = getFormattedTimestamp();
  const totalTimeS = (totalTimeMs / 1000).toFixed(1);
  console.log(`${timestamp}  INFO   ✅ ${projectName}: ${recordCount.toLocaleString()} records → ${rowCount.toLocaleString()} rows in ${totalTimeS}s`);
}

function logDebugTiming(timings) {
  const timestamp = getFormattedTimestamp();
  const parts = [];
  
  if (timings.api !== undefined) parts.push(`API: ${(timings.api / 1000).toFixed(1)}s`);
  if (timings.processing !== undefined) parts.push(`Processing: ${(timings.processing / 1000).toFixed(1)}s`);
  if (timings.format !== undefined) parts.push(`Format: ${(timings.format / 1000).toFixed(1)}s`);
  if (timings.grouping !== undefined) parts.push(`Grouping: ${(timings.grouping / 1000).toFixed(1)}s`);
  
  if (parts.length > 0) {
    console.log(`${timestamp}  DEBUG    ${parts.join(' | ')}`);
  }
}

function measureExecutionTime(fn, label = 'Function') {
  const startTime = Date.now();
  const result = fn();
  const endTime = Date.now();
  const executionTime = endTime - startTime;
  
  return { result, executionTime };
}

// Sheet Utils
function clearProjectDataSilent(projectName, preserveComments = true) {
  const maxRetries = 3;
  const baseDelay = 2000;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const config = getProjectConfig(projectName);
      const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
      const oldSheet = spreadsheet.getSheetByName(config.SHEET_NAME);
      
      if (preserveComments && oldSheet && oldSheet.getLastRow() > 1) {
        try {
          const cache = new CommentCache(projectName);
          cache.syncCommentsFromSheet();
        } catch (e) {
          if (e.toString().includes('timed out')) {
          } else {
            console.error(`${projectName}: Error caching comments:`, e);
          }
        }
      }
      
      SpreadsheetApp.flush();
      
      const tempSheetName = config.SHEET_NAME + '_temp_' + Date.now();
      const newSheet = spreadsheet.insertSheet(tempSheetName);
      
      SpreadsheetApp.flush();
      
      if (oldSheet) {
        try {
          spreadsheet.deleteSheet(oldSheet);
        } catch (deleteError) {
          console.error(`${projectName}: Error deleting old sheet, will try to continue:`, deleteError);
          const uniqueName = config.SHEET_NAME + '_' + Date.now();
          newSheet.setName(uniqueName);
         
          try {
            spreadsheet.deleteSheet(oldSheet);
            newSheet.setName(config.SHEET_NAME);
          } catch (e) {
            console.error(`${projectName}: Failed to handle sheet deletion:`, e);
            throw e;
          }
          return;
        }
      }
      
      SpreadsheetApp.flush();
      
      newSheet.setName(config.SHEET_NAME);
      
      return;
    } catch (e) {
      console.error(`${projectName} sheet recreation attempt ${attempt} failed:`, e);
      
      if (e.toString().includes('timed out') || e.toString().includes('Service Spreadsheets')) {
        const timeoutDelay = baseDelay * Math.pow(2, attempt);
        Utilities.sleep(timeoutDelay);
        
        SpreadsheetApp.flush();
      
      } else if (attempt === maxRetries) {
        throw e;
      } else {
        const delay = baseDelay * Math.pow(1.5, attempt - 1);
        Utilities.sleep(delay);
      }
    }
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

function sortProjectSheets() {
  try {
    const spreadsheet = SpreadsheetApp.openById(MAIN_SHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    const projectOrder = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Incent_traffic', 'Overall', 'Settings', 'To do'];
    
    const visibleSheets = sheets.filter(sheet => !sheet.isSheetHidden());
    
    const projectSheets = [];
    const otherVisibleSheets = [];
    
    visibleSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const projectIndex = projectOrder.indexOf(sheetName);
      
      if (projectIndex !== -1) {
        projectSheets.push({ sheet, index: projectIndex, name: sheetName });
      } else {
        otherVisibleSheets.push({ sheet, name: sheetName });
      }
    });
    
    projectSheets.sort((a, b) => a.index - b.index);
    otherVisibleSheets.sort((a, b) => a.name.localeCompare(b.name));
    
    const finalOrder = [
      ...projectSheets.map(item => item.sheet),
      ...otherVisibleSheets.map(item => item.sheet)
    ];
    
    let position = 1;
    
    finalOrder.forEach((sheet, index) => {
      try {
        spreadsheet.setActiveSheet(sheet);
        spreadsheet.moveActiveSheet(position);
        position++;
        
        if (index < finalOrder.length - 1) {
          Utilities.sleep(200);
        }
      } catch (e) {
        console.error(`Error moving sheet ${sheet.getName()}:`, e);
      }
    });
  } catch (e) {
    console.error('Error sorting project sheets:', e);
    throw e;
  }
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

function formatSmartCurrency(amount) {
  if (Math.abs(amount) >= 1) {
    return amount.toFixed(0);
  } else {
    return amount.toFixed(1);
  }
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
      Utilities.sleep(delay);
      return attempt();
    }
  }
  
  return attempt();
}

// Performance Utils
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
  
  if (projectName === 'OVERALL') {
    status.dataType = 'app-level aggregated';
    status.networkFilter = project.API_CONFIG.FILTERS.ATTRIBUTION_NETWORK_HID.length > 0 ? 
      project.API_CONFIG.FILTERS.ATTRIBUTION_NETWORK_HID.join(', ') : 'ALL NETWORKS';
  } else if (projectName === 'INCENT_TRAFFIC') {
    status.dataType = 'network-grouped traffic data';
    status.networkFilter = project.API_CONFIG.FILTERS.ATTRIBUTION_NETWORK_HID.length > 0 ? 
      project.API_CONFIG.FILTERS.ATTRIBUTION_NETWORK_HID.join(', ') : 'ALL NETWORKS';
  } else {
    status.dataType = 'campaign-level';
    status.networkFilter = project.API_CONFIG.FILTERS.ATTRIBUTION_NETWORK_HID.join(', ');
  }
  
  return status;
}

// Cache Utils
function getCacheKey(...parts) {
  return parts.filter(p => p !== null && p !== undefined).join('_');
}

function clearAllCaches() {
  try {
    PropertiesService.getScriptProperties().deleteAll();
  } catch (e) {
    console.error('Error clearing script properties:', e);
  }
}