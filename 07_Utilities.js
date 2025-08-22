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

// Sheet Utils - Unified clear function
function clearSheetDataSilent(projectName = null) {
  const maxRetries = projectName ? 3 : 2;
  const baseDelay = projectName ? 5000 : 3000;
  const config = projectName ? getProjectConfig(projectName) : getCurrentConfig();
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
      if (!spreadsheet) throw new Error(`Cannot access spreadsheet ${config.SHEET_ID}`);
      
      const oldSheet = spreadsheet.getSheetByName(config.SHEET_NAME);
      
      // Try to save comments before clearing
      if (oldSheet && oldSheet.getLastRow() > 1) {
        try {
          const cache = new CommentCache(projectName);
          cache.syncCommentsFromSheet();
          console.log(`${projectName || 'Current'}: Comments cached before clearing sheet`);
        } catch (e) {
          console.error(`${projectName || 'Current'}: Error caching comments:`, e);
        }
      }
      
      Utilities.sleep(1000);
      SpreadsheetApp.flush();
      
      // Create new sheet
      const tempSheetName = config.SHEET_NAME + '_temp_' + Date.now();
      const newSheet = spreadsheet.insertSheet(tempSheetName);
      
      // Delete old sheet if exists
      if (oldSheet) {
        try {
          spreadsheet.deleteSheet(oldSheet);
          Utilities.sleep(1000);
        } catch (deleteError) {
          console.error(`Error deleting old sheet:`, deleteError);
          // Rename old sheet as backup
          oldSheet.setName(config.SHEET_NAME + '_backup_' + Date.now());
        }
      }
      
      // Rename new sheet
      newSheet.setName(config.SHEET_NAME);
      
      // Clean up any temp/backup sheets
      spreadsheet.getSheets()
        .filter(s => s.getName().includes(config.SHEET_NAME + '_') && 
                    (s.getName().includes('_temp_') || s.getName().includes('_backup_')))
        .forEach(sheet => {
          try {
            if (spreadsheet.getSheets().length > 1) {
              spreadsheet.deleteSheet(sheet);
            }
          } catch (e) {}
        });
      
      console.log(`${projectName || 'Current'}: Sheet recreated successfully`);
      return;
      
    } catch (e) {
      console.error(`Sheet recreation attempt ${attempt} failed:`, e);
      if (attempt === maxRetries) throw e;
      
      const delay = e.toString().includes('timed out') ? 
                    baseDelay * Math.pow(2, attempt) : 
                    baseDelay * Math.pow(1.5, attempt - 1);
      console.log(`Waiting ${delay}ms before retry...`);
      Utilities.sleep(delay);
      SpreadsheetApp.flush();
    }
  }
}

// Legacy functions for compatibility
function clearAllDataSilent() { clearSheetDataSilent(); }
function clearProjectDataSilent(projectName) { clearSheetDataSilent(projectName); }

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
    
    const projectOrder = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Applovin_test', 'Mintegral', 'Incent', 'Incent_traffic', 'Overall', 'Settings', 'To do'];
    
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
        if (index < finalOrder.length - 1) Utilities.sleep(200);
      } catch (e) {
        console.error(`Error moving sheet ${sheet.getName()}:`, e);
      }
    });
    
    console.log('Visible project sheets sorted successfully');
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
  return str.length <= maxLength ? str : str.substring(0, maxLength - 3) + '...';
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
  if (Math.abs(amount) >= 10) {
    return '$' + amount.toFixed(0);
  } else if (Math.abs(amount) >= 1) {
    return '$' + amount.toFixed(1);
  } else {
    return '$' + amount.toFixed(2);
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
      Utilities.sleep(50);
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
  }
  
  return status;
}

function clearAllCommentColumnCaches() {
  const projects = ['TRICKY', 'MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'INCENT_TRAFFIC', 'OVERALL'];
  projects.forEach(proj => {
    try {
      const cache = new CommentCache(proj);
      cache.clearColumnCache();
    } catch (e) {}
  });
}