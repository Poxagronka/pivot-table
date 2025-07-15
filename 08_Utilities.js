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
  const config = getCurrentConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (sheet && sheet.getLastRow() > 1) {
    try {
      const cache = new CommentCache();
      cache.syncCommentsFromSheet();
      console.log('Comments cached before clearing sheet');
    } catch (e) {
      console.error('Error caching comments:', e);
    }
  }
  
  recreateSheetFast(spreadsheet, config.SHEET_NAME);
}

function clearProjectDataSilent(projectName) {
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  let sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (sheet && sheet.getLastRow() > 1) {
    try {
      const cache = new CommentCache(projectName);
      cache.syncCommentsFromSheet();
      console.log(`${projectName}: Comments cached before clearing sheet`);
    } catch (e) {
      console.error(`${projectName}: Error caching comments:`, e);
    }
  }
  
  recreateSheetFast(spreadsheet, config.SHEET_NAME);
  console.log(`${projectName}: Sheet recreated`);
}

function recreateSheetFast(spreadsheet, sheetName) {
  try {
    const oldSheet = spreadsheet.getSheetByName(sheetName);
    if (oldSheet) {
      spreadsheet.deleteSheet(oldSheet);
    }
    
    const newSheet = spreadsheet.insertSheet(sheetName);
    console.log(`Sheet ${sheetName} recreated`);
  } catch (e) {
    console.error(`Error recreating sheet ${sheetName}:`, e);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      sheet.clear();
      console.log(`Fallback: Sheet ${sheetName} cleared`);
    }
  }
}

function updateProjectDataOptimized(projectName) {
  if (projectName === 'TRICKY') {
    updateProjectDataOptimizedTricky();
    return;
  }
  
  updateProjectDataOptimizedStandard(projectName);
}

function updateProjectDataOptimizedTricky() {
  console.log('=== STARTING TRICKY OPTIMIZED UPDATE ===');
  
  const config = getProjectConfig('TRICKY');
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log('TRICKY: No existing data to update');
    return;
  }
  
  console.log('TRICKY: Initializing optimized cache...');
  const trickyCache = initTrickyOptimizedCache();
  
  console.log('TRICKY: Caching comments...');
  const cache = new CommentCache('TRICKY');
  cache.syncCommentsFromSheet();
  
  console.log('TRICKY: Finding earliest week date...');
  const earliestDate = findEarliestWeekDate(sheet);
  if (!earliestDate) {
    console.log('TRICKY: No week data found');
    return;
  }
  
  const today = new Date();
  const dayOfWeek = today.getDay();
  const endDate = new Date(today);
  
  if (dayOfWeek === 0) {
    endDate.setDate(today.getDate() - 1);
  } else {
    endDate.setDate(today.getDate() - dayOfWeek);
  }
  
  const dateRange = {
    from: formatDateForAPI(earliestDate),
    to: formatDateForAPI(endDate)
  };
  
  console.log(`TRICKY: Fetching optimized data ${dateRange.from} to ${dateRange.to}`);
  
  const raw = fetchProjectCampaignData('TRICKY', dateRange);
  
  if (!raw.data?.analytics?.richStats?.stats?.length) {
    console.log('TRICKY: No API data');
    return;
  }
  
  console.log('TRICKY: Processing API data with optimizations...');
  const originalProject = CURRENT_PROJECT;
  setCurrentProject('TRICKY');
  
  try {
    const processed = processApiData(raw);
    
    if (Object.keys(processed).length === 0) {
      console.log('TRICKY: No valid processed data');
      return;
    }
    
    console.log('TRICKY: Recreating sheet...');
    recreateSheetFast(spreadsheet, config.SHEET_NAME);
    
    console.log('TRICKY: Creating optimized pivot table...');
    createEnhancedPivotTable(processed);
    
    console.log('TRICKY: Applying cached comments...');
    cache.applyCommentsToSheet();
    
    console.log('=== TRICKY OPTIMIZED UPDATE COMPLETED ===');
    console.log(`TRICKY: Cache stats - ${trickyCache?.processed || 0} processed, ${trickyCache?.cacheHits || 0} cache hits`);
    
  } finally {
    setCurrentProject(originalProject);
  }
}

function updateProjectDataOptimizedStandard(projectName) {
  console.log(`Starting standard update for ${projectName}`);
  
  const config = getProjectConfig(projectName);
  const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
  const sheet = spreadsheet.getSheetByName(config.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    console.log(`${projectName}: No existing data to update`);
    return;
  }
  
  const cache = new CommentCache(projectName);
  cache.syncCommentsFromSheet();
  console.log(`${projectName}: Comments cached`);
  
  const earliestDate = findEarliestWeekDate(sheet);
  if (!earliestDate) {
    console.log(`${projectName}: No week data found`);
    return;
  }
  
  const today = new Date();
  const dayOfWeek = today.getDay();
  const endDate = new Date(today);
  
  if (dayOfWeek === 0) {
    endDate.setDate(today.getDate() - 1);
  } else {
    endDate.setDate(today.getDate() - dayOfWeek);
  }
  
  const dateRange = {
    from: formatDateForAPI(earliestDate),
    to: formatDateForAPI(endDate)
  };
  
  console.log(`${projectName}: Fetching ${dateRange.from} to ${dateRange.to}`);
  
  const raw = fetchProjectCampaignData(projectName, dateRange);
  
  if (!raw.data?.analytics?.richStats?.stats?.length) {
    console.log(`${projectName}: No API data`);
    return;
  }
  
  const processed = processProjectApiData(projectName, raw);
  
  if (Object.keys(processed).length === 0) {
    console.log(`${projectName}: No valid processed data`);
    return;
  }
  
  recreateSheetFast(spreadsheet, config.SHEET_NAME);
  
  const originalProject = CURRENT_PROJECT;
  setCurrentProject(projectName);
  try {
    if (projectName === 'OVERALL') {
      createOverallPivotTable(processed);
    } else {
      createEnhancedPivotTable(processed);
    }
    cache.applyCommentsToSheet();
  } finally {
    setCurrentProject(originalProject);
  }
  
  console.log(`${projectName}: Update completed successfully`);
}

function findEarliestWeekDate(sheet) {
  const data = sheet.getDataRange().getValues();
  let earliestDate = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'WEEK') {
      const weekRange = data[i][1];
      const startStr = weekRange.split(' - ')[0];
      const startDate = new Date(startStr);
      if (!earliestDate || startDate < earliestDate) {
        earliestDate = startDate;
      }
    }
  }
  
  return earliestDate;
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
    
    const projectOrder = ['Tricky', 'Moloco', 'Regular', 'Google_Ads', 'Applovin', 'Mintegral', 'Incent', 'Overall', 'Settings', 'To do'];
    
    const projectSheets = [];
    const visibleOtherSheets = [];
    const hiddenSheets = [];
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const projectIndex = projectOrder.indexOf(sheetName);
      const isHidden = sheet.isSheetHidden();
      
      if (projectIndex !== -1) {
        projectSheets.push({ sheet, index: projectIndex, name: sheetName });
      } else if (isHidden) {
        hiddenSheets.push({ sheet, name: sheetName });
      } else {
        visibleOtherSheets.push({ sheet, name: sheetName });
      }
    });
    
    projectSheets.sort((a, b) => a.index - b.index);
    visibleOtherSheets.sort((a, b) => a.name.localeCompare(b.name));
    hiddenSheets.sort((a, b) => a.name.localeCompare(b.name));
    
    const finalOrder = [
      ...projectSheets.map(item => item.sheet),
      ...visibleOtherSheets.map(item => item.sheet),
      ...hiddenSheets.map(item => item.sheet)
    ];
    
    console.log('Sheet ordering:');
    console.log('- Project sheets:', projectSheets.map(s => s.name));
    console.log('- Visible other sheets:', visibleOtherSheets.map(s => s.name));
    console.log('- Hidden sheets (alphabetical):', hiddenSheets.map(s => s.name));
    
    let position = 1;
    
    finalOrder.forEach((sheet, index) => {
      try {
        spreadsheet.setActiveSheet(sheet);
        spreadsheet.moveActiveSheet(position);
        position++;
        
        if (index < finalOrder.length - 1) {
          Utilities.sleep(500);
        }
      } catch (e) {
        console.error(`Error moving sheet ${sheet.getName()}:`, e);
      }
    });
    
    console.log('Project sheets sorted successfully');
  } catch (e) {
    console.error('Error sorting project sheets:', e);
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
      Utilities.sleep(500);
    }
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