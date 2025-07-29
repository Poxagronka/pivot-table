/**
 * Utility Functions - ОБНОВЛЕНО: улучшена защита от таймаутов + INCENT_TRAFFIC + умное форматирование валют
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
  const maxRetries = 2;
  const baseDelay = 3000;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const config = getCurrentConfig();
      const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
      const oldSheet = spreadsheet.getSheetByName(config.SHEET_NAME);
      
      if (oldSheet && oldSheet.getLastRow() > 1) {
        try {
          const cache = new CommentCache();
          cache.syncCommentsFromSheet();
          console.log('Comments cached before clearing sheet');
        } catch (e) {
          console.error('Error caching comments:', e);
        }
      }
      
      Utilities.sleep(1000);
      
      const tempSheetName = config.SHEET_NAME + '_temp_' + Date.now();
      const newSheet = spreadsheet.insertSheet(tempSheetName);
      
      if (oldSheet) {
        spreadsheet.deleteSheet(oldSheet);
      }
      
      newSheet.setName(config.SHEET_NAME);
      
      console.log(`Sheet ${config.SHEET_NAME} recreated successfully`);
      return;
    } catch (e) {
      console.error(`Sheet recreation attempt ${attempt} failed:`, e);
      
      if (attempt === maxRetries) {
        throw e;
      }
      
      const delay = baseDelay * Math.pow(2, attempt - 1);
      console.log(`Waiting ${delay}ms before retry...`);
      Utilities.sleep(delay);
    }
  }
}

function clearProjectDataSilent(projectName) {
  const maxRetries = 3;
  const baseDelay = 5000;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const config = getProjectConfig(projectName);
      const spreadsheet = SpreadsheetApp.openById(config.SHEET_ID);
      
      if (!spreadsheet) {
        throw new Error(`Cannot access spreadsheet ${config.SHEET_ID}`);
      }
      
      const oldSheet = spreadsheet.getSheetByName(config.SHEET_NAME);
      
      if (oldSheet && oldSheet.getLastRow() > 1) {
        try {
          const headers = oldSheet.getRange(1, 1, 1, oldSheet.getLastColumn()).getValues()[0];
          const hasCommentColumn = headers.some(h => 
            h && (h.toString().toLowerCase() === 'comments' || h.toString().toLowerCase() === 'comment')
          );
          
          if (hasCommentColumn) {
            console.log(`${projectName}: Found comment column, syncing comments...`);
            const cache = new CommentCache(projectName);
            cache.syncCommentsFromSheet();
            console.log(`${projectName}: Comments cached successfully`);
          } else {
            console.log(`${projectName}: No comment column found in headers:`, headers);
          }
        } catch (e) {
          if (e.toString().includes('timed out')) {
            console.log(`${projectName}: Timeout while caching comments, continuing anyway`);
          } else {
            console.error(`${projectName}: Error caching comments:`, e);
            console.log(`${projectName}: Continuing without comment cache`);
          }
        }
      }
      
      Utilities.sleep(2000);
      SpreadsheetApp.flush();
      
      const tempSheetName = config.SHEET_NAME + '_temp_' + Date.now();
      console.log(`${projectName}: Creating temporary sheet: ${tempSheetName}`);
      const newSheet = spreadsheet.insertSheet(tempSheetName);
      
      Utilities.sleep(1000);
      SpreadsheetApp.flush();
      
      if (oldSheet) {
        try {
          console.log(`${projectName}: Attempting to delete old sheet...`);
          spreadsheet.deleteSheet(oldSheet);
          console.log(`${projectName}: Old sheet deleted successfully`);
          
          Utilities.sleep(1000);
          SpreadsheetApp.flush();
          
          const checkSheet = spreadsheet.getSheetByName(config.SHEET_NAME);
          if (checkSheet) {
            throw new Error(`Sheet ${config.SHEET_NAME} still exists after deletion`);
          }
        } catch (deleteError) {
          console.error(`${projectName}: Error deleting old sheet:`, deleteError);
          
          const allSheets = spreadsheet.getSheets();
          const existingNames = allSheets.map(s => s.getName());
          console.log(`${projectName}: Existing sheets:`, existingNames);
          
          const tempSheets = allSheets.filter(s => s.getName().startsWith(config.SHEET_NAME + '_temp_'));
          if (tempSheets.length > 1) {
            console.log(`${projectName}: Found ${tempSheets.length} temporary sheets, cleaning up...`);
            tempSheets.slice(1).forEach(sheet => {
              try {
                spreadsheet.deleteSheet(sheet);
                console.log(`${projectName}: Deleted temp sheet: ${sheet.getName()}`);
              } catch (e) {
                console.error(`${projectName}: Failed to delete temp sheet: ${sheet.getName()}`);
              }
            });
          }
          
          if (existingNames.includes(config.SHEET_NAME)) {
            const backupName = config.SHEET_NAME + '_backup_' + Date.now();
            console.log(`${projectName}: Renaming existing sheet to ${backupName}`);
            const existingSheet = spreadsheet.getSheetByName(config.SHEET_NAME);
            existingSheet.setName(backupName);
            
            Utilities.sleep(1000);
            SpreadsheetApp.flush();
          }
        }
      }
      
      Utilities.sleep(1000);
      SpreadsheetApp.flush();
      
      console.log(`${projectName}: Renaming temp sheet to ${config.SHEET_NAME}`);
      const finalCheck = spreadsheet.getSheetByName(config.SHEET_NAME);
      if (finalCheck) {
        const backupName = config.SHEET_NAME + '_old_' + Date.now();
        console.log(`${projectName}: Sheet still exists, renaming to ${backupName}`);
        finalCheck.setName(backupName);
        Utilities.sleep(1000);
        SpreadsheetApp.flush();
      }
      
      newSheet.setName(config.SHEET_NAME);
      
      const allSheetsAfter = spreadsheet.getSheets();
      const tempSheetsAfter = allSheetsAfter.filter(s => 
        s.getName().startsWith(config.SHEET_NAME + '_temp_') || 
        s.getName().startsWith(config.SHEET_NAME + '_backup_') ||
        s.getName().startsWith(config.SHEET_NAME + '_old_')
      );
      
      if (tempSheetsAfter.length > 0) {
        console.log(`${projectName}: Cleaning up ${tempSheetsAfter.length} temporary/backup sheets...`);
        tempSheetsAfter.forEach(sheet => {
          try {
            if (spreadsheet.getSheets().length > 1) {
              spreadsheet.deleteSheet(sheet);
              console.log(`${projectName}: Deleted: ${sheet.getName()}`);
            }
          } catch (e) {
            console.error(`${projectName}: Failed to delete: ${sheet.getName()}`);
          }
        });
      }
      
      console.log(`${projectName}: Sheet recreated successfully`);
      return;
    } catch (e) {
      console.error(`${projectName} sheet recreation attempt ${attempt} failed:`, e);
      
      if (e.toString().includes('timed out') || e.toString().includes('Service Spreadsheets')) {
        const timeoutDelay = baseDelay * Math.pow(2, attempt);
        console.log(`Timeout detected. Waiting ${timeoutDelay}ms before retry...`);
        Utilities.sleep(timeoutDelay);
        
        SpreadsheetApp.flush();
        Utilities.sleep(2000);
      } else if (attempt === maxRetries) {
        throw e;
      } else {
        const delay = baseDelay * Math.pow(1.5, attempt - 1);
        console.log(`Waiting ${delay}ms before retry...`);
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
    
    console.log('Sheet ordering (visible sheets only):');
    console.log('- Project sheets:', projectSheets.map(s => s.name));
    console.log('- Other visible sheets:', otherVisibleSheets.map(s => s.name));
    
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
    return amount.toFixed(2);
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

// Comment Hash Generation
function generateCommentHash(level, appName, weekRange, projectName = null) {
  const project = projectName || CURRENT_PROJECT;
  
  // Нормализуем данные
  const normalizedLevel = (level || '').toString().toUpperCase().trim();
  const normalizedApp = (appName || '').toString().trim();
  const normalizedWeek = (weekRange || '').toString().trim();
  
  // Базовые компоненты для всех проектов
  let hashComponents = [
    project,
    normalizedLevel,
    normalizedApp,
    normalizedWeek
  ];
  
  // Создаем строку для хеширования
  const hashInput = hashComponents.join('|||');
  
  // Простая хеш функция
  let hash = 0;
  for (let i = 0; i < hashInput.length; i++) {
    const char = hashInput.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  
  // Возвращаем хеш с префиксом проекта для уникальности
  return `${project.substring(0, 3)}_${Math.abs(hash).toString(36)}`;
}

function generateDetailedCommentHash(level, appName, weekRange, identifier, sourceApp, campaignOrNetwork, projectName = null) {
  const project = projectName || CURRENT_PROJECT;
  
  if (level === 'WEEK') {
    return generateCommentHash(level, appName, weekRange, project);
  }
  
  const normalizedLevel = (level || '').toString().toUpperCase().trim();
  const normalizedApp = (appName || '').toString().trim();
  const normalizedWeek = (weekRange || '').toString().trim();
  const normalizedId = (identifier || '').toString().trim();
  const normalizedSource = (sourceApp || '').toString().trim();
  const normalizedCampaign = (campaignOrNetwork || '').toString().trim();
  
  let hashComponents = [project, normalizedLevel, normalizedApp, normalizedWeek];
  
  switch (normalizedLevel) {
    case 'CAMPAIGN':
    // Для TRICKY добавляем identifier (campaign ID), для остальных - только sourceApp
    if (project === 'TRICKY' || project === 'TRI') {
        hashComponents.push(normalizedId || 'NO_ID');
    }
    hashComponents.push(normalizedSource || 'NO_SOURCE');
    hashComponents.push(normalizedCampaign || 'NO_CAMPAIGN');
    break;
      
    case 'SOURCE_APP':
      hashComponents.push(normalizedId || normalizedSource || 'NO_SOURCE');
      break;
      
    case 'NETWORK':
      hashComponents.push(normalizedId || 'NO_NETWORK_ID');
      break;
      
    case 'APP':
      hashComponents.push(normalizedId || normalizedApp);
      break;
      
    default:
      hashComponents.push(normalizedId);
      hashComponents.push(normalizedSource);
      hashComponents.push(normalizedCampaign);
  }
  
  hashComponents = hashComponents.filter(c => c && c !== '');
  
  const hashInput = hashComponents.join('|||');
  
  let hash = 0;
  for (let i = 0; i < hashInput.length; i++) {
    const char = hashInput.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  
  const levelPrefix = normalizedLevel.substring(0, 1);
  return `${project.substring(0, 3)}_${levelPrefix}_${Math.abs(hash).toString(36)}`;
}

function migrateCampaignHashes() {
  const ui = SpreadsheetApp.getUi();
  const projects = ['MOLOCO', 'REGULAR', 'GOOGLE_ADS', 'APPLOVIN', 'MINTEGRAL', 'INCENT', 'INCENT_TRAFFIC', 'OVERALL'];
  
  const result = ui.alert(
    '🔄 Migrate Campaign Hashes', 
    `This will fix campaign comment hashes in BOTH main sheets and cache.\n\nProjects: ${projects.join(', ')}\n\nContinue?`, 
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  let totalMigrated = 0;
  let totalFixed = 0;
  
  projects.forEach(projectName => {
    try {
      console.log(`\n🔄 Migrating campaign hashes for ${projectName}...`);
      
      const config = getProjectConfig(projectName);
      const cache = new CommentCache(projectName);
      
      // Migrate main sheet hashes
      const mainRange = `${config.SHEET_NAME}!A:T`;
      const mainResponse = Sheets.Spreadsheets.Values.get(config.SHEET_ID, mainRange);
      
      if (mainResponse.values && mainResponse.values.length > 1) {
        const mainData = mainResponse.values;
        const headers = mainData[0];
        const hashCol = headers.findIndex(h => h === 'RowHash');
        const idCol = headers.findIndex(h => h === 'ID');
        
        if (hashCol !== -1) {
          const mainUpdateRequests = [];
          let currentApp = '';
          let currentWeek = '';
          
          for (let i = 1; i < mainData.length; i++) {
            const row = mainData[i];
            const level = row[0];
            const nameOrRange = row[1];
            
            if (level === 'APP') {
              currentApp = nameOrRange;
            } else if (level === 'WEEK') {
              currentWeek = nameOrRange;
            } else if (level === 'CAMPAIGN' && currentApp && currentWeek) {
              const sourceApp = nameOrRange;
              const idValue = row[idCol] || '';
              
              // Extract campaign ID from hyperlink if present
              let campaignId = 'N/A';
              if (projectName === 'TRICKY' && idValue.includes('HYPERLINK')) {
                const match = idValue.match(/campaigns\/([^"]+)/);
                campaignId = match ? match[1] : 'N/A';
              }
              
              // For non-TRICKY projects, campaign name is sourceApp
              const campaignName = projectName === 'TRICKY' ? 
                (campaignId !== 'N/A' ? campaignId : 'Unknown') : 
                sourceApp;
              
              const newHash = generateDetailedCommentHash(
                'CAMPAIGN', 
                currentApp, 
                currentWeek, 
                projectName === 'TRICKY' ? campaignId : 'N/A',
                sourceApp, 
                campaignName, 
                projectName
              );
              
              mainUpdateRequests.push({
                range: `${config.SHEET_NAME}!${cache.columnNumberToLetter(hashCol + 1)}${i + 1}`,
                values: [[newHash]]
              });
            }
          }
          
          if (mainUpdateRequests.length > 0) {
            const batchUpdateRequest = {
              valueInputOption: 'RAW',
              data: mainUpdateRequests
            };
            
            Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest, config.SHEET_ID);
            console.log(`✅ ${projectName}: Updated ${mainUpdateRequests.length} hashes in main sheet`);
            totalFixed += mainUpdateRequests.length;
          }
        }
      }
      
      // Migrate cache sheet hashes
      cache.getOrCreateCacheSheet();
      
      const cacheRange = `${cache.cacheSheetName}!A:I`;
      const cacheResponse = cache.getCachedSheetData(cache.cacheSpreadsheetId, cacheRange);
      
      if (cacheResponse.values && cacheResponse.values.length > 1) {
        const cacheData = cacheResponse.values;
        const cacheUpdateRequests = [];
        let projectCount = 0;
        
        for (let i = 1; i < cacheData.length; i++) {
          const row = cacheData[i];
          if (row.length >= 6 && row[2] === 'CAMPAIGN') {
            const [appName, weekRange, level, identifier, sourceApp, campaign] = row;
            
            // Determine correct identifier based on project
            let correctIdentifier = identifier;
            if (projectName === 'TRICKY') {
              // Keep existing identifier (campaign ID)
              correctIdentifier = identifier || 'N/A';
            } else {
              // Non-TRICKY projects don't use identifier for campaigns
              correctIdentifier = 'N/A';
            }
            
            // Determine correct campaign name
            let correctCampaign = campaign;
            if (projectName !== 'TRICKY') {
              // For non-TRICKY, campaign name should be sourceApp
              correctCampaign = sourceApp || campaign || 'Unknown';
            }
            
            const newHash = generateDetailedCommentHash(
              'CAMPAIGN', 
              appName, 
              weekRange, 
              correctIdentifier,
              sourceApp || 'N/A', 
              correctCampaign, 
              projectName
            );
            
            cacheUpdateRequests.push({
              range: `${cache.cacheSheetName}!I${i + 1}`,
              values: [[newHash]]
            });
            
            projectCount++;
          }
        }
        
        if (cacheUpdateRequests.length > 0) {
          const batchUpdateRequest = {
            valueInputOption: 'RAW',
            data: cacheUpdateRequests
          };
          
          Sheets.Spreadsheets.Values.batchUpdate(batchUpdateRequest, cache.cacheSpreadsheetId);
          
          const cacheKey = `${cache.cacheSpreadsheetId}_${cacheRange}`;
          delete COMMENT_CACHE_GLOBAL.sheetData[cacheKey];
          delete COMMENT_CACHE_GLOBAL.sheetDataTime[cacheKey];
          
          console.log(`✅ ${projectName}: Updated ${projectCount} hashes in cache`);
          totalMigrated += projectCount;
        }
      }
      
    } catch (e) {
      console.error(`❌ Error migrating ${projectName}:`, e);
    }
  });
  
  ui.alert('Migration Complete', `✅ Migration completed!\n\nFixed ${totalFixed} hashes in main sheets\nFixed ${totalMigrated} hashes in cache sheets`, ui.ButtonSet.OK);
}