var MAIN_SHEET_ID = '1sU3G0HYgv-xX1UGK4Qa_4jhpc7vndtRyKsojyVx9iaE';
var APPS_DATABASE_ID = '1Z5pJgtg--9EACJL8PVZgJsmeUemv6PKhSsyx9ArChrM';
var APPS_DATABASE_SHEET = 'Apps Database';

var BEARER_TOKEN_CACHE = null;
var BEARER_TOKEN_CACHE_TIME = null;
var FALLBACK_BEARER_TOKEN = null;

function getBearerToken() {
  try {
    const now = new Date().getTime();
    if (BEARER_TOKEN_CACHE && BEARER_TOKEN_CACHE_TIME && (now - BEARER_TOKEN_CACHE_TIME) < 300000) return BEARER_TOKEN_CACHE;
    
    const cachedToken = CacheService.getScriptCache().get('BEARER_TOKEN');
    if (cachedToken) {
      BEARER_TOKEN_CACHE = cachedToken;
      BEARER_TOKEN_CACHE_TIME = now;
      return cachedToken;
    }
    
    const settings = loadSettingsFromSheetWithRetry();
    const token = settings.bearerToken || '';
    
    if (token && token.length > 50) {
      BEARER_TOKEN_CACHE = token;
      BEARER_TOKEN_CACHE_TIME = now;
      FALLBACK_BEARER_TOKEN = token;
      CacheService.getScriptCache().put('BEARER_TOKEN', token, 3600);
    }
    
    return token;
  } catch (e) {
    if (FALLBACK_BEARER_TOKEN) return FALLBACK_BEARER_TOKEN;
    const props = PropertiesService.getScriptProperties();
    const propToken = props.getProperty('BEARER_TOKEN');
    return propToken || '';
  }
}

function getBearerTokenStrict() {
  const token = getBearerToken();
  if (!token || token.length < 50) throw new Error('Bearer token not configured. Please set it in Settings sheet.');
  return token;
}

function isBearerTokenConfigured() { return getBearerToken().length > 50; }

function clearTrickyCaches() {
  BUNDLE_ID_CACHE = {};
  APPS_DB_CACHE = null;
  APPS_DB_CACHE_TIME = null;
}

function getTargetEROAS(projectName, appName = null) {
  try {
    const settings = loadSettingsFromSheetWithRetry();
    if (projectName === 'TRICKY') return settings.targetEROAS.tricky || 250;
    if (appName && appName.toLowerCase().includes('business')) return settings.targetEROAS.business || 140;
    return settings.targetEROAS.ceg || 150;
  } catch (e) {
    if (projectName === 'TRICKY') return 250;
    if (appName && appName.toLowerCase().includes('business')) return 140;
    return 150;
  }
}

function getGrowthThresholds(projectName) {
  try {
    const settings = loadSettingsFromSheetWithRetry();
    return settings.growthThresholds[projectName] || getDefaultGrowthThresholds();
  } catch (e) {
    return getDefaultGrowthThresholds();
  }
}

function getDefaultGrowthThresholds() {
  return {
    healthyGrowth: { minSpendChange: 10, minProfitChange: 5 },
    efficiencyImprovement: { maxSpendDecline: -5, minProfitGrowth: 8 },
    inefficientGrowth: { minSpendChange: 0, maxProfitChange: -8 },
    decliningEfficiency: { minSpendStable: -2, maxSpendGrowth: 10, maxProfitDecline: -4, minProfitDecline: -7 },
    scalingDown: { 
      maxSpendChange: -15,
      efficient: { minProfitChange: 0 },
      moderate: { maxProfitDecline: -10, minProfitDecline: -1 },
      problematic: { maxProfitDecline: -15 }
    },
    moderateGrowthSpend: 3, moderateGrowthProfit: 2,
    minimalGrowth: { maxSpendChange: 2, maxProfitChange: 1 },
    moderateDecline: { 
      maxSpendDecline: -3, maxProfitDecline: -3, spendOptimizationRatio: 1.5,
      efficiencyDropRatio: 1.5, proportionalRatio: 1.3
    },
    stable: { maxAbsoluteChange: 2 }
  };
}

function isAutoCacheEnabled() {
  try { return loadSettingsFromSheetWithRetry().automation.autoCache; } catch (e) { return false; }
}

function isAutoUpdateEnabled() {
  try { return loadSettingsFromSheetWithRetry().automation.autoUpdate; } catch (e) { return false; }
}

function getUpdateTriggersStatus() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const updateFunctions = ['autoUpdateTricky','autoUpdateMoloco','autoUpdateRegular','autoUpdateGoogleAds','autoUpdateApplovin','autoUpdateMintegral','autoUpdateIncent','autoUpdateOverall'];
    const updateTriggers = triggers.filter(t => updateFunctions.includes(t.getHandlerFunction()));
    return {
      enabled: isAutoUpdateEnabled(),
      triggersCount: updateTriggers.length,
      expectedCount: 8,
      isComplete: updateTriggers.length === 8,
      triggersList: updateTriggers.map(t => t.getHandlerFunction())
    };
  } catch (e) {
    return { enabled: false, triggersCount: 0, expectedCount: 8, isComplete: false, triggersList: [] };
  }
}

function loadSettingsFromSheetWithRetry(maxRetries = 3) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try { return loadSettingsFromSheet(); } catch (e) {
      if (attempt === maxRetries) return getDefaultSettings();
      Utilities.sleep(1000 * attempt);
    }
  }
}

function getDefaultSettings() {
  return {
    bearerToken: FALLBACK_BEARER_TOKEN || '',
    targetEROAS: { tricky: 250, business: 140, ceg: 150 },
    automation: { autoCache: false, autoUpdate: false },
    growthThresholds: Object.fromEntries(['TRICKY','MOLOCO','REGULAR','GOOGLE_ADS','APPLOVIN','MINTEGRAL','INCENT','OVERALL'].map(p => [p, getDefaultGrowthThresholds()]))
  };
}

['Tricky','Moloco','Regular','GoogleAds','Applovin','Mintegral','Incent','Overall'].forEach(p => {
  const P = p.toUpperCase().replace('GOOGLEADS','GOOGLE_ADS');
  this[`get${p}TargetEROAS`] = appName => getTargetEROAS(P, appName);
  this[`get${p}GrowthThresholds`] = () => getGrowthThresholds(P);
});

var UNIFIED_MEASURES = [
  { id: "cpi", day: null }, 
  { id: "installs", day: null }, 
  { id: "ipm", day: null },
  { id: "spend", day: null }, 
  { id: "retention_rate", day: 1 }, 
  { id: "roas", day: 1 }, 
  { id: "retention_rate", day: 7 },
  { id: "roas", day: 7 }, 
  { id: "e_arpu_forecast", day: 365 },
  { id: "e_roas_forecast", day: 365 }, 
  { id: "e_profit_forecast", day: 730 },
  { id: "e_roas_forecast", day: 730 }
];

var BASE_API_CONFIG = {
  OPERATION_NAME: "RichStats",
  FILTERS: {
    ATTRIBUTION_PARTNER: ["Stack"],
    ATTRIBUTION_CAMPAIGN_SEARCH: null
  },
  MEASURES: UNIFIED_MEASURES
};

var PROJECT_CONFIGS = {
  TRICKY: { net: ["234187180623265792"], search: "/tricky/i", group: "INSTALL_DATE" },
  MOLOCO: { net: ["445856363109679104"], search: null, group: "INSTALL_DATE" },
  REGULAR: { net: ["234187180623265792"], search: "!/tricky/i", group: "INSTALL_DATE" },
  GOOGLE_ADS: { net: ["378302368699121664"], search: "!/test_creo|creo_test|SL|TL|RnD|adq/i", group: "DATE" },
  APPLOVIN: { net: ["261208778387488768"], search: "!/test_creo|creo_test|SL|TL|RnD|adq/i", group: "DATE" },
  MINTEGRAL: { net: ["756604737398243328"], search: null, group: "INSTALL_DATE" },
  INCENT: { net: ["1580763469207044096","932245122865692672","6958061424287416320","6070852297695428608","5354779956943519744"], search: "!/test_creo|creo_test|SL|TL|RnD|adq/i", group: "DATE" },
  OVERALL: { net: null, search: null, group: "DATE", special: true }
};

var PROJECTS = {};
Object.keys(PROJECT_CONFIGS).forEach(proj => {
  const cfg = PROJECT_CONFIGS[proj];
  const users = proj === 'OVERALL' || proj === 'GOOGLE_ADS' || proj === 'APPLOVIN' || proj === 'INCENT' 
    ? ["79950","127168","157350","150140"] 
    : ["79950","127168","157350","150140","11628","233863","239157"];
  
  PROJECTS[proj] = {
    SHEET_NAME: proj === 'GOOGLE_ADS' ? 'Google_Ads' : proj.charAt(0) + proj.slice(1).toLowerCase(),
    API_URL: 'https://app.appodeal.com/graphql',
    TARGET_EROAS: eval(`get${proj.charAt(0) + proj.slice(1).toLowerCase().replace('_a','A')}TargetEROAS`),
    GROWTH_THRESHOLDS: eval(`get${proj.charAt(0) + proj.slice(1).toLowerCase().replace('_a','A')}GrowthThresholds`),
    BEARER_TOKEN: getBearerTokenStrict,
    COMMENTS_CACHE_SHEET: `CommentsCache_${proj === 'GOOGLE_ADS' ? 'Google_Ads' : proj.charAt(0) + proj.slice(1).toLowerCase()}`,
    APPS_CACHE_SHEET: proj === 'TRICKY' ? 'AppsCache_Tricky' : null,
    API_CONFIG: {
      ...BASE_API_CONFIG,
      FILTERS: {
        ...BASE_API_CONFIG.FILTERS,
        USER: users,
        ATTRIBUTION_NETWORK_HID: cfg.net,
        ATTRIBUTION_CAMPAIGN_SEARCH: cfg.search
      },
      GROUP_BY: cfg.special 
        ? [{ dimension: cfg.group, timeBucket: "WEEK" }, { dimension: "ATTRIBUTION_NETWORK_HID" }, { dimension: "APP" }]
        : [{ dimension: cfg.group, timeBucket: "WEEK" }, { dimension: "ATTRIBUTION_CAMPAIGN_HID" }, { dimension: "APP" }]
    }
  };
});

var CURRENT_PROJECT = 'TRICKY';

function getCurrentConfig() {
  const p = PROJECTS[CURRENT_PROJECT];
  return {
    SHEET_ID: MAIN_SHEET_ID,
    SHEET_NAME: p.SHEET_NAME,
    API_URL: p.API_URL,
    TARGET_EROAS: p.TARGET_EROAS(),
    GROWTH_THRESHOLDS: p.GROWTH_THRESHOLDS(),
    BEARER_TOKEN: p.BEARER_TOKEN(),
    COMMENTS_CACHE_SHEET: p.COMMENTS_CACHE_SHEET,
    APPS_CACHE_SHEET: p.APPS_CACHE_SHEET
  };
}

function getCurrentApiConfig() { return PROJECTS[CURRENT_PROJECT].API_CONFIG; }

function getProjectConfig(projectName) {
  if (!PROJECTS[projectName]) throw new Error('Unknown project: ' + projectName);
  const p = PROJECTS[projectName];
  return {
    SHEET_ID: MAIN_SHEET_ID,
    SHEET_NAME: p.SHEET_NAME,
    API_URL: p.API_URL,
    TARGET_EROAS: p.TARGET_EROAS(),
    GROWTH_THRESHOLDS: p.GROWTH_THRESHOLDS(),
    BEARER_TOKEN: p.BEARER_TOKEN(),
    COMMENTS_CACHE_SHEET: p.COMMENTS_CACHE_SHEET,
    APPS_CACHE_SHEET: p.APPS_CACHE_SHEET
  };
}

function getProjectApiConfig(projectName) {
  if (!PROJECTS[projectName]) throw new Error('Unknown project: ' + projectName);
  return PROJECTS[projectName].API_CONFIG;
}

function setCurrentProject(projectName) {
  if (!PROJECTS[projectName]) throw new Error('Unknown project: ' + projectName);
  if (CURRENT_PROJECT === 'TRICKY' || projectName === 'TRICKY') {
    try { clearTrickyCaches(); } catch (e) {}
  }
  CURRENT_PROJECT = projectName;
}

var TABLE_CONFIG = {
  HEADERS: ['Level','Week Range / Source App','ID','GEO','Spend','Spend WoW %','Installs','CPI','ROAS D-1','IPM','RR D-1','RR D-7','eARPU 365d','eROAS 365d','eROAS 730d','eProfit 730d','eProfit 730d WoW %','Growth Status','Comments'],
  COLUMN_WIDTHS: [[1,80],[2,300],[3,40],[4,40],[5,75],[6,55],[7,55],[8,55],[9,55],[10,55],[11,55],[12,55],[13,55],[14,55],[15,55],[16,75],[17,85],[18,160],[19,250]].map(([c,w]) => ({c,w}))
};

var COLORS = {
  HEADER: { background: '#4285f4', fontColor: 'white' },
  APP_ROW: { background: '#d1e7fe', fontColor: 'black' },
  WEEK_ROW: { background: '#e8f0fe' },
  NETWORK_ROW: { background: '#f0f8ff' },
  SOURCE_APP_ROW: { background: '#f0f8ff' },
  CAMPAIGN_ROW: { background: '#ffffff' },
  POSITIVE: { background: '#d1f2eb', fontColor: '#0c5460' },
  NEGATIVE: { background: '#f8d7da', fontColor: '#721c24' },
  WARNING: { background: '#fff3cd', fontColor: '#856404' },
  INFO: { background: '#d1ecf1', fontColor: '#0c5460' }
};