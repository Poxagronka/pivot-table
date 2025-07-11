/**
 * Configuration file - ОБНОВЛЕНО: унифицированные метрики + динамические таргеты eROAS D730
 */

var MAIN_SHEET_ID = '1sU3G0HYgv-xX1UGK4Qa_4jhpc7vndtRyKsojyVx9iaE';
var APPS_DATABASE_ID = '1Z5pJgtg--9EACJL8PVZgJsmeUemv6PKhSsyx9ArChrM';
var APPS_DATABASE_SHEET = 'Apps Database';

// Дефолтные значения для eROAS D730 таргетов
var DEFAULT_TARGET_EROAS = {
  TRICKY: 250, MOLOCO: 150, REGULAR: 150, GOOGLE_ADS: 150,
  APPLOVIN: 150, MINTEGRAL: 150, INCENT: 150, OVERALL: 150
};

function getBearerToken() {
  try {
    const settings = loadSettingsFromSheet();
    return settings.bearerToken || '';
  } catch (e) {
    console.error('Error loading bearer token:', e);
    return '';
  }
}

function getBearerTokenStrict() {
  const token = getBearerToken();
  if (!token || token.length < 50) {
    throw new Error('Bearer token not configured. Please set it in Settings sheet.');
  }
  return token;
}

function isBearerTokenConfigured() {
  const token = getBearerToken();
  return token && token.length > 50;
}

function getTargetEROAS(projectName, appName = null) {
  try {
    const settings = loadSettingsFromSheet();
    let baseTarget = settings.targetEROAS[projectName] || DEFAULT_TARGET_EROAS[projectName] || 150;
    
    // Динамический таргет для приложений
    if (appName) {
      if (projectName === 'TRICKY') {
        return 250; // Фиксированный таргет для TRICKY
      } else if (appName.toLowerCase().includes('business')) {
        return 140; // Business приложения
      } else {
        return 150; // Остальные приложения
      }
    }
    
    return baseTarget;
  } catch (e) {
    console.error('Error loading target eROAS:', e);
    if (projectName === 'TRICKY') return 250;
    return appName && appName.toLowerCase().includes('business') ? 140 : 150;
  }
}

function getGrowthThresholds(projectName) {
  try {
    const settings = loadSettingsFromSheet();
    return settings.growthThresholds[projectName] || getDefaultGrowthThresholds();
  } catch (e) {
    console.error('Error loading growth thresholds:', e);
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
  try {
    const settings = loadSettingsFromSheet();
    return settings.automation.autoCache;
  } catch (e) {
    return false;
  }
}

function isAutoUpdateEnabled() {
  try {
    const settings = loadSettingsFromSheet();
    return settings.automation.autoUpdate;
  } catch (e) {
    return false;
  }
}

// Shortcut functions - обновлены для eROAS D730
function getTrickyTargetEROAS(appName) { return getTargetEROAS('TRICKY', appName); }
function getMolocoTargetEROAS(appName) { return getTargetEROAS('MOLOCO', appName); }
function getRegularTargetEROAS(appName) { return getTargetEROAS('REGULAR', appName); }
function getGoogleAdsTargetEROAS(appName) { return getTargetEROAS('GOOGLE_ADS', appName); }
function getApplovinTargetEROAS(appName) { return getTargetEROAS('APPLOVIN', appName); }
function getMintegralTargetEROAS(appName) { return getTargetEROAS('MINTEGRAL', appName); }
function getIncentTargetEROAS(appName) { return getTargetEROAS('INCENT', appName); }
function getOverallTargetEROAS(appName) { return getTargetEROAS('OVERALL', appName); }

function getTrickyGrowthThresholds() { return getGrowthThresholds('TRICKY'); }
function getMolocoGrowthThresholds() { return getGrowthThresholds('MOLOCO'); }
function getRegularGrowthThresholds() { return getGrowthThresholds('REGULAR'); }
function getGoogleAdsGrowthThresholds() { return getGrowthThresholds('GOOGLE_ADS'); }
function getApplovinGrowthThresholds() { return getGrowthThresholds('APPLOVIN'); }
function getMintegralGrowthThresholds() { return getGrowthThresholds('MINTEGRAL'); }
function getIncentGrowthThresholds() { return getGrowthThresholds('INCENT'); }
function getOverallGrowthThresholds() { return getGrowthThresholds('OVERALL'); }

// УНИФИЦИРОВАННЫЕ МЕТРИКИ для всех проектов
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

var PROJECTS = {
  TRICKY: {
    SHEET_NAME: 'Tricky',
    API_URL: 'https://app.appodeal.com/graphql',
    TARGET_EROAS: getTrickyTargetEROAS,
    GROWTH_THRESHOLDS: getTrickyGrowthThresholds,
    BEARER_TOKEN: getBearerTokenStrict,
    COMMENTS_CACHE_SHEET: 'CommentsCache_Tricky',
    APPS_CACHE_SHEET: 'AppsCache_Tricky',
    API_CONFIG: {
      OPERATION_NAME: "RichStats",
      FILTERS: {
        USER: ["79950","127168","157350","150140","11628","233863","239157"],
        ATTRIBUTION_PARTNER: ["Stack"],
        ATTRIBUTION_NETWORK_HID: ["234187180623265792"],
        ATTRIBUTION_CAMPAIGN_SEARCH: "/tricky/i"
      },
      GROUP_BY: [
        { dimension: "INSTALL_DATE", timeBucket: "WEEK" },
        { dimension: "ATTRIBUTION_CAMPAIGN_HID" },
        { dimension: "APP" }
      ],
      MEASURES: UNIFIED_MEASURES
    }
  },
  
  MOLOCO: {
    SHEET_NAME: 'Moloco',
    API_URL: 'https://app.appodeal.com/graphql',
    TARGET_EROAS: getMolocoTargetEROAS,
    GROWTH_THRESHOLDS: getMolocoGrowthThresholds,
    BEARER_TOKEN: getBearerTokenStrict,
    COMMENTS_CACHE_SHEET: 'CommentsCache_Moloco',
    API_CONFIG: {
      OPERATION_NAME: "RichStats",
      FILTERS: {
        USER: ["79950","127168","157350","150140","11628","233863","239157"],
        ATTRIBUTION_PARTNER: ["Stack"],
        ATTRIBUTION_NETWORK_HID: ["445856363109679104"],
        ATTRIBUTION_CAMPAIGN_SEARCH: null
      },
      GROUP_BY: [
        { dimension: "INSTALL_DATE", timeBucket: "WEEK" },
        { dimension: "ATTRIBUTION_CAMPAIGN_HID" },
        { dimension: "APP" }
      ],
      MEASURES: UNIFIED_MEASURES
    }
  },

  REGULAR: {
    SHEET_NAME: 'Regular',
    API_URL: 'https://app.appodeal.com/graphql',
    TARGET_EROAS: getRegularTargetEROAS,
    GROWTH_THRESHOLDS: getRegularGrowthThresholds,
    BEARER_TOKEN: getBearerTokenStrict,
    COMMENTS_CACHE_SHEET: 'CommentsCache_Regular',
    API_CONFIG: {
      OPERATION_NAME: "RichStats",
      FILTERS: {
        USER: ["79950","127168","157350","150140","11628","233863","239157"],
        ATTRIBUTION_PARTNER: ["Stack"],
        ATTRIBUTION_NETWORK_HID: ["234187180623265792"],
        ATTRIBUTION_CAMPAIGN_SEARCH: "!/tricky/i"
      },
      GROUP_BY: [
        { dimension: "INSTALL_DATE", timeBucket: "WEEK" },
        { dimension: "ATTRIBUTION_CAMPAIGN_HID" },
        { dimension: "APP" }
      ],
      MEASURES: UNIFIED_MEASURES
    }
  },

  GOOGLE_ADS: {
    SHEET_NAME: 'Google_Ads',
    API_URL: 'https://app.appodeal.com/graphql',
    TARGET_EROAS: getGoogleAdsTargetEROAS,
    GROWTH_THRESHOLDS: getGoogleAdsGrowthThresholds,
    BEARER_TOKEN: getBearerTokenStrict,
    COMMENTS_CACHE_SHEET: 'CommentsCache_Google_Ads',
    API_CONFIG: {
      OPERATION_NAME: "RichStats",
      FILTERS: {
        USER: ["79950","127168","157350","150140"],
        ATTRIBUTION_PARTNER: ["Stack"],
        ATTRIBUTION_NETWORK_HID: ["378302368699121664"],
        ATTRIBUTION_CAMPAIGN_SEARCH: "!/test_creo|creo_test|SL|TL|RnD|adq/i"
      },
      GROUP_BY: [
        { dimension: "DATE", timeBucket: "WEEK" },
        { dimension: "ATTRIBUTION_CAMPAIGN_HID" },
        { dimension: "APP" }
      ],
      MEASURES: UNIFIED_MEASURES
    }
  },

  APPLOVIN: {
    SHEET_NAME: 'Applovin',
    API_URL: 'https://app.appodeal.com/graphql',
    TARGET_EROAS: getApplovinTargetEROAS,
    GROWTH_THRESHOLDS: getApplovinGrowthThresholds,
    BEARER_TOKEN: getBearerTokenStrict,
    COMMENTS_CACHE_SHEET: 'CommentsCache_Applovin',
    API_CONFIG: {
      OPERATION_NAME: "RichStats",
      FILTERS: {
        USER: ["79950","127168","157350","150140"],
        ATTRIBUTION_PARTNER: ["Stack"],
        ATTRIBUTION_NETWORK_HID: ["261208778387488768"],
        ATTRIBUTION_CAMPAIGN_SEARCH: "!/test_creo|creo_test|SL|TL|RnD|adq/i"
      },
      GROUP_BY: [
        { dimension: "DATE", timeBucket: "WEEK" },
        { dimension: "ATTRIBUTION_CAMPAIGN_HID" },
        { dimension: "APP" }
      ],
      MEASURES: UNIFIED_MEASURES
    }
  },

  MINTEGRAL: {
    SHEET_NAME: 'Mintegral',
    API_URL: 'https://app.appodeal.com/graphql',
    TARGET_EROAS: getMintegralTargetEROAS,
    GROWTH_THRESHOLDS: getMintegralGrowthThresholds,
    BEARER_TOKEN: getBearerTokenStrict,
    COMMENTS_CACHE_SHEET: 'CommentsCache_Mintegral',
    API_CONFIG: {
      OPERATION_NAME: "RichStats",
      FILTERS: {
        USER: ["79950","127168","157350","150140","11628","233863","239157"],
        ATTRIBUTION_PARTNER: ["Stack"],
        ATTRIBUTION_NETWORK_HID: ["756604737398243328"],
        ATTRIBUTION_CAMPAIGN_SEARCH: null
      },
      GROUP_BY: [
        { dimension: "INSTALL_DATE", timeBucket: "WEEK" },
        { dimension: "ATTRIBUTION_CAMPAIGN_HID" },
        { dimension: "APP" }
      ],
      MEASURES: UNIFIED_MEASURES
    }
  },

  INCENT: {
    SHEET_NAME: 'Incent',
    API_URL: 'https://app.appodeal.com/graphql',
    TARGET_EROAS: getIncentTargetEROAS,
    GROWTH_THRESHOLDS: getIncentGrowthThresholds,
    BEARER_TOKEN: getBearerTokenStrict,
    COMMENTS_CACHE_SHEET: 'CommentsCache_Incent',
    API_CONFIG: {
      OPERATION_NAME: "RichStats",
      FILTERS: {
        USER: ["79950","127168","157350","150140"],
        ATTRIBUTION_PARTNER: ["Stack"],
        ATTRIBUTION_NETWORK_HID: ["1580763469207044096","932245122865692672","6958061424287416320","6070852297695428608","5354779956943519744"],
        ATTRIBUTION_CAMPAIGN_SEARCH: "!/test_creo|creo_test|SL|TL|RnD|adq/i"
      },
      GROUP_BY: [
        { dimension: "DATE", timeBucket: "WEEK" },
        { dimension: "ATTRIBUTION_CAMPAIGN_HID" },
        { dimension: "APP" }
      ],
      MEASURES: UNIFIED_MEASURES
    }
  },

  OVERALL: {
    SHEET_NAME: 'Overall',
    API_URL: 'https://app.appodeal.com/graphql',
    TARGET_EROAS: getOverallTargetEROAS,
    GROWTH_THRESHOLDS: getOverallGrowthThresholds,
    BEARER_TOKEN: getBearerTokenStrict,
    COMMENTS_CACHE_SHEET: 'CommentsCache_Overall',
    API_CONFIG: {
      OPERATION_NAME: "RichStats",
      FILTERS: {
        USER: ["79950","127168","157350","150140"],
        ATTRIBUTION_PARTNER: ["Stack"],
        ATTRIBUTION_NETWORK_HID: [],
        ATTRIBUTION_CAMPAIGN_SEARCH: null
      },
      GROUP_BY: [
        { dimension: "DATE", timeBucket: "WEEK" },
        { dimension: "APP" }
      ],
      MEASURES: UNIFIED_MEASURES
    }
  }
};

var CURRENT_PROJECT = 'TRICKY';

function getCurrentConfig() {
  return {
    SHEET_ID: MAIN_SHEET_ID,
    SHEET_NAME: PROJECTS[CURRENT_PROJECT].SHEET_NAME,
    API_URL: PROJECTS[CURRENT_PROJECT].API_URL,
    TARGET_EROAS: PROJECTS[CURRENT_PROJECT].TARGET_EROAS(),
    GROWTH_THRESHOLDS: PROJECTS[CURRENT_PROJECT].GROWTH_THRESHOLDS(),
    BEARER_TOKEN: PROJECTS[CURRENT_PROJECT].BEARER_TOKEN(),
    COMMENTS_CACHE_SHEET: PROJECTS[CURRENT_PROJECT].COMMENTS_CACHE_SHEET,
    APPS_CACHE_SHEET: PROJECTS[CURRENT_PROJECT].APPS_CACHE_SHEET || null
  };
}

function getCurrentApiConfig() {
  return PROJECTS[CURRENT_PROJECT].API_CONFIG;
}

function getProjectConfig(projectName) {
  if (!PROJECTS[projectName]) {
    throw new Error('Unknown project: ' + projectName);
  }
  return {
    SHEET_ID: MAIN_SHEET_ID,
    SHEET_NAME: PROJECTS[projectName].SHEET_NAME,
    API_URL: PROJECTS[projectName].API_URL,
    TARGET_EROAS: PROJECTS[projectName].TARGET_EROAS(),
    GROWTH_THRESHOLDS: PROJECTS[projectName].GROWTH_THRESHOLDS(),
    BEARER_TOKEN: PROJECTS[projectName].BEARER_TOKEN(),
    COMMENTS_CACHE_SHEET: PROJECTS[projectName].COMMENTS_CACHE_SHEET,
    APPS_CACHE_SHEET: PROJECTS[projectName].APPS_CACHE_SHEET || null
  };
}

function getProjectApiConfig(projectName) {
  if (!PROJECTS[projectName]) {
    throw new Error('Unknown project: ' + projectName);
  }
  return PROJECTS[projectName].API_CONFIG;
}

function setCurrentProject(projectName) {
  if (!PROJECTS[projectName]) {
    throw new Error('Unknown project: ' + projectName);
  }
  
  if (CURRENT_PROJECT === 'TRICKY' || projectName === 'TRICKY') {
    try {
      clearTrickyCaches();
    } catch (e) {
      console.log('Cache clear function not available');
    }
  }
  
  CURRENT_PROJECT = projectName;
}

var TABLE_CONFIG = {
  HEADERS: [
    'Level', 'Week Range / Source App', 'ID', 'GEO',
    'Spend', 'Spend WoW %', 'Installs', 'CPI', 'ROAS D-1', 'IPM',
    'RR D-1', 'RR D-7', 'eARPU 365d', 'eROAS 365d', 'eROAS 730d', 'eProfit 730d', 'eProfit 730d WoW %', 'Growth Status', 'Comments'
  ],
  COLUMN_WIDTHS: [
    { c: 1, w: 80 }, { c: 2, w: 300 }, { c: 3, w: 50 }, { c: 4, w: 50 },
    { c: 5, w: 75 }, { c: 6, w: 80 }, { c: 7, w: 60 }, { c: 8, w: 60 },
    { c: 9, w: 60 }, { c: 10, w: 50 }, { c: 11, w: 50 }, { c: 12, w: 50 },
    { c: 13, w: 75 }, { c: 14, w: 75 }, { c: 15, w: 75 }, { c: 16, w: 75 }, 
    { c: 17, w: 85 }, { c: 18, w: 160 }, { c: 19, w: 250 }
  ]
};

var COLORS = {
  HEADER: { background: '#4285f4', fontColor: 'white' },
  APP_ROW: { background: '#d1e7fe', fontColor: 'black' },
  WEEK_ROW: { background: '#e8f0fe' },
  SOURCE_APP_ROW: { background: '#f0f8ff' },
  CAMPAIGN_ROW: { background: '#ffffff' },
  POSITIVE: { background: '#d1f2eb', fontColor: '#0c5460' },
  NEGATIVE: { background: '#f8d7da', fontColor: '#721c24' },
  WARNING: { background: '#fff3cd', fontColor: '#856404' },
  INFO: { background: '#d1ecf1', fontColor: '#0c5460' }
};