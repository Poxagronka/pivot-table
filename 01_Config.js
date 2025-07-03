/**
 * Configuration file - ИСПРАВЛЕНО: токен спрятан в Properties без ошибок при инициализации
 */

var MAIN_SHEET_ID = '1sU3G0HYgv-xX1UGK4Qa_4jhpc7vndtRyKsojyVx9iaE';

var PROPERTY_KEYS = {
  BEARER_TOKEN: 'BEARER_TOKEN',
  TARGET_EROAS_TRICKY: 'TARGET_EROAS_TRICKY',
  TARGET_EROAS_MOLOCO: 'TARGET_EROAS_MOLOCO',
  TARGET_EROAS_REGULAR: 'TARGET_EROAS_REGULAR',
  TARGET_EROAS_GOOGLE_ADS: 'TARGET_EROAS_GOOGLE_ADS',
  TARGET_EROAS_APPLOVIN: 'TARGET_EROAS_APPLOVIN',
  TARGET_EROAS_MINTEGRAL: 'TARGET_EROAS_MINTEGRAL',
  AUTO_CACHE_ENABLED: 'AUTO_CACHE_ENABLED',
  AUTO_UPDATE_ENABLED: 'AUTO_UPDATE_ENABLED',
  GROWTH_THRESHOLDS_TRICKY: 'GROWTH_THRESHOLDS_TRICKY',
  GROWTH_THRESHOLDS_MOLOCO: 'GROWTH_THRESHOLDS_MOLOCO',
  GROWTH_THRESHOLDS_REGULAR: 'GROWTH_THRESHOLDS_REGULAR',
  GROWTH_THRESHOLDS_GOOGLE_ADS: 'GROWTH_THRESHOLDS_GOOGLE_ADS',
  GROWTH_THRESHOLDS_APPLOVIN: 'GROWTH_THRESHOLDS_APPLOVIN',
  GROWTH_THRESHOLDS_MINTEGRAL: 'GROWTH_THRESHOLDS_MINTEGRAL'
};

var DEFAULT_TARGET_EROAS = {
  TRICKY: 160,
  MOLOCO: 140,
  REGULAR: 140,
  GOOGLE_ADS: 140,
  APPLOVIN: 140,
  MINTEGRAL: 140
};

var DEFAULT_GROWTH_THRESHOLDS_BASE = {
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
  moderateGrowthSpend: 3,
  moderateGrowthProfit: 2,
  minimalGrowth: { maxSpendChange: 2, maxProfitChange: 1 },
  moderateDecline: { 
    maxSpendDecline: -3, maxProfitDecline: -3, spendOptimizationRatio: 1.5,
    efficiencyDropRatio: 1.5, proportionalRatio: 1.3
  },
  stable: { maxAbsoluteChange: 2 }
};

var DEFAULT_GROWTH_THRESHOLDS = {
  TRICKY: JSON.parse(JSON.stringify(DEFAULT_GROWTH_THRESHOLDS_BASE)),
  MOLOCO: JSON.parse(JSON.stringify(DEFAULT_GROWTH_THRESHOLDS_BASE)),
  REGULAR: JSON.parse(JSON.stringify(DEFAULT_GROWTH_THRESHOLDS_BASE)),
  GOOGLE_ADS: JSON.parse(JSON.stringify(DEFAULT_GROWTH_THRESHOLDS_BASE)),
  APPLOVIN: JSON.parse(JSON.stringify(DEFAULT_GROWTH_THRESHOLDS_BASE)),
  MINTEGRAL: JSON.parse(JSON.stringify(DEFAULT_GROWTH_THRESHOLDS_BASE))
};

// TOKEN MANAGEMENT - ИСПРАВЛЕНО: без ошибок при инициализации
function getBearerToken() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty(PROPERTY_KEYS.BEARER_TOKEN);
  return token || '';
}

function getBearerTokenStrict() {
  var token = getBearerToken();
  if (!token || token.length < 50) {
    throw new Error('Bearer token not configured. Please set it using Settings menu.');
  }
  return token;
}

function setBearerToken(token) {
  if (!token || token.length < 50) {
    throw new Error('Invalid token provided. Token must be at least 50 characters long.');
  }
  var props = PropertiesService.getScriptProperties();
  props.setProperty(PROPERTY_KEYS.BEARER_TOKEN, token);
}

function isBearerTokenConfigured() {
  var token = getBearerToken();
  return token && token.length > 50;
}

function clearBearerToken() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROPERTY_KEYS.BEARER_TOKEN);
}

// TARGET EROAS MANAGEMENT
function getTargetEROAS(projectName) {
  var props = PropertiesService.getScriptProperties();
  var key = 'TARGET_EROAS_' + projectName;
  var value = props.getProperty(key);
  return value ? parseInt(value) : DEFAULT_TARGET_EROAS[projectName];
}

function setTargetEROAS(projectName, value) {
  var props = PropertiesService.getScriptProperties();
  var key = 'TARGET_EROAS_' + projectName;
  props.setProperty(key, value.toString());
}

// GROWTH THRESHOLDS MANAGEMENT
function getGrowthThresholds(projectName) {
  var props = PropertiesService.getScriptProperties();
  var key = 'GROWTH_THRESHOLDS_' + projectName;
  var value = props.getProperty(key);
  
  if (value) {
    try {
      return JSON.parse(value);
    } catch (e) {
      console.error('Error parsing growth thresholds for ' + projectName + ':', e);
      return DEFAULT_GROWTH_THRESHOLDS[projectName];
    }
  }
  
  return DEFAULT_GROWTH_THRESHOLDS[projectName];
}

function setGrowthThresholds(projectName, thresholds) {
  var props = PropertiesService.getScriptProperties();
  var key = 'GROWTH_THRESHOLDS_' + projectName;
  props.setProperty(key, JSON.stringify(thresholds));
}

function resetGrowthThresholds(projectName) {
  var props = PropertiesService.getScriptProperties();
  var key = 'GROWTH_THRESHOLDS_' + projectName;
  props.deleteProperty(key);
}

// PROJECT GETTER FUNCTIONS
function getTrickyTargetEROAS() { return getTargetEROAS('TRICKY'); }
function getMolocoTargetEROAS() { return getTargetEROAS('MOLOCO'); }
function getRegularTargetEROAS() { return getTargetEROAS('REGULAR'); }
function getGoogleAdsTargetEROAS() { return getTargetEROAS('GOOGLE_ADS'); }
function getApplovinTargetEROAS() { return getTargetEROAS('APPLOVIN'); }
function getMintegralTargetEROAS() { return getTargetEROAS('MINTEGRAL'); }

function getTrickyGrowthThresholds() { return getGrowthThresholds('TRICKY'); }
function getMolocoGrowthThresholds() { return getGrowthThresholds('MOLOCO'); }
function getRegularGrowthThresholds() { return getGrowthThresholds('REGULAR'); }
function getGoogleAdsGrowthThresholds() { return getGrowthThresholds('GOOGLE_ADS'); }
function getApplovinGrowthThresholds() { return getGrowthThresholds('APPLOVIN'); }
function getMintegralGrowthThresholds() { return getGrowthThresholds('MINTEGRAL'); }

// PROJECT CONFIGURATIONS
var PROJECTS = {
  TRICKY: {
    SHEET_NAME: 'Tricky',
    API_URL: 'https://app.appodeal.com/graphql',
    TARGET_EROAS: getTrickyTargetEROAS,
    GROWTH_THRESHOLDS: getTrickyGrowthThresholds,
    BEARER_TOKEN: getBearerTokenStrict,
    COMMENTS_CACHE_SHEET: 'CommentsCache_Tricky',
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
      MEASURES: [
        { id: "cpi", day: null }, 
        { id: "installs", day: null }, 
        { id: "ipm", day: null },
        { id: "spend", day: null }, 
        { id: "roas", day: 1 }, 
        { id: "e_arpu_forecast", day: 365 },
        { id: "e_roas_forecast", day: 365 },
        { id: "e_profit_forecast", day: 730 }
      ]
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
      MEASURES: [
        { id: "cpi", day: null }, 
        { id: "installs", day: null }, 
        { id: "ipm", day: null },
        { id: "spend", day: null }, 
        { id: "roas", day: 1 }, 
        { id: "e_arpu_forecast", day: 365 },
        { id: "e_roas_forecast", day: 365 }, 
        { id: "e_profit_forecast", day: 730 }
      ]
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
      MEASURES: [
        { id: "cpi", day: null }, 
        { id: "installs", day: null }, 
        { id: "ipm", day: null },
        { id: "spend", day: null }, 
        { id: "roas", day: 1 }, 
        { id: "e_arpu_forecast", day: 365 },
        { id: "e_roas_forecast", day: 365 }, 
        { id: "e_profit_forecast", day: 730 }
      ]
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
      MEASURES: [
        { id: "cpi", day: null },
        { id: "installs", day: null },
        { id: "spend", day: null },
        { id: "retention_rate", day: 1 },
        { id: "roas", day: 1 },
        { id: "retention_rate", day: 7 },
        { id: "e_roas_forecast", day: 365 },
        { id: "e_profit_forecast", day: 730 }
      ]
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
      MEASURES: [
        { id: "cpi", day: null },
        { id: "installs", day: null },
        { id: "spend", day: null },
        { id: "retention_rate", day: 1 },
        { id: "roas", day: 1 },
        { id: "retention_rate", day: 7 },
        { id: "e_roas_forecast", day: 365 },
        { id: "e_profit_forecast", day: 730 }
      ]
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
      MEASURES: [
        { id: "cpi", day: null }, 
        { id: "installs", day: null }, 
        { id: "ipm", day: null },
        { id: "spend", day: null }, 
        { id: "roas", day: 1 }, 
        { id: "e_arpu_forecast", day: 365 },
        { id: "e_roas_forecast", day: 365 }, 
        { id: "e_profit_forecast", day: 730 }
      ]
    }
  }
};

var CURRENT_PROJECT = 'TRICKY';

// DYNAMIC CONFIG GETTERS
function getCurrentConfig() {
  return {
    SHEET_ID: MAIN_SHEET_ID,
    SHEET_NAME: PROJECTS[CURRENT_PROJECT].SHEET_NAME,
    API_URL: PROJECTS[CURRENT_PROJECT].API_URL,
    TARGET_EROAS: PROJECTS[CURRENT_PROJECT].TARGET_EROAS(),
    GROWTH_THRESHOLDS: PROJECTS[CURRENT_PROJECT].GROWTH_THRESHOLDS(),
    BEARER_TOKEN: PROJECTS[CURRENT_PROJECT].BEARER_TOKEN(),
    COMMENTS_CACHE_SHEET: PROJECTS[CURRENT_PROJECT].COMMENTS_CACHE_SHEET
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
    COMMENTS_CACHE_SHEET: PROJECTS[projectName].COMMENTS_CACHE_SHEET
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
  CURRENT_PROJECT = projectName;
}

// TABLE CONFIGURATION
var TABLE_CONFIG = {
  HEADERS: [
    'Level', 'Week Range / Source App', 'ID', 'GEO',
    'Spend', 'Spend WoW %', 'Installs', 'CPI', 'ROAS D-1', 'IPM',
    'eARPU 365d', 'eROAS 365d', 'eProfit 730d', 'eProfit 730d WoW %', 'Growth Status', 'Comments'
  ],
  COLUMN_WIDTHS: [
    { c: 1, w: 80 },   { c: 2, w: 300 },  { c: 3, w: 50 },   { c: 4, w: 50 },
    { c: 5, w: 75 },   { c: 6, w: 80 },   { c: 7, w: 60 },   { c: 8, w: 60 },
    { c: 9, w: 60 },   { c: 10, w: 50 },  { c: 11, w: 75 },  { c: 12, w: 75 },
    { c: 13, w: 75 },  { c: 14, w: 85 },  { c: 15, w: 160 }, { c: 16, w: 250 }
  ]
};

// COLOR SCHEME
var COLORS = {
  HEADER: { background: '#4285f4', fontColor: 'white' },
  APP_ROW: { background: '#d1e7fe', fontColor: 'black' },
  WEEK_ROW: { background: '#e8f0fe' },
  CAMPAIGN_ROW: { background: '#ffffff' },
  POSITIVE: { background: '#d1f2eb', fontColor: '#0c5460' },
  NEGATIVE: { background: '#f8d7da', fontColor: '#721c24' },
  WARNING: { background: '#fff3cd', fontColor: '#856404' },
  INFO: { background: '#d1ecf1', fontColor: '#0c5460' }
};