var MAIN_SHEET_ID = '1sU3G0HYgv-xX1UGK4Qa_4jhpc7vndtRyKsojyVx9iaE';
var APPS_DATABASE_ID = '1Z5pJgtg--9EACJL8PVZgJsmeUemv6PKhSsyx9ArChrM';
var APPS_DATABASE_SHEET = 'Bundle IDs Database';
var COMMENTS_CACHE_SPREADSHEET_ID = '19A6woiTOP_c7XeKWuLWXKmd-4mO_nZ3aVVxk9ep6mCo';
var APP_NAME_LEGACY = { 'Block-Tok': 'Brick Blast' }; // старое название для поиска в кэшах

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
  if (!token || token.length < 50) throw new Error('Bearer token not configured. Please set it in Settings sheet.');
  return token;
}

function isBearerTokenConfigured() {
  const token = getBearerToken();
  return token && token.length > 50;
}

function getTargetEROAS(projectName, appName = null) {
  projectName = projectName.toUpperCase();
  try {
    const settings = loadSettingsFromSheet();
    if (projectName === 'TRICKY') return settings.targetEROAS.tricky || 250;
    if (appName && appName.toLowerCase().includes('business')) return settings.targetEROAS.business || 140;
    return settings.targetEROAS.ceg || 150;
  } catch (e) {
    console.error('Error loading target eROAS:', e);
    if (projectName === 'TRICKY') return 250;
    if (appName && appName.toLowerCase().includes('business')) return 140;
    return 150;
  }
}

function getGrowthThresholds() {
  try {
    const settings = loadSettingsFromSheet();
    return settings.growthThresholds || getDefaultGrowthThresholds();
  } catch (e) {
    console.error('Error loading growth thresholds:', e);
    return getDefaultGrowthThresholds();
  }
}

function getDefaultGrowthThresholds() {
  return {
    healthyGrowth: { minSpendChange: 10, minProfitChange: 5 },
    efficiencyImprovement: { minSpendChange: -2, maxSpendChange: 2, minProfitGrowth: 5 },
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
  try { return loadSettingsFromSheet().automation.autoCache; } catch (e) { return false; }
}

function isAutoUpdateEnabled() {
  try { return loadSettingsFromSheet().automation.autoUpdate; } catch (e) { return false; }
}

var UNIFIED_MEASURES = [
  { id: "cpi", day: null }, { id: "installs", day: null }, { id: "ipm", day: null },
  { id: "spend", day: null }, { id: "retention_rate", day: 1 }, { id: "roas", day: 1 }, 
  { id: "roas", day: 3 }, { id: "retention_rate", day: 7 }, { id: "roas", day: 7 }, 
  { id: "roas", day: 14 }, { id: "roas", day: 30 }, { id: "e_arpu_forecast", day: 365 },
  { id: "e_roas_forecast", day: 365 }, { id: "e_profit_forecast", day: 730 },
  { id: "e_roas_forecast", day: 730 }
];

// Общие настройки для всех проектов
const API_URL = 'https://app.appodeal.com/graphql';
const OPERATION_NAME = "RichStats";
const ATTRIBUTION_PARTNER = ["Stack"];
const COMMON_USERS = ["79950","127168","157350","150140"];
const EXTENDED_USERS = [...COMMON_USERS,"11628","233863","239157"];

// Фабрика для создания проекта
function createProject(name, network, search, users = EXTENDED_USERS, dateDim = "INSTALL_DATE", extraCfg = {}) {
  return {
    SHEET_NAME: name,
    API_URL,
    BEARER_TOKEN: getBearerTokenStrict,
    COMMENTS_CACHE_SHEET: `CommentsCache_${name}`,
    APPS_CACHE_SHEET: name === 'Tricky' ? 'AppsCache_Tricky' : undefined,
    API_CONFIG: {
      OPERATION_NAME,
      MEASURES: UNIFIED_MEASURES,
      FILTERS: {
        USER: users,
        ATTRIBUTION_PARTNER,
        ATTRIBUTION_NETWORK_HID: Array.isArray(network) ? network : [network],
        ATTRIBUTION_CAMPAIGN_SEARCH: search,
        ...extraCfg.filters
      },
      GROUP_BY: extraCfg.groupBy || [
        { dimension: dateDim, timeBucket: "WEEK" },
        { dimension: "ATTRIBUTION_CAMPAIGN_HID" },
        { dimension: "APP" }
      ]
    }
  };
}

var BASE_PROJECT_CONFIG = {
  API_URL,
  BEARER_TOKEN: getBearerTokenStrict,
  API_CONFIG: { OPERATION_NAME, MEASURES: UNIFIED_MEASURES }
};

var PROJECTS = {
  TRICKY: createProject('Tricky', "234187180623265792", "/tricky/i"),
  MOLOCO: createProject('Moloco', "445856363109679104", null),
  REGULAR: createProject('Regular', "234187180623265792", "!/tricky/i"),
  GOOGLE_ADS: createProject('Google_Ads', "378302368699121664", "!/test_creo|creo_test|SL|TL|RnD|adq/i", COMMON_USERS, "DATE"),
  APPLOVIN: createProject('Applovin', "261208778387488768", "!/test_creo|creo_test|SL|TL|RnD|adq/i", COMMON_USERS, "DATE"),
  APPLOVIN_TEST: createProject('Applovin_test', "261208778387488768", "!/test_creo|creo_test|SL|TL|RnD|adq/i", COMMON_USERS, "DATE", {
    groupBy: [
      { dimension: "DATE", timeBucket: "WEEK" },
      { dimension: "COUNTRY" },
      { dimension: "ATTRIBUTION_CAMPAIGN_HID" },
      { dimension: "APP" }
    ]
  }),
  MINTEGRAL: createProject('Mintegral', "756604737398243328", null),
  INCENT: createProject('Incent', ["1580763469207044096","932245122865692672","6958061424287416320","6070852297695428608","5354779956943519744"], 
    "!/test_creo|creo_test|SL|TL|RnD|adq/i", COMMON_USERS, "DATE"),
  INCENT_TRAFFIC: createProject('Incent_traffic', 
    ["1580763469207044096","932245122865692672","6958061424287416320","6070852297695428608","5354779956943519744","7570550973609803776"],
    null, COMMON_USERS, "DATE", {
      filters: { ATTRIBUTION_CAMPAIGN_EXCLUDE: ["3359685322857250816"] },
      groupBy: [
        { dimension: "DATE", timeBucket: "WEEK" },
        { dimension: "COUNTRY" },
        { dimension: "ATTRIBUTION_NETWORK_HID" },
        { dimension: "ATTRIBUTION_CAMPAIGN_HID" },
        { dimension: "APP" }
      ]
    }),
  OVERALL: createProject('Overall', [], null, COMMON_USERS, "DATE", {
    groupBy: [
      { dimension: "DATE", timeBucket: "WEEK" },
      { dimension: "ATTRIBUTION_NETWORK_HID" },
      { dimension: "APP" }
    ]
  })
};

var CURRENT_PROJECT = 'TRICKY';
var PREVIOUS_PROJECT = null;

function getCurrentConfig() {
  const project = PROJECTS[CURRENT_PROJECT];
  return {
    SHEET_ID: MAIN_SHEET_ID,
    SHEET_NAME: project.SHEET_NAME,
    API_URL: project.API_URL,
    TARGET_EROAS: (appName) => getTargetEROAS(CURRENT_PROJECT, appName),
    GROWTH_THRESHOLDS: () => getGrowthThresholds(),
    BEARER_TOKEN: project.BEARER_TOKEN(),
    COMMENTS_CACHE_SHEET: project.COMMENTS_CACHE_SHEET,
    APPS_CACHE_SHEET: project.APPS_CACHE_SHEET || null
  };
}

function getCurrentApiConfig() { return PROJECTS[CURRENT_PROJECT].API_CONFIG; }

function getProjectConfig(projectName) {
  projectName = projectName.toUpperCase();
  if (!PROJECTS[projectName]) throw new Error('Unknown project: ' + projectName);
  const project = PROJECTS[projectName];
  return {
    SHEET_ID: MAIN_SHEET_ID,
    SHEET_NAME: project.SHEET_NAME,
    API_URL: project.API_URL,
    TARGET_EROAS: (appName) => getTargetEROAS(projectName, appName),
    GROWTH_THRESHOLDS: () => getGrowthThresholds(),
    BEARER_TOKEN: project.BEARER_TOKEN(),
    COMMENTS_CACHE_SHEET: project.COMMENTS_CACHE_SHEET,
    APPS_CACHE_SHEET: project.APPS_CACHE_SHEET || null
  };
}

function getProjectApiConfig(projectName) {
  projectName = projectName.toUpperCase();
  if (!PROJECTS[projectName]) throw new Error('Unknown project: ' + projectName);
  return PROJECTS[projectName].API_CONFIG;
}

function setCurrentProject(projectName) {
  projectName = projectName.toUpperCase();
  if (!PROJECTS[projectName]) throw new Error('Unknown project: ' + projectName);
  PREVIOUS_PROJECT = CURRENT_PROJECT;
  CURRENT_PROJECT = projectName;
}

// Централизованная конфигурация колонок
var COLUMN_CONFIG = {
  // Индексы колонок (1-based для Google Sheets)
  COLUMNS: {
    LEVEL: 1,
    WEEK_RANGE: 2,
    ID: 3,
    GEO: 4,
    SPEND: 5,
    SPEND_WOW: 6,
    INSTALLS: 7,
    CPI: 8,
    ROAS_COMBINED: 9,
    IPM: 10,
    RR_COMBINED: 11,  // Объединенная RR D1→D7
    EARPU: 12,
    EROAS_365: 13,
    EROAS_730: 14,
    EPROFIT_730: 15,
    EPROFIT_WOW: 16,
    GROWTH_STATUS: 17,
    COMMENTS: 18
  },
  
  // Заголовки колонок
  HEADERS: [
    'Level',
    'Week Range / Source App',
    'ID',
    'GEO',
    'Spend',
    'Spend WoW %',
    'Installs',
    'CPI',
    'ROAS D1→D3→D7→D14→D30',
    'IPM',
    'RR D1→D7',  // Объединенная колонка вместо двух отдельных
    'eARPU 365d',
    'eROAS 365d',
    'eROAS 730d (initial → actual)',
    'eProfit 730d (initial → actual)',
    'eProfit 730d WoW %',
    'Growth Status',
    'Comments'
  ],
  
  // Ширины колонок
  WIDTHS: {
    1: 80,   // Level
    2: 350,  // Week Range / Source App
    3: 40,   // ID
    4: 40,   // GEO
    5: 65,   // Spend
    6: 55,   // Spend WoW %
    7: 55,   // Installs
    8: 45,   // CPI
    9: 250,  // ROAS (увеличено для D14)
    10: 37,  // IPM
    11: 85,  // RR D1→D7 (объединенная)
    12: 55,  // eARPU 365d
    13: 55,  // eROAS 365d
    14: 115, // eROAS 730d
    15: 120, // eProfit 730d
    16: 85,  // eProfit 730d WoW %
    17: 160, // Growth Status
    18: 450  // Comments
  },
  
  // Колонки для скрытия по умолчанию
  HIDDEN_COLUMNS: [1, 3, 7, 12, 13]  // Level, ID, Installs, eARPU, eROAS 365d
};

// Обновленный TABLE_CONFIG использующий COLUMN_CONFIG
var TABLE_CONFIG = {
  HEADERS: COLUMN_CONFIG.HEADERS,
  COLUMN_WIDTHS: Object.entries(COLUMN_CONFIG.WIDTHS).map(([c, w]) => ({ c: parseInt(c), w }))
};

var COLORS = {
  HEADER: { background: '#4285f4', fontColor: 'white' },
  APP_ROW: { background: '#d1e7fe', fontColor: 'black' },
  WEEK_ROW: { background: '#e8f0fe' },
  NETWORK_ROW: { background: '#d1e7fe', fontColor: 'black' },
  SOURCE_APP_ROW: { background: '#f0f8ff' },
  CAMPAIGN_ROW: { background: '#ffffff' },
  POSITIVE: { background: '#d1f2eb', fontColor: '#0c5460' },
  NEGATIVE: { background: '#f8d7da', fontColor: '#721c24' },
  WARNING: { background: '#fff3cd', fontColor: '#856404' },
  INFO: { background: '#d1ecf1', fontColor: '#0c5460' }
};

// Централизованная функция для получения всех названий проектов
function getAllProjectNames() {
  return Object.keys(PROJECTS).map(key => {
    // Специальная обработка для APPLOVIN_TEST
    if (key === 'APPLOVIN_TEST') return 'Applovin_test';
    
    // Для остальных - стандартная логика
    return key.split('_').map((part, i) => 
      i === 0 ? part.charAt(0) + part.slice(1).toLowerCase() : part.toLowerCase()
    ).join('_');
  });
}
// Экспорт для обратной совместимости
var ALL_PROJECT_NAMES = getAllProjectNames();