/*
 * Version 0.01 made by yippym - 2024-12-18 01:05
 * https://github.com/Yippy/signal-tally-sheet
 */
// Signal Tally Const
var SIGNAL_TALLY_SHEET_SOURCE_REDIRECT_ID = '1DOlkMVWDi6-EOyZwSJ55gtpyOdMGo6eQmGYLGZhC0Kk';
var SIGNAL_TALLY_SHEET_SUPPORTED_LOCALE = "en_GB";
var SIGNAL_TALLY_SHEET_TOOLBAR_NAME = "Signal Tally";
var SIGNAL_TALLY_SHEET_SCRIPT_VERSION = 0.01;
var SIGNAL_TALLY_SHEET_SCRIPT_IS_ADD_ON = false;

// Auto Import Const
/* Add URL here to avoid showing on Sheet */
var AUTO_IMPORT_URL_FOR_API_BYPASS = ""; // Optional

class BannerSettings {
  constructor(status_cell_cell, is_enabled_cell, gacha_type) {
    this.status_cell_cell = status_cell_cell;
    this.is_enabled_cell = is_enabled_cell;
    this.gacha_type = gacha_type;
  }

  getStatusText(settingsSheet) {
    return settingsSheet.getRange(this.status_cell_cell).getValue();
  }

  setStatusText(value, settingsSheet) {
    return settingsSheet.getRange(this.status_cell_cell).setValue(value);
  }

  isEnabled(settingsSheet) {
    return settingsSheet.getRange(this.is_enabled_cell).getValue();
  }
}

var AUTO_IMPORT_BANNER_SETTINGS_FOR_IMPORT = {
  "Exclusive Channel Signal History": new BannerSettings("E44", "E37", 2002),
  "Stable Channel Signal History": new BannerSettings("E45", "E38", 1001),
  "W-Engine Channel Signal History": new BannerSettings("E46", "E39", 3002),
  "Bangboo Channel Signal History": new BannerSettings("E47", "E40", 5001)
};

// Auto Import Banner Setting gacha_type is not the same for banner API data gacha_type
var AUTO_IMPORT_LANGUAGE_SETTINGS_FOR_IMPORT = {
  English: {
    code: "en",
    full_code: "en-us",
    "B": " (B-Rank)",
    "A": " (A-Rank)",
    "S": " (S-Rank)",
    gacha_type_2: "Exclusive Channel",
    gacha_type_1: "Stable Channel",
    gacha_type_3: "W-Engine Channel",
    gacha_type_5: "Bangboo Channel",
    item_type_agent: "Agents",
    item_type_w_engine: "W-Engines",
    item_type_bangboo: "Bangboo",
  },
  German: {
    code: "de",
    full_code: "de-de",
    "B": " (Rang-B)",
    "A": " (Rang-A)",
    "S": " (Rang-S)",
    gacha_type_2: "Exklusiver Kanal",
    gacha_type_1: "Stabiler Kanal",
    gacha_type_3: "W-Motor-Kanal",
    gacha_type_5: "Bangboo-Kanal",
    item_type_agent: "Agenten",
    item_type_w_engine: "W-Motoren",
    item_type_bangboo: "Bangboos",
  },
  French: {
    code: "fr",
    full_code: "fr-fr",
    "B": " (rang B)",
    "A": " (rang A)",
    "S": " (rang S)",
    gacha_type_2: "Canal exclusif",
    gacha_type_1: "Canal stable",
    gacha_type_3: "Canal de moteurs-amplis",
    gacha_type_5: "Canal de Bangbous",
    item_type_agent: "Agents",
    item_type_w_engine: "Moteurs-amplis",
    item_type_bangboo: "Bangbous",
  },
  Spanish: {
    code: "es",
    full_code: "es-es",
    "B": " (de grado B)",
    "A": " (de grado A)",
    "S": " (de grado S)",
    gacha_type_2: "Canal Exclusivo",
    gacha_type_1: "Canal Estable",
    gacha_type_3: "Canal Amplificado",
    gacha_type_5: "Canal Bangbú",
    item_type_agent: "Agentes",
    item_type_w_engine: "Amplificadores",
    item_type_bangboo: "Bangbús",
  },
  "Chinese Traditional": {
    code: "zh-tw",
    full_code: "zh-tw",
    "B": " (B級)",
    "A": " (A級)",
    "S": " (S級)",
    gacha_type_2: "獨家頻道",
    gacha_type_1: "常駐頻道",
    gacha_type_3: "音擎頻道",
    gacha_type_5: "邦布頻道",
    item_type_agent: "代理人",
    item_type_w_engine: "音擎",
    item_type_bangboo: "邦布",
  },
  "Chinese Simplified": {
    code: "zh-cn",
    full_code: "zh-cn",
    "B": " (B级)",
    "A": " (A级)",
    "S": " (S级)",
    gacha_type_2: "独家频段",
    gacha_type_1: "常驻频段",
    gacha_type_3: "音擎频段",
    gacha_type_5: "邦布频段",
    item_type_agent: "代理人",
    item_type_w_engine: "音擎",
    item_type_bangboo: "邦布",
  },
  Indonesian: {
    code: "id",
    full_code: "id-id",
    "B": " (Tier-B)",
    "A": " (Tier-A)",
    "S": " (Tier-S)",
    gacha_type_2: "FM Eksklusif",
    gacha_type_1: "FM Stabil",
    gacha_type_3: "FM W-Engine",
    gacha_type_5: "FM Bangboo",
    item_type_agent: "Agen",
    item_type_w_engine: "W-Engine",
    item_type_bangboo: "Bangboo",
  },
  Japanese: {
    code: "ja",
    full_code: "ja-jp",
    "B": " (B級)",
    "A": " (A級)",
    "S": " (S級)",
    gacha_type_2: "独占チャンネル",
    gacha_type_1: "常設チャンネル",
    gacha_type_3: "音動機チャンネル",
    gacha_type_5: "ボンプチャンネル",
    item_type_agent: "エージェント",
    item_type_w_engine: "音動機",
    item_type_bangboo: "ボンプ",
  },
  Vietnamese: {
    code: "vi",
    full_code: "vi-vn",
    "B": " (cấp B)",
    "A": " (cấp A)",
    "S": " (cấp S)",
    gacha_type_2: "Kênh Độc Quyền",
    gacha_type_1: "Kênh Thường Trực",
    gacha_type_3: "Kênh W-Engine",
    gacha_type_5: "Kênh Bangboo",
    item_type_agent: "Người Đại Diện",
    item_type_w_engine: "W-Engine",
    item_type_bangboo: "Bangboo",
  },
  Korean: {
    code: "ko",
    full_code: "ko-kr",
    "B": " (B급)",
    "A": " (A급)",
    "S": " (S급)",
    gacha_type_2: "독점 채널",
    gacha_type_1: "일반 채널",
    gacha_type_3: "W-엔진 채널",
    gacha_type_5: "「Bangboo」 채널",
    item_type_agent: "에이전트",
    item_type_w_engine: "W-엔진",
    item_type_bangboo: "「Bangboo」",
  },
  Portuguese: {
    code: "pt",
    full_code: "pt-pt",
    "B": " (Classe B)",
    "A": " (Classe A)",
    "S": " (Classe S)",
    gacha_type_2: "Canal Exclusivo",
    gacha_type_1: "Canal Estável",
    gacha_type_3: "Canal Motor-W",
    gacha_type_5: "Canal Bangboo",
    item_type_agent: "Agentes",
    item_type_w_engine: "Motores-W",
    item_type_bangboo: "Bangboos",
  },
  Thai: {
    code: "th",
    full_code: "th-th",
    "B": " (แรงก์ B)",
    "A": " (แรงก์ A)",
    "S": " (แรงก์ S)",
    gacha_type_2: "Exclusive Channel",
    gacha_type_1: "Stable Channel",
    gacha_type_3: "W-Engine Channel",
    gacha_type_5: "Bangboo Channel",
    item_type_agent: "Agent",
    item_type_w_engine: "W-Engine",
    item_type_bangboo: "Bangboo",
  },
  Russian: {
    code: "ru",
    full_code: "ru-ru",
    "B": " (ранга B)",
    "A": " (ранга A)",
    "S": " (ранга S)",
    gacha_type_2: "Выделенный канал",
    gacha_type_1: "Стабильный канал",
    gacha_type_3: "Канал амплификаторов",
    gacha_type_5: "Канал банбу",
    item_type_agent: "Агенты",
    item_type_w_engine: "Амплификаторы",
    item_type_bangboo: "Банбу",
  },
};

var AUTO_IMPORT_ADDITIONAL_QUERY = [
  "authkey_ver=1",
  "sign_type=2",
  "auth_appid=webview_gacha",
  "device_type=pc"
];

var AUTO_IMPORT_URL = "https://public-operation-nap-sg.hoyoverse.com/common/gacha_record/api/getGachaLog";
var AUTO_IMPORT_URL_CHINA = "https://public-operation-nap.mihoyo.com/common/gacha_record/api/getGachaLog";

var AUTO_IMPORT_URL_ERROR_CODE_INVALID = -1;
var AUTO_IMPORT_URL_ERROR_CODE_AUTH_KEY_TIMEOUT_MESSAGE = "auth key time out";
var AUTO_IMPORT_URL_ERROR_CODE_AUTH_KEY_INVALID_MESSAGE =  "illegal base64 data at input byte 764";
var AUTO_IMPORT_URL_ERROR_CODE_AUTH_KEY_MISSING_MESSAGE = "auth key or sign type empty";
var AUTO_IMPORT_URL_ERROR_CODE_AUTHKEY_VER_INVALID_MESSAGE = "public key is missing"; // authkey_ver must be 1
var AUTO_IMPORT_URL_ERROR_CODE_AUTHKEY_VER_MISSING_MESSAGE = "auth key or sign type empty"; // authkey_ver is missing

var AUTO_IMPORT_URL_ERROR_CODE_MISSING_PARAMETER = -111;

// Signal Tally Const
var SIGNAL_TALLY_REDIRECT_SOURCE_ABOUT_SHEET_NAME = "About";
var SIGNAL_TALLY_REDIRECT_SOURCE_AUTO_IMPORT_SHEET_NAME = "Auto Import";
var SIGNAL_TALLY_REDIRECT_SOURCE_MAINTENANCE_SHEET_NAME = "Maintenance";
var SIGNAL_TALLY_REDIRECT_SOURCE_HOYOLAB_SHEET_NAME = "HoYoLAB";
var SIGNAL_TALLY_REDIRECT_SOURCE_BACKUP_SHEET_NAME = "Backup";
var SIGNAL_TALLY_EXCLUSIVE_CHANNEL_SIGNAL_SHEET_NAME = "Exclusive Channel Signal History";
var SIGNAL_TALLY_W_ENGINES_SIGNAL_SHEET_NAME = "W-Engine Channel Signal History";
var SIGNAL_TALLY_STABLE_SIGNAL_SHEET_NAME = "Stable Channel Signal History";
var SIGNAL_TALLY_BANGBOO_SIGNAL_SHEET_NAME = "Bangboo Channel Signal History";
var SIGNAL_TALLY_SIGNAL_HISTORY_SHEET_NAME = "Signal History";
var SIGNAL_TALLY_SETTINGS_SHEET_NAME = "Settings";
var SIGNAL_TALLY_DASHBOARD_SHEET_NAME = "Dashboard";
var SIGNAL_TALLY_CHANGELOG_SHEET_NAME = "Changelog";
var SIGNAL_TALLY_PITY_CHECKER_SHEET_NAME = "Pity Checker";

// Optional sheets
var SIGNAL_TALLY_EVENTS_SHEET_NAME = "Events";
var SIGNAL_TALLY_AGENTS_SHEET_NAME = "Agents";
var SIGNAL_TALLY_BANGBOOS_SHEET_NAME = "Bangboos";
var SIGNAL_TALLY_W_ENGINES_SHEET_NAME = "W-Engines";
var SIGNAL_TALLY_RESULTS_SHEET_NAME = "Results";
// Must match optional sheets names
var SETTINGS_FOR_OPTIONAL_SHEET = {
  "Events": {"setting_option": "B14"},
  "Results": {"setting_option": "B15"},
  "Agents": {"setting_option": "B16"},
  "W-Engines": {"setting_option": "B13"},
  "Bangboos": {"setting_option": "B12"},
}

var SIGNAL_TALLY_README_SHEET_NAME = "README";
var SIGNAL_TALLY_AVAILABLE_SHEET_NAME = "Available";
var SIGNAL_TALLY_MONOCHROME_CALCULATOR_SHEET_NAME = "Monochrome Calculator";
var SIGNAL_TALLY_ALL_SIGNAL_HISTORY_SHEET_NAME = "All Signal History";
var SIGNAL_TALLY_ITEMS_SHEET_NAME = "Items";
var SIGNAL_TALLY_NAME_OF_SIGNAL_HISTORY = [
  SIGNAL_TALLY_EXCLUSIVE_CHANNEL_SIGNAL_SHEET_NAME,
  SIGNAL_TALLY_STABLE_SIGNAL_SHEET_NAME,
  SIGNAL_TALLY_W_ENGINES_SIGNAL_SHEET_NAME,
  SIGNAL_TALLY_BANGBOO_SIGNAL_SHEET_NAME,
];

// Import Const
var IMPORT_STATUS_COMPLETE = "COMPLETE";
var IMPORT_STATUS_FAILED = "FAILED";
var IMPORT_STATUS_WISH_HISTORY_COMPLETE = "DONE";
var IMPORT_STATUS_WISH_HISTORY_NOT_FOUND = "NOT FOUND";
var IMPORT_STATUS_WISH_HISTORY_EMPTY = "EMPTY";

// Scheduler Const
var SCHEDULER_TRIGGER_ON_TEXT = "ON";
var SCHEDULER_TRIGGER_OFF_TEXT = "OFF";
var SCHEDULER_RUN_AT_WHICH_DAY = {
  "Monday": ScriptApp.WeekDay.MONDAY,
  "Tuesday": ScriptApp.WeekDay.TUESDAY,
  "Wednesday": ScriptApp.WeekDay.WEDNESDAY,
  "Thursday": ScriptApp.WeekDay.THURSDAY,
  "Friday": ScriptApp.WeekDay.FRIDAY,
  "Saturday": ScriptApp.WeekDay.SATURDAY,
  "Sunday": ScriptApp.WeekDay.SUNDAY
};
var SCHEDULER_RUN_AT_HOUR = {
  "Run at 1:00": 1,
  "Run at 2:00": 2,
  "Run at 3:00": 3,
  "Run at 4:00": 4,
  "Run at 5:00": 5,
  "Run at 6:00": 6,
  "Run at 7:00": 7,
  "Run at 8:00": 8,
  "Run at 9:00": 9,
  "Run at 10:00": 10,
  "Run at 11:00": 11,
  "Run at 12:00": 12,
  "Run at 13:00": 13,
  "Run at 14:00": 14,
  "Run at 15:00": 15,
  "Run at 16:00": 16,
  "Run at 17:00": 17,
  "Run at 18:00": 18,
  "Run at 19:00": 19,
  "Run at 20:00": 20,
  "Run at 21:00": 21,
  "Run at 22:00": 22,
  "Run at 23:00": 23,
  "Run at Midnight": 0
};
var SCHEDULER_RUN_AT_EVERY_HOUR = {
  "Every hour": 1,
  "Every 2 hours": 2,
  "Every 3 hours": 3,
  "Every 4 hours": 4,
  "Every 5 hours": 5,
  "Every 6 hours": 6,
  "Every 7 hours": 7,
  "Every 8 hours": 8,
  "Every 9 hours": 9,
  "Every 10 hours": 10,
  "Every 11 hours": 11,
  "Every 12 hours": 12,
  "Every 13 hours": 13,
  "Every 14 hours": 14,
  "Every 15 hours": 15,
  "Every 16 hours": 16,
  "Every 17 hours": 17,
  "Every 18 hours": 18,
  "Every 19 hours": 19,
  "Every 20 hours": 20,
  "Every 21 hours": 21,
  "Every 22 hours": 22,
  "Every 23 hours": 23,
  "Every 24 hours": 24
};

const CACHED_AUTHKEY_TIMEOUT = 1000 * 60 * 60 * 24;
