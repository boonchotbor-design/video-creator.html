// ============================================================
// config.gs — เวอร์ชั่นกู้คืนระบบ (Resilient Config)
// ============================================================

const CONFIG = {

  // --- Google Sheets ---
  SHEET_ID: '15tjRtRJRx7owdZf6yEwRUiOIQWUnCkt0vd7BHjWR2i4',
  SHEET_NAME: 'Control',

  // --- Gemini API ---
  // 🔒 ดึงคีย์จาก Script Properties (Apps Script > Project Settings > Script Properties)
  //    เพิ่ม: GEMINI_API_KEY = <คีย์ใหม่ของคุณ>
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || 'ใส่คีย์ที่นี่-หรือ-ตั้ง-Script-Properties',
  GEMINI_MODEL: 'gemini-1.5-flash',

  // --- LINE Messaging API ---
  // 🔒 ดึงโทเคนจาก Script Properties: LINE_CHANNEL_TOKEN, LINE_USER_ID
  LINE_CHANNEL_TOKEN: PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_TOKEN') || '',
  LINE_USER_ID: PropertiesService.getScriptProperties().getProperty('LINE_USER_ID') || '',

  // --- Facebook Graph API ---
  FB_PAGE_ID: '102412824529229',
  // สำคัญ: ดึงรหัสผ่าน Page Token จากระบบหลังบ้านของ Google โดยตรง
  FB_PAGE_TOKEN: PropertiesService.getScriptProperties().getProperty('FB_PAGE_TOKEN'),
  FB_API_VERSION: 'v21.0',

  // --- คอลัมน์ใน Google Sheets ---
  COL: {
    PRODUCT_ID:      1,
    PRODUCT_NAME:    2,
    KEY_FEATURES:    3,
    MEDIA_URL:       4,
    MEDIA_TYPE:      5,
    AFFILIATE_LINK:  6,
    TARGET_PLATFORM: 7,
    SCHEDULED_DATE:  8,
    EXTRA_NOTES:     9,
    AI_CAPTION:      10,
    AI_GENERATED_AT: 11,
    STATUS:          12,
    LINE_MSG_ID:     13,
    FB_POST_ID:      14,
    POSTED_AT:       16,
    ERROR_LOG:       17
  },

  STATUS: {
    PENDING:        'PENDING',
    AI_DONE:        'AI_DONE',
    PENDING_REVIEW: 'PENDING_REVIEW',
    APPROVED:       'APPROVED',
    REJECTED:       'REJECTED',
    POSTED:         'POSTED',
    ERROR:          'ERROR'
  }
};

function getSheet() {
  return SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
}

function logError(row, message) {
  const sheet = getSheet();
  sheet.getRange(row, CONFIG.COL.STATUS).setValue(CONFIG.STATUS.ERROR);
  sheet.getRange(row, CONFIG.COL.ERROR_LOG).setValue(message);
  console.error(`Row ${row}: ${message}`);
}
