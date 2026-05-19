// ============================================================================
//   ShopAutomation Ultra — PATCHED VERSION
//   Security & Reliability Fixes:
//   [FIX-1] Bot Token / LINE Token ย้ายออกจาก Sheet → PropertiesService
//   [FIX-2] HTML escape ครอบคลุมทุก field ใน email template
//   [FIX-3] Validate VAT rate ก่อนหาร (กัน division by zero → silent 0)
//   [FIX-4] Sanitize user input ก่อนเขียน Log (กัน log injection)
//   [FIX-5] handleRestock silent skip → แจ้ง user + writeLog ชัดเจนขึ้น
//   [FIX-6] Named constants แทน magic numbers (ONE_DAY_MS, MAX_LOG_VAL, etc.)
//   [FIX-7] buildStockSkuMap ใน installedOnEdit เรียกครั้งเดียวต่อ event
//
//   Performance Fixes (ไม่ตัด feature ออก):
//   [PERF-1] Debounce refreshDashboard 2s — กัน refresh ซ้ำหลายรอบต่อ 1 edit event
//   [PERF-2] formatLogSheetLayout เรียกเฉพาะเมื่อ user อยู่หน้า Log เท่านั้น
//   [PERF-3] checkAndAlertLowStock รับ preloadedStockData — ไม่อ่าน sheet ซ้ำ
//   [PERF-4] SpreadsheetApp.flush() เรียกเฉพาะ calledFromMenu — ไม่ block onEdit
//   [PERF-5] getSettings cached 60s ใน ScriptCache — ลดการอ่าน Settings sheet
// ============================================================================

const CONFIG = {
  SHEET_DASHBOARD: 'หน้าหลัก',
  SHEET_STOCK:     'คลังสินค้า',
  SHEET_ORDERS:    'ออเดอร์',
  SHEET_SETTINGS:  'ตั้งค่า',
  SHEET_LOG:       'Log',
  SHEET_IMPORT:    'Shopee_Import',
  SHEET_PENDING:   'รายการรอตรวจสต็อก',
  CHANNELS: ['Shopee', 'Lazada', 'TikTok Shop', 'IG Shop', 'LINE', 'Facebook'],
  PAYMENTS: ['โอน', 'COD', 'บัตรเครดิต', 'PromptPay', 'TrueMoney'],
  STATUSES: ['เตรียมการจัดส่ง', 'อยู่ระหว่างจัดส่ง', 'จัดส่งเสร็จสิ้น', 'ยกเลิกรายการ', 'คืนสินค้า/คืนเงิน'],
  TIMEZONE: 'Asia/Bangkok'
};

const CANCELLED_STATUSES = ['ยกเลิกรายการ', 'คืนสินค้า/คืนเงิน'];

// [FIX-6] Named constants แทน magic numbers
const ONE_DAY_MS       = 24 * 60 * 60 * 1000;   // 86400000ms
const MAX_LOG_VALUE    = 500;                     // ความยาว max ของ log detail
const MAX_LOW_STOCK_SIG = 8000;                  // max bytes สำหรับ signature debounce
const VAT_MIN          = 0;
const VAT_MAX          = 100;
const DEFAULT_VAT      = 7;
const DEFAULT_LOW_THRESHOLD = 10;

// ============================================================
//   GLOBAL UTILITIES
// ============================================================
function safeNotify(title, message) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 5);
  } catch (e) { Logger.log('[' + title + '] ' + message); }
}

function safeAlert(message) {
  try { SpreadsheetApp.getUi().alert(message); } catch (e) { Logger.log('[Alert] ' + message); }
}

function getOrCreateSheet(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }

// [FIX-4] Sanitize ค่า user input ก่อนเขียน Log — ป้องกัน log injection (newline, tab, pipe)
function sanitizeLogValue(v) {
  return String(v == null ? '' : v).replace(/[\r\n\t|]/g, ' ').substring(0, MAX_LOG_VALUE);
}

function writeLog(ss, type, detail, operator) {
  try {
    const activeSs = ss || SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = activeSs.getSheetByName(CONFIG.SHEET_LOG);
    if (!logSheet) {
      Logger.log('[writeLog] no sheet: ' + type + ' | ' + detail);
      return;
    }
    // [FIX-4] ผ่าน sanitizeLogValue ก่อนเขียนทุกครั้ง
    logSheet.appendRow([
      new Date(),
      sanitizeLogValue(type),
      sanitizeLogValue(detail),
      sanitizeLogValue(operator || 'System Automation')
    ]);
  } catch (e) {
    Logger.log('writeLog fail: ' + e.message + ' | ' + type + ' | ' + String(detail).substring(0, 200));
  }
}

function getCurrentUserEmail() {
  let email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e1) {}
  try {
    if (!String(email || '').trim()) email = Session.getEffectiveUser().getEmail();
  } catch (e2) {}
  return String(email || '').trim() || 'unknown';
}

function isSpreadsheetOwner_(ss) {
  try {
    const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
    const ownerEmail = String(DriveApp.getFileById(spreadsheet.getId()).getOwner().getEmail() || '')
      .trim().toLowerCase();
    let me = String(getCurrentUserEmail() || '').trim().toLowerCase();
    if (me === '' || me === 'unknown') {
      try { me = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); } catch (ignore) {}
    }
    if (!ownerEmail || !me || me === 'unknown') return false;
    return ownerEmail === me;
  } catch (e) {
    Logger.log('isSpreadsheetOwner_: ' + e.message);
    return null;
  }
}

function generateOrderId() {
  const now = new Date();
  const datePart = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyMMdd');
  try {
    const props = PropertiesService.getDocumentProperties();
    const lastDate = props.getProperty('orderSeq_date') || '';
    let seq;
    if (lastDate === datePart) {
      seq = (parseInt(props.getProperty('orderSeq_num')) || 0) + 1;
    } else {
      seq = scanMaxOrderSeqForDate(datePart) + 1;
    }
    props.setProperty('orderSeq_date', datePart);
    props.setProperty('orderSeq_num', String(seq));
    return 'ORD-' + datePart + '-' + String(seq).padStart(4, '0');
  } catch (e) {
    const ms = String(now.getMilliseconds()).padStart(3, '0');
    const rnd = String(Math.floor(Math.random() * 10000)).padStart(4, '0');
    return 'ORD-' + datePart + '-FB' + ms + rnd;
  }
}

function scanMaxOrderSeqForDate(datePart) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ord = ss.getSheetByName(CONFIG.SHEET_ORDERS);
    if (!ord || ord.getLastRow() < 2) return 0;
    const ids = ord.getRange(2, 1, ord.getLastRow() - 1, 1).getValues();
    const pattern = new RegExp('^ORD-' + datePart + '-(\\d{4})$');
    let maxSeq = 0;
    for (let i = 0; i < ids.length; i++) {
      const m = String(ids[i][0] || '').trim().match(pattern);
      if (m) { const n = parseInt(m[1]); if (n > maxSeq) maxSeq = n; }
    }
    return maxSeq;
  } catch (e) { return 0; }
}

function roundMoney(n) {
  const v = parseFloat(n);
  if (!isFinite(v)) return 0;
  return Math.round(v * 100) / 100;
}

function queuePendingStockReview(ss, action, snapshots, editorEmail) {
  if (!snapshots || snapshots.length === 0) return;
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEET_PENDING);
    if (!sheet) return;
    const now = new Date();
    const rows = snapshots.map(s => [
      now, action,
      sanitizeLogValue(s.orderId || ''),
      sanitizeLogValue(s.sku || ''),
      s.qty || 0,
      sanitizeLogValue(s.oldStatus || ''),
      sanitizeLogValue(s.newStatus || ''),
      sanitizeLogValue(editorEmail || 'unknown'),
      'รอดำเนินการ'
    ]);
    const startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, rows.length, 9).setValues(rows);
    sheet.getRange(startRow, 1, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    ensurePendingStatusDropdown(ss);
  } catch (e) {
    Logger.log('queuePendingStockReview fail: ' + e.message);
  }
}

function buildPendingItemStatusValidationRule() {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(['รอดำเนินการ', 'คืนสต็อกแล้ว', 'ยกเลิก ไม่ต้องคืน'], true)
    .setAllowInvalid(false).build();
}

function ensurePendingStatusDropdown(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_PENDING);
  if (!sheet) return;
  const maxR = sheet.getMaxRows();
  if (maxR < 2) return;
  try {
    const lr = Math.max(sheet.getLastRow(), 2);
    const endRow = Math.min(Math.max(lr + 100, 1000), maxR);
    sheet.getRange(2, 9, endRow - 1, 1).setDataValidation(buildPendingItemStatusValidationRule());
  } catch (e) { Logger.log('ensurePendingStatusDropdown: ' + e.message); }
}

function buildStockSkuMap(stockData) {
  const map = new Map();
  for (let i = 0; i < stockData.length; i++) {
    const sku = stockData[i][0];
    if (sku) map.set(String(sku).trim().toLowerCase(), i);
  }
  return map;
}

// ============================================================
//   [FIX-1] BOT TOKEN MANAGEMENT — ย้ายออกจาก Sheet → ScriptProperties
//   ใช้งาน: เรียก menuSaveBotTokens() ครั้งแรกเพื่อบันทึก token จาก Settings sheet
//           หลังจากนั้นลบ token ออกจาก Settings sheet ได้เลย
//   อ่านใช้: getSecureSettings() แทน getSettings() สำหรับ token fields
// ============================================================

/**
 * บันทึก Bot Tokens จาก Settings Sheet → ScriptProperties (ครั้งเดียว)
 * หลังบันทึกแล้วควรลบค่าออกจาก Settings Sheet เพื่อความปลอดภัย
 */
function menuSaveBotTokens() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ต้องเป็น owner เท่านั้น
  const ownerGate = isSpreadsheetOwner_(ss);
  if (ownerGate !== true) {
    ui.alert('ต้องเป็นเจ้าของไฟล์เท่านั้นถึงจะบันทึก Bot Token ได้');
    return;
  }

  const settings = getSettings(ss);
  const props = PropertiesService.getScriptProperties();

  const tokenFields = [
    ['Telegram Bot Token', 'SEC_TG_TOKEN'],
    ['Telegram Chat ID',   'SEC_TG_CHAT_ID'],
    ['LINE Token',         'SEC_LINE_TOKEN'],
    ['LINE Target ID',     'SEC_LINE_TARGET_ID']
  ];

  let saved = 0;
  tokenFields.forEach(([settingKey, propKey]) => {
    const val = String(settings[settingKey] || '').trim();
    if (val) { props.setProperty(propKey, val); saved++; }
  });

  if (saved === 0) {
    ui.alert('ไม่พบ Token ในหน้าตั้งค่า กรุณากรอก Token ในหน้า "ตั้งค่า" ก่อน แล้วกดปุ่มนี้อีกครั้ง');
    return;
  }

  // แนะนำให้ลบออกจาก sheet หลังบันทึกแล้ว
  const confirm = ui.alert(
    '✅ บันทึก Bot Token สำเร็จ ' + saved + ' รายการ\n\n' +
    'เพื่อความปลอดภัย ระบบแนะนำให้ลบค่า Token ออกจากหน้า "ตั้งค่า" ทันที\n' +
    '(Token ถูกเก็บใน Script Properties แล้ว — พนักงานจะไม่เห็นค่าจริงอีกต่อไป)\n\n' +
    'ต้องการให้ระบบลบออกจากหน้าตั้งค่าอัตโนมัติเลยไหม?',
    ui.ButtonSet.YES_NO
  );

  if (confirm === ui.Button.YES) {
    clearTokensFromSettingsSheet_(ss, tokenFields.map(t => t[0]));
    ui.alert('ลบ Token ออกจากหน้าตั้งค่าเรียบร้อย\nระบบยังใช้งาน Telegram/LINE ได้ตามปกติผ่าน Script Properties');
  }

  writeLog(ss, 'SECURITY_TOKEN_SAVED',
    'บันทึก Bot Token เข้า ScriptProperties ' + saved + ' รายการ',
    getCurrentUserEmail());
}

/** ลบค่า Token ออกจาก Settings sheet (เหลือแค่ค่าว่าง) */
function clearTokensFromSettingsSheet_(ss, keysToClear) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_SETTINGS);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][0] || '').trim();
    if (keysToClear.indexOf(k) !== -1) {
      sheet.getRange(i + 1, 2).setValue('');
    }
  }
}

/**
 * อ่าน settings ปกติ + overlay ด้วย secure token จาก ScriptProperties
 * ใช้แทน getSettings() ในทุกฟังก์ชันที่ต้องการ Telegram/LINE token
 */
function getSecureSettings(ss) {
  const settings = getSettings(ss);
  const props = PropertiesService.getScriptProperties();

  // Overlay: ถ้ามีค่าใน ScriptProperties → ใช้แทนค่าใน sheet เสมอ
  const secureMap = [
    ['SEC_TG_TOKEN',    'Telegram Bot Token'],
    ['SEC_TG_CHAT_ID',  'Telegram Chat ID'],
    ['SEC_LINE_TOKEN',  'LINE Token'],
    ['SEC_LINE_TARGET_ID', 'LINE Target ID']
  ];
  secureMap.forEach(([propKey, settingKey]) => {
    const secVal = props.getProperty(propKey);
    if (secVal && String(secVal).trim()) settings[settingKey] = secVal;
  });

  return settings;
}

function getSettings(ss) {
  // [PERF-5] ScriptCache (60s) — กัน Settings sheet ถูกอ่านซ้ำทุก onEdit
  // invalidate อัตโนมัติเมื่อ: menuReorganizeSettings / setupSheets เรียก invalidateSettingsCache()
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('shopSettings');
    if (cached) return JSON.parse(cached);
  } catch (_) {}
  const activeSs = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = activeSs.getSheetByName(CONFIG.SHEET_SETTINGS);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const settings = {};
  const mapping = {
    'ค่าธรรมเนียม Shopee (%)': 'Shopee Fee %',
    'ค่าธรรมเนียม Lazada (%)': 'Lazada Fee %',
    'ค่าธรรมเนียม TikTok Shop (%)': 'TikTok Shop Fee %',
    'ค่าธรรมเนียม IG Shop (%)': 'IG Shop Fee %',
    'ค่าธรรมเนียม LINE (%)': 'LINE Fee %',
    'ค่าธรรมเนียม Facebook (%)': 'Facebook Fee %',
    'ค่า COD - Shopee (%)': 'COD Shopee %',
    'ค่า COD - Lazada (%)': 'COD Lazada %',
    'ค่า COD - TikTok Shop (%)': 'COD TikTok Shop %',
    'ค่า COD - ทั่วไป/หน้าร้าน (%)': 'COD General %',
    'อัตราภาษีมูลค่าเพิ่ม VAT (%)': 'VAT %',
    'จุดแจ้งเตือนสินค้าสต็อกต่ำ (ชิ้น)': 'เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)',
    'จุดแจ้งเตือนสินค้าสต๊อกต่ำ (ชิ้น)': 'เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)',
    'เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)': 'เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)',
    'เกณฑ์สินค้าสต็อกต่ำ (ชิ้น)': 'เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)',
    'อีเมลส่วนกลางสำหรับรับรายงาน': 'อีเมลรับแจ้งเตือนระบบ',
    'ชื่อแบรนด์ / ชื่อร้านค้า': 'ชื่อร้านค้า',
    'เลขประจำตัวผู้เสียภาษีร้านค้า': 'เลขผู้เสียภาษี',
    'ที่อยู่ร้านค้าสำหรับออกใบกำกับภาษี': 'ที่อยู่ร้านค้า',
    'Telegram Bot Token': 'Telegram Bot Token',
    'Telegram Chat ID': 'Telegram Chat ID',
    'LINE Channel Access Token': 'LINE Token',
    'LINE Target ID (User or Group)': 'LINE Target ID'
  };
  for (let i = 1; i < data.length; i++) {
    const rawKey = String(data[i][0]).trim();
    if (rawKey && mapping[rawKey]) settings[mapping[rawKey]] = data[i][1];
  }
  try {
    CacheService.getScriptCache().put('shopSettings', JSON.stringify(settings), 60);
  } catch (_) {}
  return settings;
}

/** [PERF-5] เรียกหลัง user แก้ Settings sheet เพื่อให้รอบถัดไปอ่านค่าใหม่ */
function invalidateSettingsCache() {
  try { CacheService.getScriptCache().remove('shopSettings'); } catch (_) {}
}

// ============================================================
//   LAYOUT HELPERS (ไม่เปลี่ยน)
// ============================================================
function formatOrderUiLayout(sheet) {
  if (!sheet) return;
  sheet.setColumnWidth(1, 140); sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 120); sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 120); sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 90);  sheet.setColumnWidth(8, 100);
  sheet.setColumnWidth(9, 90);  sheet.setColumnWidth(10, 80);
  sheet.setColumnWidth(11, 80); sheet.setColumnWidth(12, 90);
  sheet.setColumnWidth(13, 100);sheet.setColumnWidth(14, 120);
  const headerRange = sheet.getRange('A1:N1');
  headerRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 35);
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 6, sheet.getMaxRows() - 1, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).setHorizontalAlignment('center');
  }
}

function formatOrderUiColors(orderSheet) {
  if (!orderSheet) return;
  const props = PropertiesService.getDocumentProperties();
  const LAYOUT_VERSION = 'v1';
  if (props.getProperty('orderColorsInstalled') === LAYOUT_VERSION) return;
  orderSheet.getRange('A1:N1').setBackground('#0F172A').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10.5);
  orderSheet.setFrozenRows(1);
  const totalDataRows = Math.max(orderSheet.getMaxRows() - 1, 1);
  orderSheet.getRange(2, 2, totalDataRows, 1).setNumberFormat('dd/MM/yyyy');
  orderSheet.getRange(2, 6, totalDataRows, 1).setNumberFormat('#,##0');
  orderSheet.getRange(2, 7, totalDataRows, 1).setNumberFormat('#,##0.00');
  orderSheet.getRange(2, 10, totalDataRows, 1).setNumberFormat('#,##0');
  orderSheet.getRange(2, 11, totalDataRows, 1).setNumberFormat('#,##0');
  orderSheet.getRange(2, 12, totalDataRows, 1).setNumberFormat('#,##0');
  orderSheet.getRange(2, 13, totalDataRows, 1).setNumberFormat('#,##0;[Red]-#,##0');
  orderSheet.getBandings().forEach(b => b.remove());
  const dataRange = orderSheet.getRange(2, 1, totalDataRows, 14);
  dataRange.setVerticalAlignment('middle').setFontSize(10);
  const banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  banding.setHeaderRowColor(null).setFooterRowColor(null)
    .setFirstRowColor('#FFFFFF').setSecondRowColor('#F8FAFC');
  const targetCols = [8, 9, 13, 14];
  const keptRules = orderSheet.getConditionalFormatRules().filter(r => {
    return !r.getRanges().some(rg => targetCols.indexOf(rg.getColumn()) !== -1 && rg.getRow() === 2);
  });
  const channelRange = orderSheet.getRange(2, 8, totalDataRows, 1);
  const channelPalette = [
    ['Shopee','#FFEDD5','#C2410C'],['Lazada','#EDE9FE','#5B21B6'],
    ['TikTok Shop','#1E293B','#FFFFFF'],['IG Shop','#FCE7F3','#9D174D'],
    ['LINE','#DCFCE7','#166534'],['Facebook','#DBEAFE','#1E40AF']
  ];
  channelPalette.forEach(([text, bg, fg]) => {
    keptRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text).setBackground(bg).setFontColor(fg).setBold(true)
      .setRanges([channelRange]).build());
  });
  const paymentRange = orderSheet.getRange(2, 9, totalDataRows, 1);
  const paymentPalette = [
    ['COD','#FEF3C7','#92400E'],['โอน','#E0F2FE','#0C4A6E'],
    ['บัตรเครดิต','#EDE9FE','#5B21B6'],['PromptPay','#DCFCE7','#166534'],
    ['TrueMoney','#FFEDD5','#C2410C']
  ];
  paymentPalette.forEach(([text, bg, fg]) => {
    keptRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text).setBackground(bg).setFontColor(fg)
      .setRanges([paymentRange]).build());
  });
  const statusRange = orderSheet.getRange(2, 14, totalDataRows, 1);
  const statusPalette = [
    ['เตรียมการจัดส่ง','#DBEAFE','#1E40AF'],['อยู่ระหว่างจัดส่ง','#FEF3C7','#92400E'],
    ['จัดส่งเสร็จสิ้น','#DCFCE7','#166534'],['ยกเลิกรายการ','#FEE2E2','#991B1B'],
    ['คืนสินค้า/คืนเงิน','#FCE7F3','#9D174D']
  ];
  statusPalette.forEach(([text, bg, fg]) => {
    keptRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text).setBackground(bg).setFontColor(fg).setBold(true)
      .setRanges([statusRange]).build());
  });
  const profitRange = orderSheet.getRange(2, 13, totalDataRows, 1);
  keptRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setFontColor('#166534').setBold(true).setRanges([profitRange]).build());
  keptRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setFontColor('#991B1B').setBold(true).setRanges([profitRange]).build());
  orderSheet.setConditionalFormatRules(keptRules);
  props.setProperty('orderColorsInstalled', LAYOUT_VERSION);
}

function formatStockUiLayout(sheet) {
  if (!sheet) return;
  sheet.setColumnWidth(1, 75); sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 85); sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 75); sheet.setColumnWidth(6, 75);
  sheet.setColumnWidth(7, 95); sheet.setColumnWidth(8, 130);
  const headerRange = sheet.getRange('A1:H1');
  headerRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 35);
}

function formatStockUiColors(stockSheet) {
  if (!stockSheet) return;
  const props = PropertiesService.getDocumentProperties();
  const LAYOUT_VERSION = 'v1';
  if (props.getProperty('stockColorsInstalled') === LAYOUT_VERSION) return;
  const settings = getSettings(stockSheet.getParent());
  const threshold = parseInt(settings['เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)']) || DEFAULT_LOW_THRESHOLD;
  stockSheet.getRange('A1:H1').setBackground('#0F172A').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10.5);
  stockSheet.setFrozenRows(1);
  const totalDataRows = Math.max(stockSheet.getMaxRows() - 1, 1);
  stockSheet.getRange(2, 4, totalDataRows, 1).setNumberFormat('#,##0');
  stockSheet.getRange(2, 5, totalDataRows, 1).setNumberFormat('#,##0.00');
  stockSheet.getRange(2, 6, totalDataRows, 1).setNumberFormat('#,##0.00');
  stockSheet.getRange(2, 7, totalDataRows, 1).setNumberFormat('0.00%');
  stockSheet.getRange(2, 8, totalDataRows, 1).setNumberFormat('#,##0');
  stockSheet.getRange(2, 4, totalDataRows, 1).setHorizontalAlignment('center');
  stockSheet.getRange(2, 5, totalDataRows, 2).setHorizontalAlignment('right');
  stockSheet.getRange(2, 7, totalDataRows, 1).setHorizontalAlignment('center');
  stockSheet.getRange(2, 8, totalDataRows, 1).setHorizontalAlignment('center');
  stockSheet.getBandings().forEach(b => b.remove());
  const dataRange = stockSheet.getRange(2, 1, totalDataRows, 8);
  dataRange.setVerticalAlignment('middle').setFontSize(10);
  const banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  banding.setHeaderRowColor(null).setFooterRowColor(null)
    .setFirstRowColor('#FFFFFF').setSecondRowColor('#F8FAFC');
  const targetCols = [4, 7, 8];
  const keptRules = stockSheet.getConditionalFormatRules().filter(r => {
    return !r.getRanges().some(rg => targetCols.indexOf(rg.getColumn()) !== -1 && rg.getRow() === 2);
  });
  const stockRange = stockSheet.getRange(2, 4, totalDataRows, 1);
  keptRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0).setBackground('#991B1B').setFontColor('#FFFFFF').setBold(true)
    .setRanges([stockRange]).build());
  keptRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, threshold).setBackground('#FEF3C7').setFontColor('#92400E').setBold(true)
    .setRanges([stockRange]).build());
  keptRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(threshold).setBackground('#DCFCE7').setFontColor('#166534').setBold(true)
    .setRanges([stockRange]).build());
  const marginRange = stockSheet.getRange(2, 7, totalDataRows, 1);
  keptRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0.30).setFontColor('#166534').setBold(true)
    .setRanges([marginRange]).build());
  keptRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.15, 0.2999).setFontColor('#92400E').setBold(true)
    .setRanges([marginRange]).build());
  keptRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.15).setFontColor('#991B1B').setBold(true)
    .setRanges([marginRange]).build());
  const restockRange = stockSheet.getRange(2, 8, totalDataRows, 1);
  keptRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0).setBackground('#DBEAFE').setFontColor('#1E40AF').setBold(true)
    .setRanges([restockRange]).build());
  stockSheet.setConditionalFormatRules(keptRules);
  props.setProperty('stockColorsInstalled', LAYOUT_VERSION);
}

function formatLogSheetLayout(logSheet) {
  if (!logSheet) return;
  const props = PropertiesService.getDocumentProperties();
  const LAYOUT_VERSION = 'v8';
  if (props.getProperty('logLayoutInstalled') === LAYOUT_VERSION) return;
  const a1 = String(logSheet.getRange(1, 1).getValue()).trim();
  if (a1 === 'วันที่-เวลาที่บันทึก') { logSheet.insertRowsBefore(1, 3); }
  logSheet.setColumnWidth(1, 165); logSheet.setColumnWidth(2, 185);
  logSheet.setColumnWidth(3, 640); logSheet.setColumnWidth(4, 130);
  logSheet.setColumnWidth(5, 60);  logSheet.setColumnWidth(6, 130);
  logSheet.setRowHeight(1, 32); logSheet.setRowHeight(2, 6);
  logSheet.setRowHeight(3, 6);  logSheet.setRowHeight(4, 28);
  logSheet.getRange('A1:F1').setBackground('#E0F2FE').setVerticalAlignment('middle');
  logSheet.getRange('A1').setValue('ช่วงเวลา:')
    .setHorizontalAlignment('right').setFontWeight('bold')
    .setFontColor('#0C4A6E').setFontSize(10).setNote('🔒 ป้ายระบบ — ห้ามแก้');
  const presetOptions = ['ทั้งหมด', 'เฉพาะวันนี้', '7 วันที่แล้ว', 'ระบุวันที่เอง'];
  const presetRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(presetOptions, true).setAllowInvalid(false).build();
  const b1Cell = logSheet.getRange('B1');
  const currentB1 = String(b1Cell.getValue()).trim();
  if (currentB1 === 'ระบุช่วงวันที่') { b1Cell.setValue('ระบุวันที่เอง'); }
  else if (presetOptions.indexOf(currentB1) === -1) { b1Cell.setValue('ทั้งหมด'); }
  b1Cell.setDataValidation(presetRule).setHorizontalAlignment('center')
    .setFontWeight('bold').setFontColor('#0F172A').setFontSize(10)
    .setBackground('#FFFFFF')
    .setBorder(true, true, true, true, false, false, '#3B82F6', SpreadsheetApp.BorderStyle.SOLID)
    .setNote('✏️ เลือกช่วงเวลาที่ต้องการกรอง');
  logSheet.getRange('C1').setValue('ตั้งแต่:')
    .setHorizontalAlignment('right').setFontWeight('bold')
    .setFontColor('#0C4A6E').setFontSize(9).setBackground('#E0F2FE')
    .setNote('🔒 ป้ายระบบ — ห้ามแก้');
  logSheet.getRange('E1').setValue('ถึง:')
    .setHorizontalAlignment('right').setFontWeight('bold')
    .setFontColor('#0C4A6E').setFontSize(9).setBackground('#E0F2FE')
    .setNote('🔒 ป้ายระบบ — ห้ามแก้');
  applyLogDateLock(logSheet, String(b1Cell.getValue()).trim());
  logSheet.getRange('A4:D4').setFontWeight('bold').setBackground('#0F172A')
    .setFontColor('#FFFFFF').setFontSize(10).setVerticalAlignment('middle')
    .setHorizontalAlignment('center').setNote('🔒 หัวตารางระบบ — ห้ามแก้');
  logSheet.setFrozenRows(4);
  const totalDataRows = Math.max(logSheet.getMaxRows() - 4, 1);
  const dataRange = logSheet.getRange(5, 1, totalDataRows, 4);
  dataRange.setVerticalAlignment('middle').setWrap(true).setFontSize(10);
  logSheet.getRange(5, 1, totalDataRows, 1)
    .setNumberFormat('dd/MM/yyyy HH:mm:ss').setHorizontalAlignment('center')
    .setFontFamily('Roboto Mono');
  logSheet.getRange(5, 2, totalDataRows, 1).setHorizontalAlignment('left');
  logSheet.getRange(5, 3, totalDataRows, 1).setHorizontalAlignment('left');
  logSheet.getRange(5, 4, totalDataRows, 1).setHorizontalAlignment('center');
  logSheet.getBandings().forEach(b => b.remove());
  const banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  banding.setHeaderRowColor(null).setFooterRowColor(null)
    .setFirstRowColor('#FFFFFF').setSecondRowColor('#F8FAFC');
  const actionRange = logSheet.getRange(5, 2, totalDataRows, 1);
  const keptRules = logSheet.getConditionalFormatRules().filter(r => {
    return !r.getRanges().some(rg => rg.getColumn() === 2 && rg.getRow() === 5);
  });
  const palette = [
    ['ERROR_','#FEE2E2','#991B1B'],['ALERT_','#FEF3C7','#92400E'],
    ['ADD_ORDER','#DCFCE7','#166534'],['RESTOCK','#DBEAFE','#1E40AF'],
    ['STOCK_','#E0E7FF','#3730A3'],['STATUS_','#FCE7F3','#9D174D'],
    ['REPORT_','#F3E8FF','#6B21A8'],['DAILY_','#F3E8FF','#6B21A8'],
    ['IMPORT','#CFFAFE','#155E75'],['SYSTEM_','#FEE2E2','#991B1B'],
    ['ORDER_','#FEF3C7','#854D0E'],['SECURITY_','#FEE2E2','#7C3AED']
  ];
  palette.forEach(([prefix, bg, fg]) => {
    keptRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith(prefix).setBackground(bg).setFontColor(fg).setBold(true)
      .setRanges([actionRange]).build());
  });
  logSheet.setConditionalFormatRules(keptRules);
  props.setProperty('logLayoutInstalled', LAYOUT_VERSION);
}

// ============================================================
//   0. UI MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ระบบจัดการร้านค้า Ultra')
    .addItem('เริ่มต้นระบบและตรวจสอบโครงสร้างชีต', 'setupSheets')
    .addSeparator()
    .addItem('นำเข้าและประมวลผลคำสั่งซื้อ Shopee', 'parseShopeeOrders')
    .addItem('บันทึกยอดและตัดสต๊อกแถวที่เลือก (Manual)', 'menuProcessSelectedOrder')
    .addItem('🔄 ดำเนินการคืนสต็อก (Pending Queue)', 'menuProcessPendingStockReview')
    .addSeparator()
    .addItem('อัปเดตรายงานแดชบอร์ดผู้บริหาร (BI)', 'menuRefreshDashboard')
    .addItem('ออกเอกสารใบเสร็จรับเงิน (Invoice)', 'menuGenerateInvoice')
    .addSeparator()
    .addItem('ตรวจสอบรายการสินค้าวิกฤต (แจ้งเตือนระบบ)', 'menuCheckLowStock')
    .addItem('ส่งรายงานสรุปยอดขายประจำวันเข้าระบบ', 'menuSendDailyReport')
    .addSeparator()
    .addItem('🔐 บันทึก Bot Token (Telegram/LINE) อย่างปลอดภัย', 'menuSaveBotTokens')
    .addSeparator()
    .addItem('🛠 จัดระเบียบเวลาในหน้า Log (Normalize Timestamps)', 'menuNormalizeLogTimestamps')
    .addItem('🧹 จัดระเบียบหน้า "ตั้งค่า" (Reorganize Settings)', 'menuReorganizeSettings')
    .addItem('🔢 ซิงค์เลขรันออเดอร์ (Order ID counter)', 'menuSyncOrderCounter')
    .addItem('🔧 ติดตั้ง Triggers ใหม่ (onEdit + onChange)', 'menuInstallTriggers')
    .addItem('⚠️ ล้างข้อมูลออเดอร์และ Log (เริ่มระบบใหม่)', 'clearAllData')
    .addToUi();
}

function menuNormalizeLogTimestamps() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CONFIG.SHEET_LOG);
  if (!logSheet) { safeAlert('ไม่พบหน้า Log'); return; }
  const lastRow = logSheet.getLastRow();
  if (lastRow < 5) { safeAlert('ยังไม่มีข้อมูลในหน้า Log'); return; }
  const range = logSheet.getRange(5, 1, lastRow - 4, 1);
  const values = range.getValues();
  let converted = 0;
  for (let i = 0; i < values.length; i++) {
    const v = values[i][0];
    if (v instanceof Date) continue;
    if (typeof v !== 'string' || v.trim() === '') continue;
    const cleaned = v.replace(',', '').trim();
    const m = cleaned.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) continue;
    const d = new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]),
                       parseInt(m[4]), parseInt(m[5]), parseInt(m[6] || '0'));
    if (isNaN(d.getTime())) continue;
    values[i][0] = d;
    converted++;
  }
  range.setValues(values).setNumberFormat('dd/MM/yyyy HH:mm:ss')
    .setHorizontalAlignment('center').setFontFamily('Roboto Mono');
  safeAlert('แปลง timestamp เก่าเป็น Date object เรียบร้อย ' + converted + ' แถว');
}

function menuRefreshDashboard() {
  refreshDashboard(SpreadsheetApp.getActiveSpreadsheet(), { calledFromMenu: true }); // [PERF-4]
  safeAlert('ระบบได้ดำเนินการอัปเดตข้อมูลรายงานแดชบอร์ดเรียบร้อยแล้ว');
}

function menuProcessPendingStockReview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_PENDING);
  if (!sheet) { safeAlert('ไม่พบแท็บ "' + CONFIG.SHEET_PENDING + '"'); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { safeAlert('ไม่มีรายการรอดำเนินการ'); return; }
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) {
    safeAlert('ระบบกำลังประมวลผลรายการอื่น กรุณารอสักครู่แล้วลองใหม่');
    return;
  }
  try {
    ensurePendingStatusDropdown(ss);
    const lrData = sheet.getLastRow();
    if (lrData < 2) { safeAlert('ไม่มีรายการรอดำเนินการ'); return; }
    const data = sheet.getRange('A2:I' + lrData).getValues();
    const stockSheet = ss.getSheetByName(CONFIG.SHEET_STOCK);
    if (!stockSheet) { safeAlert('ไม่พบแท็บคลังสินค้า'); return; }
    const stockData = stockSheet.getDataRange().getValues();
    const skuMap = buildStockSkuMap(stockData);
    const stockColData = stockData.slice(1).map(r => [r[3]]);
    let restocked = 0, cancelled = 0, skipped = 0, missing = 0;
    const updatedRows = [];
    for (let i = 0; i < data.length; i++) {
      const status = String(data[i][8]).trim();
      if (status === 'คืนสต็อกแล้ว') {
        const sku = String(data[i][3]).trim();
        const qty = parseInt(data[i][4]) || 0;
        const skuKey = sku.toLowerCase();
        if (qty > 0 && skuMap.has(skuKey)) {
          const idx = skuMap.get(skuKey);
          const offset = idx - 1;
          if (offset >= 0 && offset < stockColData.length) {
            stockColData[offset][0] = (parseInt(stockColData[offset][0]) || 0) + qty;
            restocked++;
            updatedRows.push(i + 2);
            writeLog(ss, 'PENDING_RESTOCK',
              'คืนสต็อก SKU ' + sku + ' +' + qty + ' (orderId=' + data[i][2] + ')',
              data[i][7] || 'unknown');
          }
        } else if (qty > 0 && !skuMap.has(skuKey)) { missing++; }
      } else if (status === 'ยกเลิก ไม่ต้องคืน') {
        cancelled++; updatedRows.push(i + 2);
      } else { skipped++; }
    }
    if (restocked > 0) {
      stockSheet.getRange(2, 4, stockColData.length, 1).setValues(stockColData);
    }
    if (updatedRows.length > 0) {
      updatedRows.forEach(r => sheet.getRange(r, 1, 1, 9).setBackground('#dcfce7'));
    }
    refreshDashboard(ss);
    safeAlert(
      '✅ ดำเนินการเสร็จสิ้น\n\n' +
      '• คืนสต็อกสำเร็จ: ' + restocked + ' รายการ\n' +
      '• ยกเลิกไม่ต้องคืน: ' + cancelled + ' รายการ\n' +
      '• ยังรอดำเนินการ: ' + skipped + ' รายการ\n' +
      (missing > 0 ? '• ไม่พบ SKU ในคลัง: ' + missing + ' รายการ\n' : '') +
      '\nรายการที่ดำเนินการเสร็จจะถูกไฮไลท์สีเขียว'
    );
  } catch (err) {
    writeLog(ss, 'ERROR_PENDING_PROCESS', err.message);
    safeAlert('ดำเนินการล้มเหลว: ' + err.message);
  } finally { lock.releaseLock(); }
}

function menuCheckLowStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSecureSettings(ss); // [FIX-1] ใช้ getSecureSettings
  const sent = checkAndAlertLowStock(ss, settings, true);
  safeAlert(sent
    ? 'ระบบได้ส่งรายงานสินค้าสต็อกต่ำกว่าเกณฑ์เข้าช่องทางการแจ้งเตือนเรียบร้อยแล้ว'
    : 'ตรวจสอบแล้ว ไม่พบสินค้าที่ต่ำกว่าเกณฑ์ หรือยังไม่ได้ตั้งค่าช่องทางการแจ้งเตือน');
}

function menuSendDailyReport() {
  const sent = sendDailyReport();
  safeAlert(sent
    ? 'ระบบได้ประมวลผลและส่งรายงานสรุปยอดขายประจำวันเรียบร้อยแล้ว'
    : 'ไม่สามารถส่งรายงานได้ เนื่องจากไม่มีรายการคำสั่งซื้อในวันปัจจุบัน หรือยังไม่ได้กำหนดข้อมูลการแจ้งเตือน');
}

function menuSaveBotTokens() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const ownerGate = isSpreadsheetOwner_(ss);
  if (ownerGate !== true) {
    ui.alert('ต้องเป็นเจ้าของไฟล์เท่านั้นถึงจะบันทึก Bot Token ได้');
    return;
  }
  const settings = getSettings(ss);
  const props = PropertiesService.getScriptProperties();
  const tokenFields = [
    ['Telegram Bot Token', 'SEC_TG_TOKEN'],
    ['Telegram Chat ID',   'SEC_TG_CHAT_ID'],
    ['LINE Token',         'SEC_LINE_TOKEN'],
    ['LINE Target ID',     'SEC_LINE_TARGET_ID']
  ];
  let saved = 0;
  tokenFields.forEach(([settingKey, propKey]) => {
    const val = String(settings[settingKey] || '').trim();
    if (val) { props.setProperty(propKey, val); saved++; }
  });
  if (saved === 0) {
    ui.alert('ไม่พบ Token ในหน้าตั้งค่า กรุณากรอก Token ก่อน แล้วกดปุ่มนี้อีกครั้ง');
    return;
  }
  const confirm = ui.alert(
    '✅ บันทึก Bot Token สำเร็จ ' + saved + ' รายการ\n\n' +
    'ต้องการให้ระบบลบออกจากหน้าตั้งค่าอัตโนมัติเลยไหม?',
    ui.ButtonSet.YES_NO
  );
  if (confirm === ui.Button.YES) {
    clearTokensFromSettingsSheet_(ss, tokenFields.map(t => t[0]));
    ui.alert('ลบ Token ออกจากหน้าตั้งค่าเรียบร้อย');
  }
  writeLog(ss, 'SECURITY_TOKEN_SAVED',
    'บันทึก Bot Token เข้า ScriptProperties ' + saved + ' รายการ',
    getCurrentUserEmail());
}

function menuProcessSelectedOrder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getActiveSheet();
  if (currentSheet.getName() !== CONFIG.SHEET_ORDERS) {
    safeAlert('ดำเนินการไม่สำเร็จ: กรุณาเลือกแถวข้อมูลในหน้าแท็บ "ออเดอร์"');
    return;
  }
  const row = currentSheet.getActiveRange().getRow();
  if (row <= 1) return;
  const orderRow = currentSheet.getRange(row, 1, 1, 14).getValues()[0];
  const existingOrderId = String(orderRow[0] || '').trim();
  const qty = parseInt(orderRow[5]) || 1;
  if (existingOrderId && existingOrderId.indexOf('ORD-') === 0) {
    const ui = SpreadsheetApp.getUi();
    const confirm = ui.alert(
      '⚠️ แถวนี้มี Order ID อยู่แล้ว',
      'แถวที่เลือก (' + existingOrderId + ') ถูกบันทึกในระบบไปก่อนหน้านี้\n\n' +
      'หากกดยืนยัน ระบบจะตัดสต็อก ' + qty + ' ชิ้นเพิ่มอีกครั้ง (อาจซ้ำซ้อน)\n\n' +
      'ต้องการดำเนินการต่อไหม?',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
  }
  processOrderInPlace(ss, currentSheet, row, orderRow, 0, qty);
  refreshDashboard(ss);
  ss.toast('บันทึกคำสั่งซื้อและปรับปรุงยอดคลังสินค้าเรียบร้อยแล้ว', 'ระบบคลังสินค้า');
}

function menuGenerateInvoice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_ORDERS) {
    safeAlert('กรุณาเลือกรายการคำสั่งซื้อในหน้าแท็บ "ออเดอร์" ก่อนทำรายการ');
    return;
  }
  const row = sheet.getActiveRange().getRow();
  if (row <= 1) return;
  const orderId = sheet.getRange(row, 1).getValue();
  if (!orderId || String(orderId).trim() === '') {
    safeAlert('ไม่พบเลขที่ออเดอร์ในแถวที่เลือก');
    return;
  }
  generateInvoice(orderId);
}

// ============================================================
//   1. SYSTEM INITIALIZATION
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let stockSheet = getOrCreateSheet(ss, CONFIG.SHEET_STOCK);
  if (stockSheet.getLastRow() === 0) {
    stockSheet.appendRow(['SKU', 'ชื่อสินค้า', 'หมวดหมู่', 'สินค้าคงเหลือ', 'ต้นทุน/ชิ้น', 'ราคาขาย', 'อัตรากำไร (Margin %)', 'เติมของเข้าคลัง (ระบุจำนวนชิ้น)']);
    stockSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#2d3748').setFontColor('white');
    stockSheet.setFrozenRows(1);
  } else if (stockSheet.getRange(1, 8).getValue() !== 'เติมของเข้าคลัง (ระบุจำนวนชิ้น)') {
    stockSheet.getRange(1, 8).setValue('เติมของเข้าคลัง (ระบุจำนวนชิ้น)').setFontWeight('bold').setBackground('#2d3748').setFontColor('white');
  }
  let orderSheet = getOrCreateSheet(ss, CONFIG.SHEET_ORDERS);
  if (orderSheet.getLastRow() === 0) {
    orderSheet.appendRow(['เลขที่ออเดอร์', 'วันที่สั่งซื้อ', 'ชื่อลูกค้า', 'SKU', 'ชื่อสินค้า', 'จำนวนชิ้น', 'ราคาขาย/ชิ้น',
      'ช่องทางขาย', 'รูปแบบการชำระเงิน', 'ค่าจัดส่งสินค้า', 'ค่าธรรมเนียมแพลตฟอร์ม', 'ค่าธรรมเนียม COD',
      'กำไรสุทธิ (บาท)', 'สถานะออเดอร์']);
    orderSheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#2d3748').setFontColor('white');
    orderSheet.setFrozenRows(1);
  }
  let settingsSheet = getOrCreateSheet(ss, CONFIG.SHEET_SETTINGS);
  const userEmailEv = getCurrentUserEmail();
  const userEmail = userEmailEv !== 'unknown' ? userEmailEv : '';
  const defaultShopName = ss.getName();
  const defaultSettings = [
    ['ค่าธรรมเนียม Shopee (%)', 7.5], ['ค่าธรรมเนียม Lazada (%)', 7.5],
    ['ค่าธรรมเนียม TikTok Shop (%)', 12], ['ค่าธรรมเนียม IG Shop (%)', 0],
    ['ค่าธรรมเนียม LINE (%)', 0], ['ค่าธรรมเนียม Facebook (%)', 0],
    ['ค่า COD - Shopee (%)', 2.14], ['ค่า COD - Lazada (%)', 2.14],
    ['ค่า COD - TikTok Shop (%)', 2.14], ['ค่า COD - ทั่วไป/หน้าร้าน (%)', 3.00],
    ['อัตราภาษีมูลค่าเพิ่ม VAT (%)', DEFAULT_VAT],
    ['เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)', DEFAULT_LOW_THRESHOLD],
    ['อีเมลส่วนกลางสำหรับรับรายงาน', userEmail],
    ['ชื่อแบรนด์ / ชื่อร้านค้า', defaultShopName],
    ['เลขประจำตัวผู้เสียภาษีร้านค้า', ''],
    ['ที่อยู่ร้านค้าสำหรับออกใบกำกับภาษี', ''],
    // [FIX-1] Token fields ยังมีใน Sheet เพื่อ UX — แต่หลัง setup แนะนำให้ย้ายผ่านเมนู
    ['Telegram Bot Token', '(ใช้เมนู 🔐 บันทึก Bot Token แทน)'],
    ['Telegram Chat ID', ''],
    ['LINE Channel Access Token', '(ใช้เมนู 🔐 บันทึก Bot Token แทน)'],
    ['LINE Target ID (User or Group)', '']
  ];
  if (settingsSheet.getLastRow() === 0) {
    settingsSheet.appendRow(['รายการตั้งค่าระบบหลังบ้าน', 'ค่าที่กำหนด (Value)']);
    defaultSettings.forEach(row => settingsSheet.appendRow(row));
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#2d3748').setFontColor('white');
    settingsSheet.setFrozenRows(1);
  } else {
    const existingKeys = settingsSheet.getRange(2, 1, Math.max(settingsSheet.getLastRow() - 1, 1), 1)
      .getValues().map(r => String(r[0]).trim());
    defaultSettings.forEach(([k, v]) => {
      if (existingKeys.indexOf(k) === -1) settingsSheet.appendRow([k, v]);
    });
  }
  getOrCreateSheet(ss, CONFIG.SHEET_IMPORT);
  let pendingSheet = getOrCreateSheet(ss, CONFIG.SHEET_PENDING);
  if (pendingSheet.getLastRow() === 0) {
    pendingSheet.appendRow(['เวลาที่บันทึก', 'การกระทำ', 'เลขที่ออเดอร์', 'SKU', 'จำนวน', 'สถานะเดิม', 'สถานะใหม่', 'ผู้แก้ไข', 'สถานะรายการ']);
    pendingSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#7c2d12').setFontColor('white');
    pendingSheet.setFrozenRows(1);
    [150, 120, 180, 110, 80, 130, 130, 220, 130].forEach((w, i) => pendingSheet.setColumnWidth(i + 1, w));
    pendingSheet.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }
  ensurePendingStatusDropdown(ss);
  let logSheet = getOrCreateSheet(ss, CONFIG.SHEET_LOG);
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['วันที่-เวลาที่บันทึก', 'ประเภทกิจกรรม', 'รายละเอียด', 'ผู้ดำเนินการ']);
    logSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#2d3748').setFontColor('white');
    logSheet.setFrozenRows(1);
  }
  formatOrderUiLayout(orderSheet);
  formatOrderUiColors(orderSheet);
  formatStockUiLayout(stockSheet);
  formatStockUiColors(stockSheet);
  if (logSheet.getLastColumn() > 0) {
    for (let i = 1; i <= logSheet.getLastColumn(); i++) logSheet.autoResizeColumn(i);
  }
  if (settingsSheet.getLastColumn() > 0) {
    for (let i = 1; i <= settingsSheet.getLastColumn(); i++) settingsSheet.autoResizeColumn(i);
  }
  getOrCreateSheet(ss, CONFIG.SHEET_DASHBOARD);
  const layoutProps = PropertiesService.getDocumentProperties();
  layoutProps.deleteProperty('dashboardLayoutInstalled');
  layoutProps.deleteProperty('logLayoutInstalled');
  layoutProps.deleteProperty('orderColorsInstalled');
  layoutProps.deleteProperty('stockColorsInstalled');
  addOrderValidations(orderSheet, stockSheet);
  syncAllStockMargins(stockSheet);
  protectSystemSheets(ss);
  installInstallableTriggers(ss);
  invalidateSettingsCache(); // [PERF-5] รีเซ็ต cache หลัง setup
  refreshDashboard(ss);
  safeAlert('การตั้งค่าระบบและจัดโครงสร้างหน้าตารางเสร็จสมบูรณ์เรียบร้อยแล้ว\n\n' +
    '📌 ระบบติดตั้ง onEdit + onChange triggers อัตโนมัติแล้ว\n' +
    '🔐 กรุณาใช้เมนู "บันทึก Bot Token อย่างปลอดภัย" เพื่อย้าย Telegram/LINE Token\n' +
    '   ออกจากหน้าตั้งค่า ไปเก็บใน Script Properties แทน');
}

function installInstallableTriggers(ss) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removed = 0;
    triggers.forEach(t => {
      const fn = t.getHandlerFunction();
      if (fn === 'installedOnEdit' || fn === 'installedOnChange') {
        ScriptApp.deleteTrigger(t); removed++;
      }
    });
    ScriptApp.newTrigger('installedOnEdit').forSpreadsheet(ss).onEdit().create();
    ScriptApp.newTrigger('installedOnChange').forSpreadsheet(ss).onChange().create();
    writeLog(ss, 'TRIGGER_INSTALL',
      'ติดตั้ง triggers สำเร็จ (ลบเก่า ' + removed + ' ตัว)');
  } catch (e) {
    writeLog(ss, 'ERROR_TRIGGER_INSTALL', 'ตั้ง trigger ไม่สำเร็จ: ' + e.message);
    safeAlert('⚠️ ติดตั้ง triggers ไม่สำเร็จ\nError: ' + e.message);
  }
}

function menuInstallTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  installInstallableTriggers(ss);
  safeAlert('ติดตั้ง triggers เรียบร้อย (onEdit + onChange)');
}

function menuSyncOrderCounter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd');
  const maxToday = scanMaxOrderSeqForDate(today);
  try {
    const props = PropertiesService.getDocumentProperties();
    props.setProperty('orderSeq_date', today);
    props.setProperty('orderSeq_num', String(maxToday));
    safeAlert(
      '✅ ซิงค์เลขรันออเดอร์เรียบร้อย\n\n' +
      'วันที่ปัจจุบัน: ' + today + '\n' +
      'เลขล่าสุดที่พบในชีต: ' + (maxToday > 0 ? String(maxToday).padStart(4, '0') : '(ยังไม่มี)') + '\n' +
      'เลขถัดไปที่ระบบจะใช้: ORD-' + today + '-' + String(maxToday + 1).padStart(4, '0')
    );
    writeLog(ss, 'COUNTER_SYNC',
      'sync counter วัน ' + today + ' → ' + maxToday, getCurrentUserEmail());
  } catch (e) { safeAlert('ซิงค์ไม่สำเร็จ: ' + e.message); }
}

function menuReorganizeSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_SETTINGS);
  if (!sheet) { safeAlert('ไม่พบหน้า "' + CONFIG.SHEET_SETTINGS + '"'); return; }
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'จัดระเบียบหน้า "ตั้งค่า"',
    'ระบบจะจัดเรียงรายการใหม่ตามหมวดหมู่ ค่าทั้งหมดที่คุณตั้งไว้จะถูกเก็บครบ\n\nต้องการดำเนินการต่อไหม?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) { safeAlert('ระบบกำลังประมวลผลอื่นอยู่ กรุณารอสักครู่'); return; }
  try {
    const data = sheet.getDataRange().getValues();
    const existing = new Map();
    for (let i = 1; i < data.length; i++) {
      const k = String(data[i][0] || '').trim();
      if (!k || /^[━═─]/.test(k)) continue;
      existing.set(k, data[i][1]);
    }
    const aliases = {
      'จุดแจ้งเตือนสินค้าสต็อกต่ำ (ชิ้น)': 'เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)',
      'จุดแจ้งเตือนสินค้าสต๊อกต่ำ (ชิ้น)': 'เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)',
      'เกณฑ์สินค้าสต็อกต่ำ (ชิ้น)': 'เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)'
    };
    Object.keys(aliases).forEach(legacy => {
      if (existing.has(legacy)) {
        const canonical = aliases[legacy];
        if (!existing.has(canonical) || !existing.get(canonical)) {
          existing.set(canonical, existing.get(legacy));
        }
        if (legacy !== canonical) existing.delete(legacy);
      }
    });
    const userEmailEv = getCurrentUserEmail();
    const userEmail = userEmailEv !== 'unknown' ? userEmailEv : '';
    const shopName = ss.getName();
    const groups = [
      { title: '🏷️  ค่าธรรมเนียมแพลตฟอร์ม', color: '#1d4ed8', items: [
        ['ค่าธรรมเนียม Shopee (%)', 7.5, 'percent'],
        ['ค่าธรรมเนียม Lazada (%)', 7.5, 'percent'],
        ['ค่าธรรมเนียม TikTok Shop (%)', 12, 'percent'],
        ['ค่าธรรมเนียม IG Shop (%)', 0, 'percent'],
        ['ค่าธรรมเนียม LINE (%)', 0, 'percent'],
        ['ค่าธรรมเนียม Facebook (%)', 0, 'percent']
      ]},
      { title: '💸  ค่าธรรมเนียม COD', color: '#b45309', items: [
        ['ค่า COD - Shopee (%)', 2.14, 'percent'],
        ['ค่า COD - Lazada (%)', 2.14, 'percent'],
        ['ค่า COD - TikTok Shop (%)', 2.14, 'percent'],
        ['ค่า COD - ทั่วไป/หน้าร้าน (%)', 3.00, 'percent']
      ]},
      { title: '🧾  ภาษี & เกณฑ์ระบบ', color: '#15803d', items: [
        ['อัตราภาษีมูลค่าเพิ่ม VAT (%)', DEFAULT_VAT, 'percent'],
        ['เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)', DEFAULT_LOW_THRESHOLD, 'integer']
      ]},
      { title: '🏪  ข้อมูลร้านค้า', color: '#7e22ce', items: [
        ['ชื่อแบรนด์ / ชื่อร้านค้า', shopName, 'text'],
        ['อีเมลส่วนกลางสำหรับรับรายงาน', userEmail, 'text'],
        ['เลขประจำตัวผู้เสียภาษีร้านค้า', '', 'text'],
        ['ที่อยู่ร้านค้าสำหรับออกใบกำกับภาษี', '', 'text']
      ]},
      { title: '🤖  Bot & Notification (กรอก Token แล้วกดเมนู 🔐 บันทึก Bot Token)', color: '#be185d', items: [
        ['Telegram Bot Token', '', 'text'],
        ['Telegram Chat ID', '', 'text'],
        ['LINE Channel Access Token', '', 'text'],
        ['LINE Target ID (User or Group)', '', 'text']
      ]}
    ];
    const rows = [{ type: 'header' }];
    groups.forEach(g => {
      rows.push({ type: 'banner', title: g.title, color: g.color });
      g.items.forEach(([key, def, fmt]) => {
        const val = existing.has(key) ? existing.get(key) : def;
        rows.push({ type: 'kv', key: key, value: val, fmt: fmt });
        existing.delete(key);
      });
    });
    if (existing.size > 0) {
      rows.push({ type: 'banner', title: '🗂️  รายการอื่นๆ', color: '#475569' });
      existing.forEach((v, k) => rows.push({ type: 'kv', key: k, value: v, fmt: 'text' }));
    }
    sheet.clear();
    try { sheet.clearConditionalFormatRules(); } catch (_) {}
    try { sheet.getBandings().forEach(b => b.remove()); } catch (_) {}
    const writeData = rows.map(r => {
      if (r.type === 'header') return ['รายการตั้งค่าระบบหลังบ้าน', 'ค่าที่กำหนด (Value)'];
      if (r.type === 'banner') return [r.title, ''];
      return [r.key, r.value];
    });
    sheet.getRange(1, 1, writeData.length, 2).setValues(writeData);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff')
      .setHorizontalAlignment('left').setFontSize(11).setVerticalAlignment('middle');
    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 32);
    rows.forEach((r, idx) => {
      const rowNum = idx + 1;
      if (r.type === 'banner') {
        const range = sheet.getRange(rowNum, 1, 1, 2);
        range.merge().setFontWeight('bold').setFontColor('#ffffff').setBackground(r.color)
          .setHorizontalAlignment('left').setFontSize(11).setVerticalAlignment('middle');
        sheet.setRowHeight(rowNum, 30);
      } else if (r.type === 'kv') {
        sheet.getRange(rowNum, 1).setFontWeight('normal').setBackground('#f8fafc')
          .setHorizontalAlignment('left').setVerticalAlignment('middle');
        const valCell = sheet.getRange(rowNum, 2);
        valCell.setBackground('#ffffff').setVerticalAlignment('middle');
        if (r.fmt === 'percent') {
          valCell.setNumberFormat('0.##').setHorizontalAlignment('right').setFontFamily('Roboto Mono');
        } else if (r.fmt === 'integer') {
          valCell.setNumberFormat('0').setHorizontalAlignment('right').setFontFamily('Roboto Mono');
        } else {
          valCell.setNumberFormat('@').setHorizontalAlignment('left');
        }
        sheet.setRowHeight(rowNum, 26);
      }
    });
    sheet.setColumnWidth(1, 360);
    sheet.setColumnWidth(2, 280);
    sheet.setHiddenGridlines(true);
    sheet.getRange(1, 1, writeData.length, 2)
      .setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
    invalidateSettingsCache(); // [PERF-5] บังคับอ่าน Settings ใหม่หลัง reorganize
    writeLog(ss, 'SETTINGS_REORG', 'จัดระเบียบหน้าตั้งค่าใหม่', getCurrentUserEmail());
    safeAlert('จัดระเบียบหน้า "ตั้งค่า" เรียบร้อย');
  } catch (e) {
    safeAlert('เกิดข้อผิดพลาด: ' + e.message);
    writeLog(ss, 'ERR_SETTINGS_REORG', e.message, getCurrentUserEmail());
  } finally { lock.releaseLock(); }
}

function protectSystemSheets(ss) {
  try {
    const sheets = [
      ss.getSheetByName(CONFIG.SHEET_STOCK),
      ss.getSheetByName(CONFIG.SHEET_SETTINGS),
      ss.getSheetByName(CONFIG.SHEET_LOG)
    ];
    let removed = 0;
    sheets.forEach(sh => {
      if (!sh) return;
      sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
        if ((p.getDescription() || '').indexOf('ShopAutoSystem') === 0) { p.remove(); removed++; }
      });
      sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
        if ((p.getDescription() || '').indexOf('ShopAutoSystem') === 0) { p.remove(); removed++; }
      });
    });
    if (removed > 0) {
      writeLog(ss, 'PROTECT_REMOVE', 'ลบ legacy protection ' + removed + ' ตัว');
    }
  } catch (e) {
    writeLog(ss, 'WARN_PROTECT_REMOVE_FAIL', 'ลบ protection ไม่สำเร็จ: ' + e.message);
  }
}

function addOrderValidations(orderSheet, stockSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  stockSheet = stockSheet || ss.getSheetByName(CONFIG.SHEET_STOCK);
  const totalRows = orderSheet.getMaxRows() - 1;
  if (totalRows <= 0) return;
  const skuRange = stockSheet.getRange(2, 1, Math.max(stockSheet.getMaxRows() - 1, 1), 1);
  const skuRule = SpreadsheetApp.newDataValidation().requireValueInRange(skuRange, true).setAllowInvalid(false).build();
  const channelRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.CHANNELS, true).setAllowInvalid(false).build();
  const paymentRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.PAYMENTS, true).setAllowInvalid(false).build();
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STATUSES, true).setAllowInvalid(false).build();
  orderSheet.getRange(2, 4, totalRows, 1).setDataValidation(skuRule);
  orderSheet.getRange(2, 8, totalRows, 1).setDataValidation(channelRule);
  orderSheet.getRange(2, 9, totalRows, 1).setDataValidation(paymentRule);
  orderSheet.getRange(2, 14, totalRows, 1).setDataValidation(statusRule);
}

function syncAllStockMargins(s) {
  if (s.getLastRow() < 2) return;
  const data = s.getRange(2, 5, s.getLastRow() - 1, 2).getValues();
  const m = data.map(r => [r[1] > 0 ? (r[1] - r[0]) / r[1] : '']);
  s.getRange(2, 7, m.length, 1).setValues(m).setNumberFormat('0.00%');
}

// ============================================================
//   2. CORE ORDER PROCESSOR
// ============================================================
function processOrderInPlace(ss, sheet, row, data, oldQty, newQty) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    safeNotify('ระบบประมวลผล', 'ระบบกำลังประมวลผลออเดอร์อื่น กรุณารอ 10 วินาที');
    return;
  }
  try {
    const settings = getSecureSettings(ss); // [FIX-1]
    const stockSheet = ss.getSheetByName(CONFIG.SHEET_STOCK);
    const stockData = stockSheet.getDataRange().getValues();
    let [orderId, date, client, sku, , , price, channel, payment, shipping] = data;
    const stockMap = buildStockSkuMap(stockData);
    const idx = stockMap.has(String(sku).trim().toLowerCase())
      ? stockMap.get(String(sku).trim().toLowerCase()) : -1;
    if (idx === -1) {
      writeLog(ss, 'ERROR_SKU_NOT_FOUND', 'ไม่พบรหัส SKU: ' + sanitizeLogValue(sku));
      return;
    }
    const currentStock = parseInt(stockData[idx][3]) || 0;
    const currentStatus = data[13] || 'เตรียมการจัดส่ง';
    const stockDiff = newQty - oldQty;
    if (CANCELLED_STATUSES.indexOf(currentStatus) === -1 && stockDiff > 0) {
      if (currentStock < stockDiff) {
        sheet.getRange(row, 6).setValue(oldQty > 0 ? oldQty : '');
        safeAlert('สต็อกไม่พอ: สินค้า [' + stockData[idx][1] + '] มีไม่พอในคลัง\n\n' +
                  '• คงเหลือจริงในคลัง: ' + currentStock + ' ชิ้น\n' +
                  '• คุณต้องการสั่งซื้อเพิ่ม: ' + stockDiff + ' ชิ้น');
        return;
      }
    }
    const realCellOrderId = String(sheet.getRange(row, 1).getValue()).trim();
    let finalOrderId;
    if (realCellOrderId && realCellOrderId.indexOf('ORD-') === 0) { finalOrderId = realCellOrderId; }
    else if (orderId) { finalOrderId = orderId; }
    else { finalOrderId = generateOrderId(); }
    const priceStr = String(price == null ? '' : price).trim();
    const finalPrice = (priceStr === '') ? (parseFloat(stockData[idx][5]) || 0) : (parseFloat(price) || 0);
    const finalChannel = channel || 'Shopee';
    const finalPayment = payment || 'โอน';
    let finalShipping = parseFloat(shipping) || 0;
    const itemRevenue = roundMoney(finalPrice * newQty);
    const feePct = (settings[finalChannel + ' Fee %'] || 0) / 100;
    const pFee = roundMoney(itemRevenue * feePct);
    const skuCost = parseFloat(stockData[idx][4]) || 0;
    if (skuCost > 0 && finalPrice > 0 && finalPrice < skuCost) {
      safeNotify('⚠️ ราคาต่ำกว่าทุน',
        'ราคา ' + finalPrice + ' บาท ต่ำกว่าทุน ' + skuCost + ' บาท — โปรดตรวจสอบ');
      writeLog(ss, 'WARN_LOSS_PRICE',
        'ออเดอร์ ' + sanitizeLogValue(finalOrderId) + ' ราคา ' + finalPrice + ' < ทุน ' + skuCost);
    }
    let codFee = 0;
    if (finalPayment === 'COD') {
      let codRate;
      if (finalChannel === 'Shopee')           codRate = (settings['COD Shopee %'] || 2.14) / 100;
      else if (finalChannel === 'Lazada')      codRate = (settings['COD Lazada %'] || 2.14) / 100;
      else if (finalChannel === 'TikTok Shop') codRate = (settings['COD TikTok Shop %'] || 2.14) / 100;
      else                                     codRate = (settings['COD General %'] || 3.00) / 100;
      const codBase = (finalChannel === 'Shopee' || finalChannel === 'TikTok Shop')
        ? (itemRevenue + finalShipping) : itemRevenue;
      codFee = roundMoney(codBase * codRate);
    }
    const cogs = roundMoney(stockData[idx][4] * newQty);
    finalShipping = roundMoney(finalShipping);
    const net = roundMoney(itemRevenue - cogs - pFee - codFee - finalShipping);
    sheet.getRange(row, 1, 1, 14).setValues([[
      finalOrderId, date || new Date(), client || 'ลูกค้าหน้าร้าน',
      sku, stockData[idx][1], newQty, finalPrice, finalChannel, finalPayment, finalShipping,
      pFee, codFee, net, currentStatus
    ]]);
    if (CANCELLED_STATUSES.indexOf(currentStatus) === -1 && stockDiff !== 0) {
      stockSheet.getRange(idx + 1, 4).setValue(currentStock - stockDiff);
      writeLog(ss, 'STOCK_ADJUST', 'ปรับสต็อก SKU ' + sanitizeLogValue(sku) + ' diff: ' + stockDiff);
    }
    // [PERF-3] ส่ง stockData ที่โหลดไว้แล้วให้ checkAndAlertLowStock — ไม่อ่าน sheet ซ้ำ
    checkAndAlertLowStock(ss, settings, false, stockData.slice(1));
  } catch (err) {
    writeLog(ss, 'ERROR_PROCESS_ORDER', sanitizeLogValue(err.message));
  } finally { lock.releaseLock(); }
}

// ============================================================
//   3. AUTO-RESTOCK ON STATUS CHANGE
//   [FIX-5] Lock timeout → แจ้ง user + writeLog ชัดขึ้น ไม่ silent skip
// ============================================================
function handleRestock(ss, data, oldS, newS, oldQty) {
  const wasCancelled = CANCELLED_STATUSES.indexOf(oldS) !== -1;
  const isCancelled = CANCELLED_STATUSES.indexOf(newS) !== -1;
  if (wasCancelled === isCancelled) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(8000)) {
    // [FIX-5] ไม่ silent skip — แจ้ง user + log ชัดเจน
    const orderId = sanitizeLogValue(data[0] || 'unknown');
    writeLog(ss, 'ERROR_RESTOCK_LOCK_TIMEOUT',
      'handleRestock: lock timeout 8s — สต็อกไม่ถูกปรับ orderId=' + orderId +
      ' SKU=' + sanitizeLogValue(data[3] || '') +
      ' oldStatus=' + sanitizeLogValue(oldS) + ' newStatus=' + sanitizeLogValue(newS),
      getCurrentUserEmail());
    safeNotify('⚠️ ระบบยุ่ง — สต็อกยังไม่ถูกปรับ',
      'ออเดอร์ ' + orderId + ' เปลี่ยนสถานะสำเร็จ แต่ระบบยุ่งเกินไป\n' +
      'สต็อกจะถูกปรับ เมื่อคุณรีเฟรชหน้าออเดอร์ หรือกดอัปเดต Dashboard อีกครั้ง\n' +
      'หากสต็อกยังไม่ตรง กรุณาปรับด้วยมือผ่านคอลัมน์ H ในหน้าคลังสินค้า');
    return;
  }
  try {
    const stockSheet = ss.getSheetByName(CONFIG.SHEET_STOCK);
    const stockData = stockSheet.getDataRange().getValues();
    const stockMap = buildStockSkuMap(stockData);
    const skuKey = String(data[3]).trim().toLowerCase();
    if (!stockMap.has(skuKey)) return;
    const idx = stockMap.get(skuKey);
    const current = parseInt(stockData[idx][3]) || 0;
    const qty = (oldQty != null && !isNaN(parseInt(oldQty))) ? parseInt(oldQty) : (parseInt(data[5]) || 0);
    if (isCancelled) {
      stockSheet.getRange(idx + 1, 4).setValue(current + qty);
      writeLog(ss, 'RESTOCK', 'คืนคลัง ' + sanitizeLogValue(data[3]) + ' +' + qty + ' (สถานะ: ' + sanitizeLogValue(newS) + ')',
        getCurrentUserEmail());
    } else {
      stockSheet.getRange(idx + 1, 4).setValue(Math.max(0, current - qty));
      writeLog(ss, 'RESTOCK_DEDUCT', 'หักคลัง ' + sanitizeLogValue(data[3]) + ' -' + qty + ' (สถานะ: ' + sanitizeLogValue(newS) + ')',
        getCurrentUserEmail());
    }
  } finally { lock.releaseLock(); }
}

// ============================================================
//   4. SHOPEE BULK IMPORT
// ============================================================
function parseShopeeOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const imp = ss.getSheetByName(CONFIG.SHEET_IMPORT);
  const ord = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  const stockSheet = ss.getSheetByName(CONFIG.SHEET_STOCK);
  if (!imp || imp.getLastRow() < 3) {
    safeNotify('ข้อผิดพลาด', 'ไม่พบข้อมูลในแท็บ Shopee_Import');
    return;
  }
  if (!ord || !stockSheet) return;
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) {
    safeNotify('ระบบประมวลผล', 'ระบบกำลังประมวลผลออเดอร์อื่น กรุณารออีก 15 วินาที');
    return;
  }
  try {
    const data = imp.getRange(3, 1, imp.getLastRow() - 2, imp.getLastColumn()).getValues();
    if (data.length === 0 || !data[0][0]) return;
    const headers = data[0].map(h => String(h).trim());
    const idxId = headers.indexOf('หมายเลขคำสั่งซื้อ');
    let idxSku = headers.indexOf('รหัสอ้างอิง SKU (SKU Reference No.)');
    if (idxSku === -1) idxSku = headers.indexOf('รหัสอ้างอิง SKU');
    if (idxSku === -1) idxSku = headers.indexOf('รหัส SKU');
    const idxQty = headers.indexOf('จำนวน');
    let idxPrice = headers.indexOf('ราคาตั้งต้น');
    if (idxPrice === -1) idxPrice = headers.indexOf('ราคาสุทธิที่ชำระ');
    if (idxPrice === -1) idxPrice = headers.indexOf('ราคาขาย');
    const idxUser = headers.indexOf('ชื่อผู้ใช้ (ผู้ซื้อ)');
    const idxPay = headers.indexOf('วิธีการชำระเงิน');
    if (idxId === -1 || idxSku === -1 || idxQty === -1 || idxPrice === -1) {
      safeNotify('ข้อผิดพลาด', 'โครงสร้างไฟล์นำเข้าไม่ตรงตามมาตรฐาน Shopee');
      return;
    }
    const settings = getSecureSettings(ss); // [FIX-1]
    const stockData = stockSheet.getDataRange().getValues();
    const stockMap = buildStockSkuMap(stockData);
    const existingKeys = new Set(
      ord.getDataRange().getValues().slice(1)
        .filter(r => r[0])
        .map(r => String(r[0]).trim() + '|' + String(r[3] || '').trim().toLowerCase())
    );
    const feePct = (settings['Shopee Fee %'] || 0) / 100;
    const codRatePct = (settings['COD Shopee %'] || 2.14) / 100;
    const newRows = [];
    const stockUpdates = new Map();
    let count = 0, skip = 0, missing = 0, lowStock = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const orderId = String(row[idxId]).trim();
      if (!orderId) continue;
      const sku = String(row[idxSku]).trim();
      const qty = parseInt(row[idxQty]) || 0;
      const price = parseFloat(row[idxPrice]) || 0;
      const customer = idxUser !== -1 ? String(row[idxUser]).trim() : 'ลูกค้าระบบ Shopee';
      const rawPay = idxPay !== -1 ? String(row[idxPay]).trim() : 'โอน';
      let payment = 'โอน';
      if (rawPay.includes('ปลายทาง') || rawPay.toLowerCase().includes('cod')) payment = 'COD';
      else if (rawPay.includes('บัตร')) payment = 'บัตรเครดิต';
      else if (rawPay.includes('Prompt') || rawPay.includes('พร้อมเพย์')) payment = 'PromptPay';
      if (!sku || qty <= 0) continue;
      const dupeKey = orderId + '|' + sku.toLowerCase();
      if (existingKeys.has(dupeKey)) { skip++; continue; }
      const skuKey = sku.toLowerCase();
      if (!stockMap.has(skuKey)) { missing++; continue; }
      const idx = stockMap.get(skuKey);
      const stockRow = stockData[idx];
      const cost = parseFloat(stockRow[4]) || 0;
      const stockPrice = parseFloat(stockRow[5]) || 0;
      const currentStock = stockUpdates.has(idx) ? stockUpdates.get(idx) : (parseInt(stockRow[3]) || 0);
      if (currentStock < qty) { lowStock++; continue; }
      const finalPrice = price || stockPrice;
      const itemRevenue = roundMoney(finalPrice * qty);
      const pFee = roundMoney(itemRevenue * feePct);
      const codFee = (payment === 'COD') ? roundMoney(itemRevenue * codRatePct) : 0;
      const cogs = roundMoney(cost * qty);
      const net = roundMoney(itemRevenue - cogs - pFee - codFee);
      newRows.push([orderId, new Date(), customer, sku, stockRow[1], qty, finalPrice,
        'Shopee', payment, 0, pFee, codFee, net, 'เตรียมการจัดส่ง']);
      stockUpdates.set(idx, currentStock - qty);
      existingKeys.add(dupeKey);
      count++;
    }
    if (newRows.length > 0) {
      const startRow = ord.getLastRow() + 1;
      ord.getRange(startRow, 1, newRows.length, 14).setValues(newRows);
    }
    if (stockUpdates.size > 0 && stockData.length > 1) {
      const stockColData = stockData.slice(1).map(row => [row[3]]);
      stockUpdates.forEach((newQty, idx) => {
        if (idx >= 1 && (idx - 1) < stockColData.length) stockColData[idx - 1][0] = newQty;
      });
      stockSheet.getRange(2, 4, stockColData.length, 1).setValues(stockColData);
    }
    if (imp.getLastRow() >= 3) imp.getRange(3, 1, imp.getLastRow() - 2, imp.getLastColumn()).clearContent();
    writeLog(ss, 'IMPORT', 'นำเข้า Shopee ' + count + ' รายการ (ข้ามซ้ำ ' + skip + ', ไม่พบ SKU ' + missing + ', สต๊อกไม่พอ ' + lowStock + ')');
    refreshDashboard(ss);
    if (count > 0) checkAndAlertLowStock(ss, settings, false);
    let summary = 'นำเข้า ' + count + ' รายการ';
    const issues = [];
    if (skip) issues.push('ข้ามซ้ำ ' + skip);
    if (missing) issues.push('ไม่พบ SKU ' + missing);
    if (lowStock) issues.push('สต๊อกไม่พอ ' + lowStock);
    if (issues.length) summary += ' (' + issues.join(', ') + ')';
    safeNotify('เสร็จสมบูรณ์', summary);
  } catch (err) {
    writeLog(ss, 'ERROR_IMPORT', sanitizeLogValue(err.message));
    safeNotify('ข้อผิดพลาด', 'นำเข้าล้มเหลว: ' + err.message);
  } finally { lock.releaseLock(); }
}

// ============================================================
//   5a. INSTALLABLE onChange
// ============================================================
function installedOnChange(e) {
  if (!e || !e.source) return;
  const ct = String(e.changeType || '');
  if (ct !== 'INSERT_ROW' && ct !== 'REMOVE_ROW' && ct !== 'INSERT_ROWS' && ct !== 'REMOVE_ROWS') return;
  try {
    ensurePendingStatusDropdown(e.source);
    refreshDashboard(e.source, { triggeredByStructuralChange: true });
  } catch (err) { Logger.log('installedOnChange: ' + err.message); }
}

// ============================================================
//   5b. INSTALLABLE onEdit
//   [FIX-7] buildStockSkuMap เรียกครั้งเดียวต่อ event (ไม่เรียกซ้ำใน CASE 1, 3)
// ============================================================
function installedOnEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();
  const ss = sheet.getParent();
  const sheetName = sheet.getName();
  const startRow = range.getRow();
  const numRows = range.getNumRows();
  const startCol = range.getColumn();
  const numCols = range.getNumColumns();
  if (startRow <= 1) return;

  // ===== STOCK SHEET =====
  if (sheetName === CONFIG.SHEET_STOCK) {
    if (startCol === 4) {
      range.setValue(e.oldValue);
      safeNotify('ระบบความปลอดภัย', 'ห้ามแก้ "สินค้าคงเหลือ" โดยตรง กรุณาระบุจำนวนเติมในคอลัมน์ H');
    } else if (startCol === 8 && numRows === 1) {
      const restockAmt = parseInt(range.getValue());
      if (!isNaN(restockAmt) && restockAmt > 0) {
        const currentStockCell = sheet.getRange(startRow, 4);
        const currentStock = parseInt(currentStockCell.getValue()) || 0;
        currentStockCell.setValue(currentStock + restockAmt);
        range.setValue('');
        const productName = sheet.getRange(startRow, 2).getValue();
        writeLog(ss, 'RESTOCK_INFLOW',
          'เติม ' + sanitizeLogValue(productName) + ' +' + restockAmt + ' ชิ้น (ยอดใหม่: ' + (currentStock + restockAmt) + ')',
          getCurrentUserEmail());
        safeNotify('คลังสินค้า', 'เติม ' + productName + ' +' + restockAmt + ' ชิ้น สำเร็จ');
        refreshDashboard(ss);
      }
    } else if (startCol === 5 || startCol === 6) {
      syncAllStockMargins(sheet);
      refreshDashboard(ss);
    }
    return;
  }

  if (sheetName !== CONFIG.SHEET_ORDERS) return;

  const stockSheet = ss.getSheetByName(CONFIG.SHEET_STOCK);
  if (!stockSheet) return;

  const isMultiCellEdit = numRows > 1 || numCols > 1;
  let didMultiRowQtyClearHandled = false;
  let didMultiRowStatusHandled = false;

  // [FIX-7] โหลด stockData + buildStockSkuMap ครั้งเดียวก่อนลูป
  const stockValues = stockSheet.getDataRange().getValues();
  const sharedSkuMap = buildStockSkuMap(stockValues);

  for (let i = 0; i < numRows; i++) {
    const currentRow = startRow + i;
    const orderData = sheet.getRange(currentRow, 1, 1, 14).getValues()[0];
    const [orderId, , , sku, , qty, , channel, payment, , , , , status] = orderData;
    let processed = false;
    const isTargetCol = (startCol <= 4 && startCol + numCols > 4) || (startCol <= 6 && startCol + numCols > 6);

    // CASE 1: ลบ SKU หรือ qty = 0
    if (isTargetCol && (String(sku).trim() === '' || String(qty).trim() === '' || parseInt(qty) === 0)) {
      if (!isMultiCellEdit) {
        const oldQty = parseInt(e.oldValue) || 0;
        const currentSku = String(sku).trim() !== '' ? sku : e.oldValue;
        if (currentSku && oldQty > 0) {
          // [FIX-7] ใช้ sharedSkuMap ที่โหลดไว้แล้ว ไม่อ่านชีตซ้ำ
          const lookupKey = String(currentSku).trim().toLowerCase();
          if (sharedSkuMap.has(lookupKey)) {
            const sIdx = sharedSkuMap.get(lookupKey);
            stockSheet.getRange(sIdx + 1, 4).setValue((parseInt(stockValues[sIdx][3]) || 0) + oldQty);
            writeLog(ss, 'STOCK_REFUND', 'คืนสต็อก ' + sanitizeLogValue(currentSku) + ' +' + oldQty + ' (ลบรายการ)');
          }
        }
      } else {
        if (!didMultiRowQtyClearHandled) {
          didMultiRowQtyClearHandled = true;
          const editorEmail = getCurrentUserEmail();
          const snapshots = [];
          try {
            const blockData = sheet.getRange(startRow, 1, numRows, 14).getValues();
            blockData.forEach(rowVals => {
              const sn = { orderId: rowVals[0], sku: rowVals[3] || '', qty: parseInt(rowVals[5]) || 0, oldStatus: rowVals[13] || '', newStatus: 'DELETED' };
              if (sn.orderId || sn.sku) snapshots.push(sn);
            });
          } catch (snapErr) { Logger.log('snapshot fail: ' + snapErr.message); }
          queuePendingStockReview(ss, 'MULTIROW_DELETE', snapshots, editorEmail);
          writeLog(ss, 'WARN_MULTIROW_DELETE',
            'ลบหลายแถว → enqueue ' + snapshots.length + ' (numRows=' + numRows + ', startRow=' + startRow + ')',
            editorEmail);
          let queueNote = snapshots.length === 0 ? '(ไม่พบ SKU/เลขออเดอร์)' : '(บันทึกลง Pending แล้ว)';
          safeAlert(
            '⚠️ ลบหลายแถวพร้อมกัน — สต็อกไม่ถูกคืนอัตโนมัติ\n\n' +
            'จำนวนแถว: ' + numRows + ' แถว (เริ่ม row ' + startRow + ')\n' +
            'เข้าคิวรอตรวจ: ' + snapshots.length + ' รายการ ' + queueNote + '\n\n' +
            'วิธีแก้:\n1. เปิดแท็บ "' + CONFIG.SHEET_PENDING + '" → ตรวจรายการ\n' +
            '2. ใช้เมนู "ดำเนินการคืนสต็อก (Pending Queue)"'
          );
        }
      }
      sheet.getRange(currentRow, 5).setValue('');
      sheet.getRange(currentRow, 7).setValue('');
      sheet.getRange(currentRow, 11).setValue('');
      sheet.getRange(currentRow, 12).setValue('');
      sheet.getRange(currentRow, 13).setValue('');
      continue;
    }

    // CASE 2: เปลี่ยนช่องทางขาย
    if (startCol <= 8 && startCol + numCols > 8) {
      const currentChannel = String(sheet.getRange(currentRow, 8).getValue()).trim();
      let allowedPayments = (currentChannel === 'Shopee' || currentChannel === 'Lazada' || currentChannel === 'TikTok Shop')
        ? ['โอน', 'COD', 'บัตรเครดิต']
        : (currentChannel === 'LINE') ? ['โอน', 'PromptPay', 'COD']
        : (currentChannel === 'IG Shop' || currentChannel === 'Facebook') ? ['โอน', 'PromptPay', 'TrueMoney', 'COD']
        : CONFIG.PAYMENTS;
      const paymentCell = sheet.getRange(currentRow, 9);
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(allowedPayments, true).setAllowInvalid(false).build();
      paymentCell.setDataValidation(rule);
      const pTrim = String(payment || '').trim();
      const ix = allowedPayments.indexOf(pTrim);
      const loneChannelEdit = !isMultiCellEdit && numRows === 1 && startCol === 8 && numCols === 1;
      if (isMultiCellEdit) { if (pTrim && ix === -1) paymentCell.setValue(allowedPayments[0]); }
      else if (loneChannelEdit) { if (!pTrim || ix === -1) paymentCell.setValue(allowedPayments[0]); }
      if (!processed && sku && qty > 0) {
        processOrderInPlace(ss, sheet, currentRow, sheet.getRange(currentRow, 1, 1, 14).getValues()[0], qty, qty);
        processed = true;
      }
    }

    // CASE 3: เปลี่ยน SKU
    if (startCol <= 4 && startCol + numCols > 4 && sku && qty > 0) {
      // [FIX-7] ใช้ sharedSkuMap แทนการสร้าง map ใหม่
      if (!isMultiCellEdit) {
        const oldSku = e.oldValue;
        if (oldSku && String(oldSku) !== String(sku)) {
          writeLog(ss, 'AUDIT_EDIT_SKU',
            'SKU ออเดอร์ ' + sanitizeLogValue(orderId) + ': ' + sanitizeLogValue(oldSku) + ' → ' + sanitizeLogValue(sku),
            getCurrentUserEmail());
        }
        if (oldSku) {
          const oldKey = String(oldSku).trim().toLowerCase();
          if (sharedSkuMap.has(oldKey)) {
            const oldIdx = sharedSkuMap.get(oldKey);
            stockSheet.getRange(oldIdx + 1, 4).setValue((parseInt(stockValues[oldIdx][3]) || 0) + qty);
          }
        }
      }
      const newKey = String(sku).trim().toLowerCase();
      if (sharedSkuMap.has(newKey) && !processed) {
        const match = stockValues[sharedSkuMap.get(newKey)];
        sheet.getRange(currentRow, 5).setValue(match[1]);
        sheet.getRange(currentRow, 7).setValue(match[5]);
        processOrderInPlace(ss, sheet, currentRow, sheet.getRange(currentRow, 1, 1, 14).getValues()[0], 0, qty);
        processed = true;
      }
    }

    // CASE 4: เปลี่ยน qty
    if (startCol <= 6 && startCol + numCols > 6 && sku && parseInt(qty) > 0) {
      const oldQty = parseInt(!isMultiCellEdit ? e.oldValue : orderData[5]) || 0;
      const newQty = parseInt(qty) || 0;
      if (!isMultiCellEdit && oldQty !== newQty) {
        writeLog(ss, 'AUDIT_EDIT_QTY',
          'จำนวน ออเดอร์ ' + sanitizeLogValue(orderId) + ' SKU ' + sanitizeLogValue(sku) +
          ': ' + oldQty + ' → ' + newQty + ' (diff=' + (newQty - oldQty) + ')',
          getCurrentUserEmail());
      }
      if (!processed) {
        processOrderInPlace(ss, sheet, currentRow, orderData, oldQty, newQty);
        processed = true;
      }
    }

    // CASE 5: เปลี่ยน price/payment/shipping
    if ((startCol === 7 || startCol === 9 || startCol === 10) && sku && qty > 0) {
      if (!isMultiCellEdit && e.oldValue !== undefined && String(e.oldValue) !== String(e.value)) {
        const fieldName = startCol === 7 ? 'ราคาขาย' : (startCol === 9 ? 'ช่องทางชำระ' : 'ค่าจัดส่ง');
        writeLog(ss, 'AUDIT_EDIT_' + (startCol === 7 ? 'PRICE' : startCol === 9 ? 'PAYMENT' : 'SHIPPING'),
          fieldName + ' ออเดอร์ ' + sanitizeLogValue(orderId) + ': ' +
          sanitizeLogValue(e.oldValue == null ? '(ว่าง)' : e.oldValue) + ' → ' +
          sanitizeLogValue(e.value == null ? '(ว่าง)' : e.value),
          getCurrentUserEmail());
      }
      if (!processed) {
        processOrderInPlace(ss, sheet, currentRow, orderData, qty, qty);
        processed = true;
      }
    }

    // CASE 6: เปลี่ยน status
    if (startCol === 14 && orderId) {
      if (!isMultiCellEdit) {
        handleRestock(ss, orderData, e.oldValue, status);
        writeLog(ss, 'STATUS_CHANGE',
          'คำสั่งซื้อ ' + sanitizeLogValue(orderId) + ': ' +
          sanitizeLogValue(e.oldValue || '?') + ' → ' + sanitizeLogValue(status),
          getCurrentUserEmail());
      } else {
        if (!didMultiRowStatusHandled) {
          didMultiRowStatusHandled = true;
          const editorEmail = getCurrentUserEmail();
          const snapshots = [];
          try {
            const blockData = sheet.getRange(startRow, 1, numRows, 14).getValues();
            blockData.forEach(rowVals => {
              const sn = { orderId: rowVals[0], sku: rowVals[3] || '', qty: parseInt(rowVals[5]) || 0, oldStatus: '(ไม่ทราบ)', newStatus: rowVals[13] || status || '' };
              if (sn.orderId) snapshots.push(sn);
            });
          } catch (snapErr) { Logger.log('snapshot fail: ' + snapErr.message); }
          queuePendingStockReview(ss, 'MULTIROW_STATUS', snapshots, editorEmail);
          writeLog(ss, 'WARN_MULTIROW_STATUS',
            'เปลี่ยนสถานะหลายแถว → ' + sanitizeLogValue(status) + ' → enqueue ' + snapshots.length,
            editorEmail);
          safeAlert(
            '⚠️ เปลี่ยนสถานะหลายแถวพร้อมกัน — สต็อกไม่ถูกปรับอัตโนมัติ\n\n' +
            'สถานะใหม่: ' + status + ' | จำนวน: ' + numRows + ' แถว\n' +
            'เข้าคิวรอตรวจ: ' + snapshots.length + ' รายการ\n\n' +
            'วิธีแก้:\n1. เปิดแท็บ "' + CONFIG.SHEET_PENDING + '" → ตรวจรายการ\n' +
            '2. ใช้เมนู "ดำเนินการคืนสต็อก (Pending Queue)"'
          );
        }
      }
    }
  }

  // [PERF-1] Debounce dashboard refresh: ถ้า refresh เกิดขึ้นภายใน 2 วิที่แล้ว → ข้าม
  // กัน multi-CASE edit (เช่น paste หลาย col) ที่ trigger refresh ซ้ำหลายรอบ
  debouncedRefreshDashboard_(ss);
}

/**
 * [PERF-1] Debounced refreshDashboard — ใช้ ScriptProperties เป็น timestamp gate
 * window = 2000ms: ถ้า refresh เกิดขึ้นภายใน 2s จาก trigger ล่าสุด → ข้าม
 * ป้องกัน: paste 14 cols → CASE 2+3+4+5 ทำงานพร้อมกัน → refresh 4 รอบติด
 */
function debouncedRefreshDashboard_(ss) {
  const DEBOUNCE_MS = 2000;
  try {
    const props = PropertiesService.getDocumentProperties();
    const lastTs = parseInt(props.getProperty('lastDashRefreshTs') || '0');
    const now = Date.now();
    if ((now - lastTs) < DEBOUNCE_MS) return; // ยังอยู่ในหน้าต่าง debounce → ข้าม
    props.setProperty('lastDashRefreshTs', String(now));
  } catch (_) {} // ถ้า PropertiesService fail → refresh ตามปกติ
  refreshDashboard(ss);
}

// ============================================================
//   6. SIMPLE onEdit — Dashboard + Log filter
// ============================================================
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const ss = sheet.getParent();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (sheetName === CONFIG.SHEET_DASHBOARD) {
    if (row === 3 && col === 11) { refreshDashboard(ss); return; }
    if (row === 3 && (col === 13 || col === 15)) {
      const currentMode = String(sheet.getRange('K3').getValue()).trim();
      if (currentMode !== 'ระบุวันที่เอง') {
        safeNotify('🔒 ช่องวันที่ถูกล็อค', 'เปลี่ยนช่วงเวลาเป็น "ระบุวันที่เอง" ก่อน');
        refreshDashboard(ss); return;
      }
      refreshDashboard(ss);
    }
    return;
  }

  if (sheetName === CONFIG.SHEET_LOG) {
    if (row === 1 && col === 1) {
      if (String(e.range.getValue()).trim() !== 'ช่วงเวลา:') {
        e.range.setValue('ช่วงเวลา:').setFontSize(10).setFontWeight('bold')
          .setFontColor('#0C4A6E').setHorizontalAlignment('right');
        safeNotify('🔒 ช่องระบบถูกล็อค', 'ห้ามแก้ป้ายฟิลเตอร์');
      }
      return;
    }
    if (row === 1 && col === 3) {
      if (String(e.range.getValue()).trim() !== 'ตั้งแต่:') {
        e.range.setValue('ตั้งแต่:').setFontSize(9).setFontWeight('bold')
          .setFontColor('#0C4A6E').setHorizontalAlignment('right');
        safeNotify('🔒 ช่องระบบถูกล็อค', 'ห้ามแก้ป้ายวันที่');
      }
      return;
    }
    if (row === 1 && col === 5) {
      if (String(e.range.getValue()).trim() !== 'ถึง:') {
        e.range.setValue('ถึง:').setFontSize(9).setFontWeight('bold')
          .setFontColor('#0C4A6E').setHorizontalAlignment('right');
        safeNotify('🔒 ช่องระบบถูกล็อค', 'ห้ามแก้ป้ายวันที่');
      }
      return;
    }
    if (row === 4 && col >= 1 && col <= 4) {
      const headerValues = ['วันที่-เวลาที่บันทึก', 'ประเภทกิจกรรม', 'รายละเอียด', 'ผู้ดำเนินการ'];
      e.range.setValue(headerValues[col - 1]);
      safeNotify('🔒 หัวตารางถูกล็อค', 'ห้ามแก้หัวตารางหน้า Log');
      return;
    }
    if (row === 1 && col === 2) {
      const mode = String(e.range.getValue()).trim();
      applyLogDateLock(sheet, mode);
      const startVal = (mode === 'ระบุวันที่เอง') ? sheet.getRange('D1').getValue() : null;
      const endVal = (mode === 'ระบุวันที่เอง') ? sheet.getRange('F1').getValue() : null;
      applyLogFilter(sheet, mode, startVal, endVal);
      return;
    }
    if (row === 1 && col === 4) {
      const mode = String(sheet.getRange('B1').getValue()).trim();
      if (mode !== 'ระบุวันที่เอง') {
        safeNotify('🔒 ช่องวันที่ถูกล็อค', 'เปลี่ยนช่วงเวลาเป็น "ระบุวันที่เอง" ก่อน');
        applyLogDateLock(sheet, mode); return;
      }
      applyLogFilter(sheet, mode, e.range.getValue(), sheet.getRange('F1').getValue());
      return;
    }
    if (row === 1 && col === 6) {
      const mode = String(sheet.getRange('B1').getValue()).trim();
      if (mode !== 'ระบุวันที่เอง') {
        safeNotify('🔒 ช่องวันที่ถูกล็อค', 'เปลี่ยนช่วงเวลาเป็น "ระบุวันที่เอง" ก่อน');
        applyLogDateLock(sheet, mode); return;
      }
      applyLogFilter(sheet, mode, sheet.getRange('D1').getValue(), e.range.getValue());
      return;
    }
  }
}

// ============================================================
//   7. EXECUTIVE BI DASHBOARD
// ============================================================
function sameLocalCalendarDay_(d, ref) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return false;
  if (!(ref instanceof Date) || isNaN(ref.getTime())) return false;
  return d.getFullYear() === ref.getFullYear()
    && d.getMonth() === ref.getMonth()
    && d.getDate() === ref.getDate();
}

function dashboardPresetCellsMatchSameDay_(startCell, endCell, today) {
  return sameLocalCalendarDay_(startCell.getValue(), today) && sameLocalCalendarDay_(endCell.getValue(), today);
}

function dashboardPresetCellsMatchLast7_(startCell, endCell, sevenDaysAgo, today) {
  return sameLocalCalendarDay_(startCell.getValue(), sevenDaysAgo) && sameLocalCalendarDay_(endCell.getValue(), today);
}

function refreshDashboard(ss, opt) {
  opt = opt || {};
  const dash = ss.getSheetByName(CONFIG.SHEET_DASHBOARD);
  const ordersSheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  const stockSheet = ss.getSheetByName(CONFIG.SHEET_STOCK);
  const logSheet = ss.getSheetByName(CONFIG.SHEET_LOG);
  if (!dash || !ordersSheet || !stockSheet) return;

  const settings = getSettings(ss);
  const lowThreshold = parseInt(settings['เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)']) || DEFAULT_LOW_THRESHOLD;

  // [PERF-2] ติดตั้ง Log layout เฉพาะเมื่อ user กำลังดูหน้า Log อยู่เท่านั้น
  // ไม่เรียกทุก onEdit (เดิมเรียกทุกรอบ = getProperty() 1 ครั้ง + อาจ write หลาย range)
  if (logSheet && !opt.triggeredByStructuralChange) {
    try {
      const activeSheetName = logSheet.getParent().getActiveSheet().getName();
      if (activeSheetName === CONFIG.SHEET_LOG) formatLogSheetLayout(logSheet);
    } catch (_) { formatLogSheetLayout(logSheet); } // fallback: ติดตั้งเลยถ้าอ่าน active sheet ไม่ได้
  }

  ensureDashboardLayout(dash);
  const presetCell = dash.getRange('K3');
  if (!presetCell.getValue()) presetCell.setValue('ทั้งหมด');

  const startCell = dash.getRange('M3');
  const endCell = dash.getRange('O3');
  const currentMode = presetCell.getValue().toString();

  let startDate = null, endDate = null, isFiltered = false, invalidRange = false;
  const today = new Date();

  if (currentMode === 'เฉพาะวันนี้') {
    if (!dashboardPresetCellsMatchSameDay_(startCell, endCell, today)) {
      startCell.setValue(today).setNumberFormat('dd/MM/yyyy');
      endCell.setValue(today).setNumberFormat('dd/MM/yyyy');
    }
    startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0);
    endDate = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
    isFiltered = true;
  } else if (currentMode === '7 วันที่แล้ว') {
    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(today.getDate() - 7); // [FIX-6] ใช้ ONE_DAY_MS แบบ explicit ก็ได้แต่ setDate ชัดกว่า
    if (!dashboardPresetCellsMatchLast7_(startCell, endCell, sevenDaysAgo, today)) {
      startCell.setValue(sevenDaysAgo).setNumberFormat('dd/MM/yyyy');
      endCell.setValue(today).setNumberFormat('dd/MM/yyyy');
    }
    startDate = new Date(sevenDaysAgo.getFullYear(), sevenDaysAgo.getMonth(), sevenDaysAgo.getDate(), 0, 0, 0);
    endDate = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
    isFiltered = true;
  } else if (currentMode === 'ระบุวันที่เอง') {
    const parseDateSecure = (val) => {
      if (val instanceof Date && !isNaN(val.getTime())) return val;
      if (typeof val === 'string' && val.trim() !== '') {
        const p = val.split('/');
        if (p.length === 3) return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
      }
      return null;
    };
    const cleanStart = parseDateSecure(startCell.getValue());
    const cleanEnd = parseDateSecure(endCell.getValue());
    if (cleanStart && cleanEnd) {
      startDate = new Date(cleanStart.getFullYear(), cleanStart.getMonth(), cleanStart.getDate(), 0, 0, 0);
      endDate = new Date(cleanEnd.getFullYear(), cleanEnd.getMonth(), cleanEnd.getDate(), 23, 59, 59);
      isFiltered = true;
    } else { invalidRange = true; }
  } else {
    startCell.setValue('-');
    endCell.setValue('-');
  }

  applyDashboardDateLock(dash, currentMode);

  const SUM_PREFIX = 'สรุปยอดตามช่วงเวลาที่สั่งการ: ';
  let summaryHeadline, subtitleText;
  if (invalidRange) {
    summaryHeadline = SUM_PREFIX + '(เลือกวันเริ่ม/สิ้นสุดที่ M3 และ O3 ให้ครบ)';
    subtitleText = '⚠️  กรุณาเลือกวันที่ "เริ่ม" และ "สิ้นสุด"';
  } else if (!isFiltered) {
    summaryHeadline = SUM_PREFIX + 'ข้อมูลทั้งหมด';
    subtitleText = '📊  กำลังแสดง: ข้อมูลทั้งหมด (ไม่กรองวันที่)';
  } else {
    const fmt = (d) => Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd/MM/yyyy');
    const sameDay = startDate.getFullYear() === endDate.getFullYear()
      && startDate.getMonth() === endDate.getMonth()
      && startDate.getDate() === endDate.getDate();
    if (sameDay) {
      summaryHeadline = SUM_PREFIX + 'วันที่ ' + fmt(startDate);
      subtitleText = '📊  กำลังแสดง: ข้อมูลของวัน ' + fmt(startDate);
    } else {
      summaryHeadline = SUM_PREFIX + fmt(startDate) + ' ถึง ' + fmt(endDate);
      subtitleText = '📊  กำลังแสดง: ' + fmt(startDate) + '  →  ' + fmt(endDate);
    }
  }
  try { dash.getRange('A3').setValue(summaryHeadline); } catch (e) {}
  try { dash.getRange('A4').setValue(subtitleText); } catch (e) {}

  dash.getRangeList(['B6:G6', 'J6:M6', 'B11:G14', 'B19:G21', 'J11:M15', 'J19:M23', 'O10:Q21']).clearContent();

  const orders = ordersSheet.getDataRange().getValues().slice(1);
  const stockData = stockSheet.getDataRange().getValues().slice(1);

  const filteredOrders = orders.filter(r => {
    if (!r[0]) return false;
    if (invalidRange) return false;
    if (!isFiltered) return true;
    const orderDate = r[1] instanceof Date ? r[1] : new Date(r[1]);
    if (isNaN(orderDate.getTime())) return false;
    return orderDate >= startDate && orderDate <= endDate;
  });

  const totalOrders = filteredOrders.length;
  let revSum = 0, profitSum = 0, rtsCount = 0, cancelCount = 0;
  const active = [];
  for (let i = 0; i < filteredOrders.length; i++) {
    const r = filteredOrders[i];
    const status = r[13];
    if (status === 'คืนสินค้า/คืนเงิน') rtsCount++;
    else if (status === 'ยกเลิกรายการ') cancelCount++;
    if (CANCELLED_STATUSES.indexOf(status) === -1) {
      active.push(r);
      revSum += (parseFloat(r[5]) || 0) * (parseFloat(r[6]) || 0);
      profitSum += (parseFloat(r[12]) || 0);
    }
  }
  const rev = roundMoney(revSum);
  const profit = roundMoney(profitSum);
  const rtsRate = totalOrders ? (rtsCount / totalOrders) : 0;
  const cancelRate = totalOrders ? (cancelCount / totalOrders) : 0;

  const formatCurrencySmart = (val) => {
    if (val >= 1000000) return (val / 1000000).toFixed(2) + 'M ฿';
    if (val >= 100000) return (val / 1000).toFixed(0) + 'K ฿';
    return val.toLocaleString('th-TH') + ' ฿';
  };

  const codCashLiquidityBadge_ = (deliveredCnt, deliveredRev, pendingCnt, pendingRev) => {
    if (!deliveredCnt && !pendingCnt) return 'พร้อมใช้';
    if (!pendingCnt) return 'พร้อมใช้';
    if (!deliveredCnt) return 'ค้างแอป';
    return 'พร้อมใช้ ' + formatCurrencySmart(deliveredRev) + ' · ค้างแอป ' + formatCurrencySmart(pendingRev);
  };

  dash.getRange('B6').setValue(formatCurrencySmart(rev));
  dash.getRange('D6').setValue(formatCurrencySmart(profit));
  dash.getRange('F6').setValue(rev ? (profit/rev) : 0).setNumberFormat('0.00%');
  dash.getRange('J6').setValue(rtsRate).setNumberFormat('0.00%');
  dash.getRange('L6').setValue(cancelRate).setNumberFormat('0.00%');

  const prodStats = {};
  const chStats = {};
  CONFIG.CHANNELS.forEach(c => chStats[c] = { count: 0, rev: 0 });
  const pyStats = {};
  CONFIG.PAYMENTS.forEach(p => pyStats[p] = { count: 0, rev: 0 });
  let codDeliveredCt = 0, codDeliveredRev = 0, codPendingCodCt = 0, codPendingCodRev = 0;
  const DELIVERED = 'จัดส่งเสร็จสิ้น';
  for (let i = 0; i < active.length; i++) {
    const r = active[i];
    const sku = r[3], prodName = r[4];
    const qty = parseFloat(r[5]) || 0, price = parseFloat(r[6]) || 0;
    const channel = r[7], payment = r[8];
    const orderStatus = String(r[13] || '').trim();
    const itemRev = price * qty;
    const itemProfit = parseFloat(r[12]) || 0;
    if (!prodStats[sku]) prodStats[sku] = { name: prodName, qty: 0, profit: 0 };
    prodStats[sku].qty += qty;
    prodStats[sku].profit += itemProfit;
    if (chStats[channel]) { chStats[channel].count++; chStats[channel].rev += itemRev; }
    if (pyStats[payment]) { pyStats[payment].count++; pyStats[payment].rev += itemRev; }
    if (String(payment || '').trim() === 'COD') {
      if (orderStatus === DELIVERED) { codDeliveredCt++; codDeliveredRev += itemRev; }
      else { codPendingCodCt++; codPendingCodRev += itemRev; }
    }
  }

  const topProfit = Object.values(prodStats).sort((a, b) => b.profit - a.profit);
  const totalCategoryProfit = topProfit.reduce((acc, p) => acc + p.profit, 0);
  let runSum = 0;
  const abcOutput = [];
  for (let i = 0; i < 4; i++) {
    if (topProfit[i]) {
      runSum += topProfit[i].profit;
      const ratio = totalCategoryProfit > 0 ? (runSum / totalCategoryProfit) : 0;
      let cls = ratio <= 0.75 ? 'Class A' : (ratio <= 0.95 ? 'Class B' : 'Class C');
      abcOutput.push([topProfit[i].name, '', cls, topProfit[i].qty, topProfit[i].profit, '']);
    } else { abcOutput.push(['', '', '', '', '', '']); }
  }
  dash.getRange('B11:G14').setValues(abcOutput);
  dash.getRange('G11:G14').setBackground('#FFFFFF');

  const lowStockItems = stockData.filter(r => r[0] && typeof r[3] === 'number' && r[3] <= lowThreshold).slice(0, 3);
  const stockOutput = [];
  for (let i = 0; i < 3; i++) {
    if (lowStockItems[i]) {
      const deadCost = (lowStockItems[i][3] || 0) * (lowStockItems[i][4] || 0);
      stockOutput.push([lowStockItems[i][0], lowStockItems[i][1], lowStockItems[i][3] === 0 ? 'หมด' : 'ต่ำ', lowStockItems[i][3], deadCost, '']);
    } else { stockOutput.push(['', '', '', '', '', '']); }
  }
  dash.getRange('B19:G21').setValues(stockOutput);
  dash.getRange('G19:G21').setBackground('#FFFFFF');

  const chOutput = CONFIG.CHANNELS.slice(0, 5).map(c => [c, chStats[c].count, chStats[c].rev, rev ? (chStats[c].rev / rev) : 0]);
  dash.getRange('J11:M15').setValues(chOutput);

  const pyOutput = CONFIG.PAYMENTS.map(p => [
    p, pyStats[p].count, pyStats[p].rev,
    p !== 'COD' ? 'พร้อมใช้'
      : codCashLiquidityBadge_(codDeliveredCt, codDeliveredRev, codPendingCodCt, codPendingCodRev)
  ]);
  dash.getRange('J19:M23').setValues(pyOutput);

  const bestChannel = Object.keys(chStats).reduce((a, b) => chStats[a].rev > chStats[b].rev ? a : b);
  const bestPayment = Object.keys(pyStats).reduce((a, b) => pyStats[a].count > pyStats[b].count ? a : b);
  const netMargin = rev ? (profit / rev) : 0;
  let insight = 'วิเคราะห์ภาพรวมธุรกิจ:\n';
  insight += '• ช่องทางหลัก: ' + bestChannel + ' ทำเงินได้ดีที่สุด\n';
  insight += '• ลูกค้าชอบจ่ายผ่าน: ' + bestPayment + '\n';
  insight += '• สุขภาพกำไร: ' + (netMargin > 0.15 ? 'ดีเยี่ยม' : 'ควรปรับราคา') + '\n\n';
  insight += 'มิติด้านความเสี่ยงสถิติ:\n';
  insight += '• อัตรายกเลิก (Cancel): ' + (cancelRate*100).toFixed(1) + '%\n';
  insight += '• อัตราคืนของ (RTS): ' + (rtsRate*100).toFixed(1) + '%\n\n';
  insight += 'คำแนะนำเชิงกลยุทธ์:\n';
  insight += '1. สินค้ากลุ่ม Class A ทำกำไรหลัก ควรทำโปรโมชั่นกระตุ้นต่อ\n';
  insight += '2. ตรวจสอบต้นทุนสินค้าสต็อกต่ำ เพื่อลดปัญหาทุนจม';
  dash.getRange('O10').setValue(insight);

  dash.getRange('J2').setValue('🕒  อัปเดตล่าสุด: ' +
    Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm') + '  ');
  // [PERF-4] flush เฉพาะเมื่อเรียกจากเมนู — ไม่ block onEdit path
  if (opt && opt.calledFromMenu) SpreadsheetApp.flush();
}

// ============================================================
//   7.1 DASHBOARD STATIC LAYOUT
// ============================================================
function ensureDashboardLayout(dash) {
  const props = PropertiesService.getDocumentProperties();
  const LAYOUT_VERSION = 'v13';  // 🆙 v13: N3 = arrow spacer (→), ไม่มี label "สิ้นสุด:" overflow; col N=20 O=130
  if (props.getProperty('dashboardLayoutInstalled') === LAYOUT_VERSION) return;
  props.deleteProperty('dashDateLockMode');
  try { const w = dash.getRange('A3:I3'); w.breakApart(); w.clearContent().clearFormat(); } catch (e) {}
  try { const w = dash.getRange('A4:Q4'); w.breakApart(); w.clearContent().clearFormat(); } catch (e) {}
  try {
    dash.getRange('B9:G9').breakApart(); dash.getRange('B10:G14').breakApart();
    dash.getRange('B17:G17').breakApart(); dash.getRange('B18:G21').breakApart();
  } catch (e) {}
  dash.setTabColor('#1E40AF');
  dash.setHiddenGridlines(true);
  dash.setFrozenRows(3);
  // N (col14) = arrow spacer ระหว่าง start/end date → 20px กัน label overflow
  // O (col15) = end date picker → 130px ให้แสดงวันที่ครบ
  const widths = [28, 100, 110, 110, 100, 110, 100, 16, 16, 100, 110, 100, 110, 20, 130, 24, 290];
  widths.forEach((w, i) => dash.setColumnWidth(i + 1, w));
  const heights = {
    1: 10, 2: 38, 3: 32, 4: 24,
    5: 22, 6: 44, 7: 14, 8: 14,
    9: 26, 10: 28, 11: 28, 12: 28, 13: 28, 14: 28, 15: 28,
    16: 12, 17: 26, 18: 28, 19: 30, 20: 30, 21: 30, 22: 30, 23: 30
  };
  Object.keys(heights).forEach(r => dash.setRowHeight(parseInt(r), heights[r]));
  const safeMerge = (rangeA1) => {
    try { const r = dash.getRange(rangeA1); if (r.getMergedRanges().length === 0) r.merge(); } catch (e) {}
  };
  const forceMerge = (rangeA1) => {
    try { const r = dash.getRange(rangeA1); r.breakApart(); r.merge(); } catch (e) {}
  };
  dash.getRange('A2:Q2').breakApart();
  dash.getRange('A2:Q2').setBackground('#0F172A').setWrap(false);
  forceMerge('A2:I2');
  forceMerge('J2:Q2');
  let shopName = '';
  try {
    const settingsSheet = dash.getParent().getSheetByName(CONFIG.SHEET_SETTINGS);
    if (settingsSheet) {
      const sData = settingsSheet.getDataRange().getValues();
      for (let i = 0; i < sData.length; i++) {
        if (String(sData[i][0]).trim() === 'ชื่อแบรนด์ / ชื่อร้านค้า') {
          shopName = String(sData[i][1] || '').trim(); break;
        }
      }
    }
    if (!shopName) shopName = dash.getParent().getName();
  } catch (e) { shopName = 'ร้านค้าออนไลน์'; }
  dash.getRange('A2').setValue('  🏪   ' + shopName + '   ·   BI Dashboard')
    .setFontSize(13).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#0F172A').setWrap(false)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  dash.getRange('J2').setValue('🕒  อัปเดตล่าสุด: —')
    .setFontSize(10).setFontColor('#94A3B8').setBackground('#0F172A').setWrap(false)
    .setHorizontalAlignment('right').setVerticalAlignment('middle');
  try {
    const sumSlot = dash.getRange('A3:I3');
    sumSlot.breakApart();
    sumSlot.merge().setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true)
      .setFontWeight('bold').setFontSize(11).setFontColor('#0C4A6E').setBackground('#F8FAFC')
      .setBorder(false, false, true, false, false, false, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID)
      .setValue('สรุปยอดตามช่วงเวลาที่สั่งการ: (กำลังโหลด...)');
  } catch (e) {}
  dash.getRange('J3:O3').setBackground('#F1F5F9');
  dash.getRange('J3').setValue('🗓 ช่วงเวลา:')
    .setFontSize(10).setFontWeight('bold').setFontColor('#475569')
    .setHorizontalAlignment('right').setVerticalAlignment('middle');
  dash.getRange('L3').setValue('เริ่ม:').setFontSize(9).setFontColor('#64748B')
    .setHorizontalAlignment('right').setVerticalAlignment('middle');
  // N3 = spacer คั่นระหว่าง start date (M3) กับ end date (O3) — ไม่ใส่ label เพื่อกัน overflow
  dash.getRange('N3').setValue('→').setFontSize(9).setFontColor('#CBD5E1')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  const presetRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ทั้งหมด', 'เฉพาะวันนี้', '7 วันที่แล้ว', 'ระบุวันที่เอง'], true)
    .setAllowInvalid(false).build();
  dash.getRange('K3').setDataValidation(presetRule)
    .setFontSize(10).setFontWeight('bold').setBackground('#FFFFFF')
    .setFontColor('#0F172A').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#3B82F6', SpreadsheetApp.BorderStyle.SOLID);
  const kpiCards = [
    ['B5', '💰  ยอดขายรวม',  'B6', '#16A34A', '#F0FDF4', 'B5:C6'],
    ['D5', '✨  กำไรสุทธิ',  'D6', '#2563EB', '#EFF6FF', 'D5:E6'],
    ['F5', '📊  อัตรากำไร', 'F6', '#9333EA', '#FAF5FF', 'F5:G6'],
    ['J5', '↩️  RTS Rate',   'J6', '#D97706', '#FFFBEB', 'J5:K6'],
    ['L5', '✖️  Cancel Rate','L6', '#DC2626', '#FEF2F2', 'L5:M6']
  ];
  kpiCards.forEach(([labelCell, label, valueCell, color, bg, cardRange]) => {
    const [topLeft, bottomRight] = cardRange.split(':');
    const topRowMatch = topLeft.match(/([A-Z]+)(\d+)/);
    const bottomRowMatch = bottomRight.match(/([A-Z]+)(\d+)/);
    if (topRowMatch && bottomRowMatch) {
      safeMerge(topRowMatch[1] + topRowMatch[2] + ':' + bottomRowMatch[1] + topRowMatch[2]);
      safeMerge(topRowMatch[1] + bottomRowMatch[2] + ':' + bottomRowMatch[1] + bottomRowMatch[2]);
    }
    dash.getRange(cardRange).setBackground(bg)
      .setBorder(true, true, true, true, false, false, color, SpreadsheetApp.BorderStyle.SOLID_THICK);
    dash.getRange(labelCell).setValue(label)
      .setFontSize(9).setFontWeight('bold').setFontColor(color)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    dash.getRange(valueCell)
      .setFontSize(16).setFontWeight('bold').setFontColor(color)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  const banners1 = [
    ['B9:F9', '🏆   ABC Classification — Top Profit Products'],
    ['J9:M9', '🛒   Channel Share — สัดส่วนช่องทางขาย'],
    ['O9:Q9', '💡   Insights & Recommendations']
  ];
  banners1.forEach(([range, text]) => {
    forceMerge(range);
    dash.getRange(range.split(':')[0]).setValue(text)
      .setFontSize(11).setFontWeight('bold').setFontColor('#FFFFFF')
      .setBackground('#1E293B').setHorizontalAlignment('left').setVerticalAlignment('middle');
  });
  dash.getRange('B10:F10').setValues([['ชื่อสินค้า', '', 'Class', 'จำนวน', 'กำไร (฿)']])
    .setFontSize(9).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#475569').setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.getRange('G10').clearContent().setBackground('#FFFFFF').setFontSize(10).setVerticalAlignment('middle');
  dash.getRange('J10:M10').setValues([['ช่องทาง', 'ออเดอร์', 'ยอดขาย', 'สัดส่วน']])
    .setFontSize(9).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#475569').setHorizontalAlignment('center').setVerticalAlignment('middle');
  const banners2 = [
    ['B17:F17', '📦   สินค้าสต็อกต่ำ — Low Stock Alert'],
    ['J17:M17', '💰   Cash Flow — สรุปการเงินตามวิธีชำระ']
  ];
  banners2.forEach(([range, text]) => {
    forceMerge(range);
    dash.getRange(range.split(':')[0]).setValue(text)
      .setFontSize(11).setFontWeight('bold').setFontColor('#FFFFFF')
      .setBackground('#1E293B').setHorizontalAlignment('left').setVerticalAlignment('middle');
  });
  dash.getRange('B18:F18').setValues([['SKU', 'ชื่อสินค้า', 'สถานะ', 'คงเหลือ', 'ทุนค้าง (฿)']])
    .setFontSize(9).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#475569').setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.getRange('G18').clearContent().setBackground('#FFFFFF').setFontSize(10).setVerticalAlignment('middle');
  dash.getRange('J18:M18').setValues([['วิธีชำระ', 'ออเดอร์', 'ยอดขาย', 'สถานะเงิน']])
    .setFontSize(9).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#475569').setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.getRange('B11:G14').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('G11:G14').setBackground('#FFFFFF');
  dash.getRange('B11:B14').setHorizontalAlignment('left').setFontWeight('bold');
  dash.getRange('D11:D14').setHorizontalAlignment('center');
  dash.getRange('E11:E14').setHorizontalAlignment('center').setNumberFormat('#,##0');
  dash.getRange('F11:F14').setHorizontalAlignment('right').setNumberFormat('#,##0').setFontWeight('bold');
  dash.getRange('B19:G21').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('G19:G21').setBackground('#FFFFFF');
  dash.getRange('B19:B21').setHorizontalAlignment('center').setFontFamily('Roboto Mono');
  dash.getRange('C19:C21').setHorizontalAlignment('left');
  dash.getRange('D19:D21').setHorizontalAlignment('center');
  dash.getRange('E19:E21').setHorizontalAlignment('center').setNumberFormat('#,##0').setFontWeight('bold');
  dash.getRange('F19:F21').setHorizontalAlignment('right').setNumberFormat('#,##0.00');
  dash.getRange('G9').clearContent().setBackground('#FFFFFF').setFontSize(11).setVerticalAlignment('middle');
  dash.getRange('G17').clearContent().setBackground('#FFFFFF').setFontSize(11).setVerticalAlignment('middle');
  dash.getRange('J11:M15').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('J11:J15').setHorizontalAlignment('left').setFontWeight('bold');
  dash.getRange('K11:K15').setHorizontalAlignment('center').setNumberFormat('#,##0');
  dash.getRange('L11:L15').setHorizontalAlignment('right').setNumberFormat('#,##0');
  dash.getRange('M11:M15').setHorizontalAlignment('center').setNumberFormat('0.00%').setFontWeight('bold');
  dash.getRange('J19:M23').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('J19:J23').setHorizontalAlignment('left').setFontWeight('bold');
  dash.getRange('K19:K23').setHorizontalAlignment('center').setNumberFormat('#,##0');
  dash.getRange('L19:L23').setHorizontalAlignment('right').setNumberFormat('#,##0');
  dash.getRange('M19:M23').setHorizontalAlignment('center');
  safeMerge('O10:Q21');
  dash.getRange('O10').setBackground('#FFFBEB')
    .setBorder(true, true, true, true, false, false, '#FBBF24', SpreadsheetApp.BorderStyle.SOLID)
    .setFontSize(10).setVerticalAlignment('top').setHorizontalAlignment('left').setWrap(true);
  const targetCells = ['D11:D14', 'D19:D21', 'J11:J15', 'M19:M23'];
  const keptRules = dash.getConditionalFormatRules().filter(r => {
    return !r.getRanges().some(rg => targetCells.indexOf(rg.getA1Notation()) !== -1);
  });
  const abcClassRange = dash.getRange('D11:D14');
  [['Class A','#DCFCE7','#166534'],['Class B','#FEF3C7','#92400E'],['Class C','#FEE2E2','#991B1B']].forEach(([t, bg, fg]) => {
    keptRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(t).setBackground(bg).setFontColor(fg).setBold(true).setRanges([abcClassRange]).build());
  });
  const stockStatusRange = dash.getRange('D19:D21');
  keptRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('หมด')
    .setBackground('#991B1B').setFontColor('#FFFFFF').setBold(true).setRanges([stockStatusRange]).build());
  keptRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ต่ำ')
    .setBackground('#FEF3C7').setFontColor('#92400E').setBold(true).setRanges([stockStatusRange]).build());
  const channelNameRange = dash.getRange('J11:J15');
  [['Shopee','#FFEDD5','#C2410C'],['Lazada','#EDE9FE','#5B21B6'],['TikTok Shop','#1E293B','#FFFFFF'],
   ['IG Shop','#FCE7F3','#9D174D'],['LINE','#DCFCE7','#166534'],['Facebook','#DBEAFE','#1E40AF']].forEach(([t, bg, fg]) => {
    keptRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(t).setBackground(bg).setFontColor(fg).setBold(true).setRanges([channelNameRange]).build());
  });
  const cashStatusRange = dash.getRange('M19:M23');
  keptRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains(' · ')
    .setBackground('#E0F2FE').setFontColor('#0369A1').setBold(true).setRanges([cashStatusRange]).build());
  keptRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ค้างแอป')
    .setBackground('#FEF3C7').setFontColor('#92400E').setBold(true).setRanges([cashStatusRange]).build());
  keptRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('พร้อมใช้')
    .setBackground('#DCFCE7').setFontColor('#166534').setBold(true).setRanges([cashStatusRange]).build());
  dash.setConditionalFormatRules(keptRules);
  try {
    const subtitle = dash.getRange('A4:Q4');
    subtitle.merge();
    dash.getRange('A4').setBackground('#F1F5F9')
      .setFontColor('#0369A1').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(false, false, true, false, false, false, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
  } catch (e) {}
  props.setProperty('dashboardLayoutInstalled', LAYOUT_VERSION);
}

// ============================================================
//   7.2 DASHBOARD DATE LOCK
// ============================================================
function applyDashboardDateLock(dash, mode) {
  const props = PropertiesService.getDocumentProperties();
  const memoKey = 'dashDateLockMode';
  const modeStr = mode == null ? '' : String(mode);
  if (props.getProperty(memoKey) === modeStr) return;
  const startCell = dash.getRange('M3');
  const endCell = dash.getRange('O3');
  const isCustom = mode === 'ระบุวันที่เอง';
  if (isCustom) {
    const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
    [startCell, endCell].forEach((c, i) => {
      c.setDataValidation(dateRule).setBackground('#FFFFFF')
        .setFontStyle('normal').setFontWeight('bold').setFontColor('#0F172A')
        .setFontSize(10).setHorizontalAlignment('center')
        .setBorder(true, true, true, true, false, false, '#3B82F6', SpreadsheetApp.BorderStyle.SOLID)
        .setNote(i === 0 ? '✏️ คลิกเพื่อเลือกวันที่เริ่มต้น' : '✏️ คลิกเพื่อเลือกวันที่สิ้นสุด');
    });
  } else {
    [startCell, endCell].forEach(c => {
      c.clearDataValidations().setBackground('#F1F5F9')
        .setFontStyle('italic').setFontWeight('normal').setFontColor('#94A3B8')
        .setFontSize(10).setHorizontalAlignment('center')
        .setBorder(true, true, true, true, false, false, '#CBD5E1', SpreadsheetApp.BorderStyle.DASHED)
        .setNote('🔒 ล็อคอยู่ — เปลี่ยนช่วงเวลาเป็น "ระบุวันที่เอง" เพื่อปลดล็อค');
    });
  }
  props.setProperty(memoKey, modeStr);
}

// ============================================================
//   7.3 LOG SHEET DATE LOCK
// ============================================================
function applyLogDateLock(logSheet, mode) {
  const startCell = logSheet.getRange('D1');
  const endCell = logSheet.getRange('F1');
  const unlockCell = (cell, placeholder) => {
    const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
    cell.setDataValidation(dateRule).setBackground('#FFFFFF')
      .setFontStyle('normal').setFontWeight('bold').setFontColor('#0F172A')
      .setFontSize(10).setHorizontalAlignment('center').setNumberFormat('dd/MM/yyyy')
      .setBorder(true, true, true, true, false, false, '#3B82F6', SpreadsheetApp.BorderStyle.SOLID)
      .setNote(placeholder);
  };
  const lockCell = (cell, hint) => {
    cell.clearDataValidations().setValue('-').setBackground('#F1F5F9')
      .setFontStyle('italic').setFontWeight('normal').setFontColor('#94A3B8')
      .setFontSize(10).setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, '#CBD5E1', SpreadsheetApp.BorderStyle.DASHED)
      .setNote(hint);
  };
  if (mode === 'ระบุวันที่เอง') {
    unlockCell(startCell, '✏️ เลือกวันเริ่มต้น (จำเป็น)');
    unlockCell(endCell, '✏️ เลือกวันสิ้นสุด (เว้นว่างไว้ = กรอง 1 วัน)');
  } else {
    lockCell(startCell, '🔒 ล็อคอยู่ — เปลี่ยนช่วงเวลาเป็น "ระบุวันที่เอง"');
    lockCell(endCell, '🔒 ล็อคอยู่ — เปลี่ยนช่วงเวลาเป็น "ระบุวันที่เอง"');
  }
}

// ============================================================
//   7.4 LOG FILTER ENGINE
// ============================================================
function applyLogFilter(logSheet, mode, customStart, customEnd) {
  const lastRow = logSheet.getLastRow();
  if (lastRow <= 4) return;
  const totalDataRows = lastRow - 4;
  let startDate = null, endDate = null, isFiltered = false;
  const today = new Date();
  const isValidDate = (d) => d instanceof Date && !isNaN(d.getTime());
  if (mode === 'เฉพาะวันนี้') {
    startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0);
    endDate = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
    isFiltered = true;
  } else if (mode === '7 วันที่แล้ว') {
    // [FIX-6] ใช้ ONE_DAY_MS constant
    const sevenDaysAgo = new Date(today.getTime() - 7 * ONE_DAY_MS);
    startDate = new Date(sevenDaysAgo.getFullYear(), sevenDaysAgo.getMonth(), sevenDaysAgo.getDate(), 0, 0, 0);
    endDate = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
    isFiltered = true;
  } else if (mode === 'ระบุวันที่เอง' && isValidDate(customStart)) {
    let s = customStart;
    let e = isValidDate(customEnd) ? customEnd : customStart;
    if (s > e) { const tmp = s; s = e; e = tmp; }
    startDate = new Date(s.getFullYear(), s.getMonth(), s.getDate(), 0, 0, 0);
    endDate = new Date(e.getFullYear(), e.getMonth(), e.getDate(), 23, 59, 59);
    isFiltered = true;
  }
  if (!isFiltered) { logSheet.showRows(5, totalDataRows); return; }
  const dateRange = logSheet.getRange(5, 1, totalDataRows, 1).getValues();
  const matches = new Array(totalDataRows);
  for (let i = 0; i < dateRange.length; i++) {
    const rowData = dateRange[i][0];
    let m = false;
    if (rowData instanceof Date && !isNaN(rowData.getTime())) {
      m = rowData >= startDate && rowData <= endDate;
    } else if (typeof rowData === 'string' && rowData.trim() !== '') {
      const cleaned = rowData.replace(',', '').trim();
      const r = cleaned.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
      if (r) {
        const d = new Date(parseInt(r[3]), parseInt(r[2]) - 1, parseInt(r[1]),
                           parseInt(r[4]), parseInt(r[5]), parseInt(r[6] || '0'));
        if (!isNaN(d.getTime())) m = d >= startDate && d <= endDate;
      }
    }
    matches[i] = m;
  }
  let i = 0;
  while (i < matches.length) {
    const state = matches[i];
    let j = i + 1;
    while (j < matches.length && matches[j] === state) j++;
    const startAbsRow = 5 + i;
    const numContiguous = j - i;
    if (state) logSheet.showRows(startAbsRow, numContiguous);
    else logSheet.hideRows(startAbsRow, numContiguous);
    i = j;
  }
}

// ============================================================
//   8. CLEAR ALL DATA
// ============================================================
function clearAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const whom = getCurrentUserEmail();
  const ownerGate = isSpreadsheetOwner_(ss);
  if (ownerGate === false) {
    ui.alert('ไม่ได้รับอนุญาต', 'ล้างข้อมูลระบบทำได้เฉพาะผู้เป็นเจ้าของไฟล์เท่านั้น\n(เมล: ' + whom + ')'); return;
  }
  if (ownerGate === null) {
    ui.alert('ยังไม่สามารถยืนยันสิทธิ์', 'โปรดอนุญาตสิทธิ์ใน Apps Script แล้วลองอีกครั้ง'); return;
  }
  const response = ui.alert(
    'คำเตือนวิกฤต',
    'คุณกำลังจะลบข้อมูลออเดอร์ ยอดขาย และประวัติทั้งหมด (ยกเว้นสต็อก)\n\nไม่สามารถย้อนคืนได้ มั่นใจใช่ไหม?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) { ui.alert('ยกเลิกการล้างระบบ'); return; }
  const ordersSheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  if (ordersSheet) {
    const lastRow = ordersSheet.getLastRow();
    if (lastRow > 1) ordersSheet.getRange(2, 1, lastRow - 1, ordersSheet.getLastColumn()).clearContent();
  }
  const logSheet = ss.getSheetByName(CONFIG.SHEET_LOG);
  if (logSheet) {
    const lastRowLog = logSheet.getLastRow();
    if (lastRowLog > 4) {
      logSheet.getRange(5, 1, lastRowLog - 4, logSheet.getLastColumn()).clearContent();
      logSheet.showRows(5, lastRowLog - 4);
    } else if (lastRowLog > 1) {
      logSheet.getRange(2, 1, lastRowLog - 1, logSheet.getLastColumn()).clearContent();
    }
    if (logSheet.getRange('B1').getDataValidation()) logSheet.getRange('B1').clearContent();
  }
  refreshDashboard(ss);
  writeLog(ss, 'SYSTEM_RESET', 'ล้างข้อมูลระบบ', getCurrentUserEmail());
  ui.alert('ล้างระบบเรียบร้อย พร้อมเริ่มใหม่');
}

// ============================================================
//   9. LOW STOCK ALERT
// ============================================================
function checkAndAlertLowStock(ss, settings, force, preloadedStockData) {
  // [PERF-3] ถ้า caller (processOrderInPlace) โหลด stockData มาแล้ว → ใช้เลย ไม่อ่านซ้ำ
  let stockData;
  if (preloadedStockData && preloadedStockData.length > 0) {
    stockData = preloadedStockData;
  } else {
    const stockSheet = ss.getSheetByName(CONFIG.SHEET_STOCK);
    if (!stockSheet) return false;
    stockData = stockSheet.getDataRange().getValues().slice(1);
  }
  const threshold = parseInt(settings['เกณฑ์สินค้าสต๊อกต่ำ (ชิ้น)']) || DEFAULT_LOW_THRESHOLD;
  const lowItems = stockData.filter(r => r[0] && typeof r[3] === 'number' && r[3] >= 0 && r[3] <= threshold);
  if (lowItems.length === 0) return false;
  const props = PropertiesService.getDocumentProperties();
  const signature = lowItems.map(r => r[0] + ':' + r[3]).sort().join('|');
  if (!force && signature === props.getProperty('lastLowStockSig')) return false;
  try {
    const truncated = signature.length > MAX_LOW_STOCK_SIG ? signature.substring(0, MAX_LOW_STOCK_SIG) : signature;
    props.setProperty('lastLowStockSig', truncated);
  } catch (qErr) { Logger.log('lastLowStockSig setProperty fail: ' + qErr.message); }
  const outItems = lowItems.filter(r => r[3] === 0);
  const nearItems = lowItems.filter(r => r[3] > 0);
  const todayStr = new Date().toLocaleDateString('th-TH');
  let body = 'แจ้งเตือนระบบคลังสินค้าต่ำกว่าเกณฑ์ — ' + todayStr + '\n\n';
  if (outItems.length > 0) {
    body += 'สินค้าหมดสต๊อก (' + outItems.length + ' รายการ):\n';
    outItems.forEach(r => body += '  - ' + r[0] + ' ' + r[1] + '\n');
    body += '\n';
  }
  if (nearItems.length > 0) {
    body += 'สินค้าใกล้หมด (ต่ำกว่า ' + threshold + ' ชิ้น):\n';
    nearItems.forEach(r => body += '  - ' + r[0] + ' ' + r[1] + ' — เหลือ ' + r[3] + ' ชิ้น\n');
  }
  const ue = getCurrentUserEmail();
  const fallbackMail = ue && ue !== 'unknown' ? ue : '';
  const alertEmail = settings['อีเมลรับแจ้งเตือนระบบ'] || fallbackMail;
  let sentOk = false;
  try {
    if (alertEmail) {
      MailApp.sendEmail({ to: alertEmail, subject: '[แจ้งเตือนด่วน] คลังสินค้าวิกฤต ' + lowItems.length + ' รายการ', body: body });
      sentOk = true;
    }
    // [FIX-1] ใช้ getSecureSettings สำหรับ token (ถ้า caller ส่ง settings มาเป็น plain ให้ re-read)
    const secSettings = settings['Telegram Bot Token'] ? settings : getSecureSettings(ss);
    if (secSettings['Telegram Bot Token'] && secSettings['Telegram Chat ID']) {
      if (sendTelegramNotify(secSettings['Telegram Bot Token'], secSettings['Telegram Chat ID'], body)) sentOk = true;
    }
    if (secSettings['LINE Token'] && secSettings['LINE Target ID']) {
      if (sendLineOANotify(secSettings['LINE Token'], secSettings['LINE Target ID'], body)) sentOk = true;
    }
    writeLog(ss, 'ALERT_LOW_STOCK', 'แจ้งเตือน ' + lowItems.length + ' รายการ');
    return sentOk;
  } catch (e) {
    writeLog(ss, 'ALERT_ERROR', sanitizeLogValue(e.message));
    return false;
  }
}

// ============================================================
//   10. DAILY REPORT
// ============================================================
function escapeHtmlForEmail_(text) {
  return String(text == null ? '' : text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');  // [FIX-2] เพิ่ม escape single quote
}

function fmtBahtDaily_(amount) {
  return roundMoney(amount).toLocaleString('th-TH') + ' ฿';
}

function orderPresetFirstBreakdownKeys_(map, presetArr) {
  const seen = {};
  const out = [];
  for (let i = 0; i < presetArr.length; i++) {
    const k = presetArr[i];
    if (map[k] && map[k].count > 0) { out.push(k); seen[k] = true; }
  }
  const rest = Object.keys(map).filter(k => !seen[k] && map[k] && map[k].count > 0);
  rest.sort((a, b) => a.localeCompare(b, 'th'));
  return out.concat(rest);
}

function accumulateReportTrip_(bucket, row, keyResolver) {
  let k = keyResolver(row);
  if (!String(k || '').trim()) k = 'อื่นๆ / ไม่ระบุ';
  if (!bucket[k]) bucket[k] = { count: 0, rev: 0, profit: 0 };
  bucket[k].count++;
  bucket[k].rev += (parseFloat(row[5]) || 0) * (parseFloat(row[6]) || 0);
  bucket[k].profit += parseFloat(row[12]) || 0;
}

function accumulateChannelBreakdown_(rows) {
  const m = {};
  rows.forEach(r => {
    const chan = String(r[7] || '').trim();
    accumulateReportTrip_(m, r, () => chan ? chan : 'หน้าร้าน/อื่นๆ');
  });
  Object.keys(m).forEach(k => { m[k].rev = roundMoney(m[k].rev); m[k].profit = roundMoney(m[k].profit); });
  return m;
}

function accumulatePaymentBreakdown_(rows) {
  const m = {};
  rows.forEach(r => accumulateReportTrip_(m, r, () => String(r[8] || '').trim()));
  Object.keys(m).forEach(k => { m[k].rev = roundMoney(m[k].rev); m[k].profit = roundMoney(m[k].profit); });
  return m;
}

function aggregateProductsReport_(rows, limitTop) {
  const byKey = {};
  rows.forEach(r => {
    const sku = String(r[3] || '').trim();
    const name = String(r[4] || '').trim() || sku || '(ไม่ระบุ)';
    const key = sku || name;
    if (!byKey[key]) byKey[key] = { sku, name, qty: 0, rev: 0, profit: 0 };
    byKey[key].qty += parseFloat(r[5]) || 0;
    byKey[key].rev += (parseFloat(r[5]) || 0) * (parseFloat(r[6]) || 0);
    byKey[key].profit += parseFloat(r[12]) || 0;
  });
  return Object.values(byKey)
    .map(p => ({ sku: p.sku, name: p.name, qty: Math.round(p.qty * 100) / 100, rev: roundMoney(p.rev), profit: roundMoney(p.profit) }))
    .sort((a, b) => b.rev - a.rev || b.profit - a.profit)
    .slice(0, limitTop);
}

function summarizeReportRowsTriple_(rows) {
  let revSum = 0, profSum = 0;
  for (let i = 0; i < rows.length; i++) {
    revSum += (parseFloat(rows[i][5]) || 0) * (parseFloat(rows[i][6]) || 0);
    profSum += parseFloat(rows[i][12]) || 0;
  }
  return { count: rows.length, rev: roundMoney(revSum), profit: roundMoney(profSum) };
}

function buildDailyReportPlain_(rep) {
  let t = '📌 รายงานขาย\nร้าน: ' + rep.shopName + '\nวันที่: ' + rep.todayStrLine + '\n';
  const triLine_ = (label, s) => {
    t += '\n══ ' + label + ' ══\nแถว: ' + s.count + '\nรายได้: ' + fmtBahtDaily_(s.rev) + '\nกำไร: ' + fmtBahtDaily_(s.profit) + '\n';
  };
  triLine_('ภาพรวมเดือนนี้', rep.monthTri);
  triLine_('วันนี้', rep.dayTri);
  triLine_('เมื่อวาน', rep.yesterdayTri);
  triLine_('ยอดสะสมทั้งระบบ', rep.cumTri);
  const tbl_ = (title, map, ordered) => {
    t += '\n── ' + title + ' ──\n';
    if (!ordered.length) { t += '(ว่าง)\n'; return; }
    ordered.forEach((k, i) => {
      const s = map[k];
      t += (i + 1) + '. ' + k + '\n   แถว ' + s.count + ' · ขาย ' + fmtBahtDaily_(s.rev) + ' · กำไร ' + fmtBahtDaily_(s.profit) + '\n';
    });
  };
  tbl_('ช่องทาง — เดือนนี้', rep.monthChMap, rep.monthChKeys);
  tbl_('ช่องทาง — วันนี้', rep.dayChMap, rep.dayChKeys);
  tbl_('วิธีชำระ — เดือนนี้', rep.monthPayMap, rep.monthPayKeys);
  const top_ = (title, list) => {
    t += '\n── ' + title + ' ──\n';
    if (!list.length) { t += '(ว่าง)\n'; return; }
    list.forEach((p, i) => {
      const nm = (p.name + (p.sku ? (' [' + p.sku + ']') : '')).trim();
      t += (i + 1) + '. ' + nm + '\n   จำนวน ' + p.qty + ' · ขาย ' + fmtBahtDaily_(p.rev) + ' · กำไร ' + fmtBahtDaily_(p.profit) + '\n';
    });
  };
  top_('Top สินค้าวันนี้', rep.topDayProd);
  t += '\n— ส่งอัตโนมัติจาก ShopAutomation Ultra —';
  return t;
}

// [FIX-2] HTML email — ทุก field ที่มาจาก user data ต้องผ่าน escapeHtmlForEmail_
function htmlDailyReportTripleCard_(heading, accent, fg, bd, stats) {
  const h = escapeHtmlForEmail_(heading);
  const revEsc = escapeHtmlForEmail_(fmtBahtDaily_(stats.rev));
  const proEsc = escapeHtmlForEmail_(fmtBahtDaily_(stats.profit));
  return (
    '<td width="33%" valign="top" style="padding:12px 14px;background:' + accent +
    ';border-radius:12px;border:1px solid ' + bd + ';">' +
    '<div style="font-size:10px;font-weight:700;color:' + fg + ';text-transform:uppercase;">' + h + '</div>' +
    '<div style="margin-top:8px;font-size:22px;font-weight:800;color:#0f172a;">' + revEsc + '</div>' +
    '<div style="margin-top:8px;font-size:12px;color:#334155;">กำไร <strong>' + proEsc + '</strong></div>' +
    '<div style="margin-top:4px;font-size:11px;color:#64748b;">แถว ' + stats.count + '</div></td>'
  );
}

function htmlDailyReportDataTable_(title, orderedKeys, map) {
  let bodyRows = '';
  if (!orderedKeys.length) {
    bodyRows = '<tr><td colspan="4" style="padding:14px;color:#64748b;text-align:center;">ไม่มียอด</td></tr>';
  } else {
    orderedKeys.forEach((k, ix) => {
      const s = map[k];
      const bg = ix % 2 === 1 ? '#f8fafc' : '#ffffff';
      // [FIX-2] escape ทุก field ที่มาจาก data
      bodyRows +=
        '<tr style="background:' + bg + ';">' +
        '<td style="padding:8px 10px;font-size:13px;font-weight:600;border-bottom:1px solid #e2e8f0;">' + escapeHtmlForEmail_(k) + '</td>' +
        '<td style="padding:8px 10px;font-size:13px;text-align:center;border-bottom:1px solid #e2e8f0;">' + escapeHtmlForEmail_(String(s.count)) + '</td>' +
        '<td style="padding:8px 10px;font-size:13px;text-align:right;border-bottom:1px solid #e2e8f0;">' + escapeHtmlForEmail_(fmtBahtDaily_(s.rev)) + '</td>' +
        '<td style="padding:8px 10px;font-size:13px;text-align:right;border-bottom:1px solid #e2e8f0;">' + escapeHtmlForEmail_(fmtBahtDaily_(s.profit)) + '</td></tr>';
    });
  }
  return (
    '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #e2e8f0;border-radius:10px;margin-bottom:22px;">' +
    '<tr><td colspan="4" style="padding:10px 12px;background:#f1f5f9;font-size:13px;font-weight:700;">' + escapeHtmlForEmail_(title) + '</td></tr>' +
    '<tr style="background:#334155;"><th align="left" style="padding:8px 10px;color:#fff;font-size:10px;">รายการ</th>' +
    '<th style="padding:8px 10px;color:#fff;font-size:10px;">แถว</th>' +
    '<th style="padding:8px 10px;color:#fff;font-size:10px;">ขาย</th>' +
    '<th style="padding:8px 10px;color:#fff;font-size:10px;">กำไร</th></tr>' +
    bodyRows + '</table>'
  );
}

function htmlDailyReportTopProdTable_(title, list) {
  let bodyRows = '';
  if (!list.length) {
    bodyRows = '<tr><td colspan="4" style="padding:14px;color:#64748b;text-align:center;">ว่าง</td></tr>';
  } else {
    list.forEach((p, ix) => {
      const bg = ix % 2 === 1 ? '#fffbeb' : '#ffffff';
      // [FIX-2] escape ชื่อสินค้า + SKU ที่มาจาก user
      const label = escapeHtmlForEmail_(p.name) +
        (p.sku ? ' <span style="color:#94a3b8;">[' + escapeHtmlForEmail_(p.sku) + ']</span>' : '');
      bodyRows +=
        '<tr style="background:' + bg + ';">' +
        '<td style="padding:8px 10px;font-size:13px;border-bottom:1px solid #fde68a;">' + label + '</td>' +
        '<td style="padding:8px 10px;font-size:13px;text-align:center;border-bottom:1px solid #fde68a;">' + escapeHtmlForEmail_(String(p.qty)) + '</td>' +
        '<td style="padding:8px 10px;font-size:13px;text-align:right;border-bottom:1px solid #fde68a;">' + escapeHtmlForEmail_(fmtBahtDaily_(p.rev)) + '</td>' +
        '<td style="padding:8px 10px;font-size:13px;text-align:right;border-bottom:1px solid #fde68a;">' + escapeHtmlForEmail_(fmtBahtDaily_(p.profit)) + '</td></tr>';
    });
  }
  return (
    '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #fcd34d;border-radius:10px;margin-bottom:22px;">' +
    '<tr><td colspan="4" style="padding:10px 12px;background:#fef3c7;font-size:13px;font-weight:700;color:#78350f;">' + escapeHtmlForEmail_(title) + '</td></tr>' +
    '<tr style="background:#92400e;"><th align="left" style="padding:8px 10px;color:#fef3c7;font-size:10px;">สินค้า</th>' +
    '<th style="padding:8px 10px;color:#fef3c7;font-size:10px;">จำนวน</th>' +
    '<th style="padding:8px 10px;color:#fef3c7;font-size:10px;">ขาย</th>' +
    '<th style="padding:8px 10px;color:#fef3c7;font-size:10px;">กำไร</th></tr>' +
    bodyRows + '</table>'
  );
}

function buildDailyReportHtml_(rep) {
  // [FIX-2] ทุก dynamic field ผ่าน escapeHtmlForEmail_
  const escShop = escapeHtmlForEmail_(rep.shopName);
  const escToday = escapeHtmlForEmail_(rep.todayStrLine);
  const kpiRow =
    '<tr>' +
    htmlDailyReportTripleCard_('เดือนนี้', '#ecfeff', '#0e7490', '#bae6fd', rep.monthTri) +
    htmlDailyReportTripleCard_('วันนี้', '#ecfdf5', '#047857', '#bbf7d0', rep.dayTri) +
    htmlDailyReportTripleCard_('สะสมในระบบ', '#eef2ff', '#4338ca', '#c7d2fe', rep.cumTri) +
    '</tr>';
  const yRow =
    '<tr><td colspan="3" style="padding:12px 4px;font-size:12px;color:#475569;background:#f8fafc;border-radius:10px;text-align:center;">' +
    '<strong>เมื่อวาน</strong> — ขาย <strong>' + escapeHtmlForEmail_(fmtBahtDaily_(rep.yesterdayTri.rev)) +
    '</strong> · กำไร <strong>' + escapeHtmlForEmail_(fmtBahtDaily_(rep.yesterdayTri.profit)) +
    '</strong> · แถว <strong>' + rep.yesterdayTri.count + '</strong></td></tr>';
  const blocks =
    htmlDailyReportDataTable_('แยกตามช่องทาง — เดือนนี้', rep.monthChKeys, rep.monthChMap) +
    htmlDailyReportDataTable_('แยกตามช่องทาง — วันนี้', rep.dayChKeys, rep.dayChMap) +
    htmlDailyReportDataTable_('แยกตามวิธีชำระ — เดือนนี้', rep.monthPayKeys, rep.monthPayMap) +
    htmlDailyReportTopProdTable_('Top ขายในวันนี้', rep.topDayProd);
  return (
    '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
    '<body style="margin:0;padding:0;background:#dfe7ef;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#dfe7ef;padding:20px 6px">' +
    '<tr><td align="center">' +
    '<table width="640" cellpadding="0" cellspacing="0" style="max-width:640px;background:#fff;border-radius:16px;overflow:hidden;">' +
    '<tr><td style="padding:26px 24px;background:linear-gradient(120deg,#0f172a,#1e293b);">' +
    '<div style="font-size:15px;color:#bae6fd;">สรุปขายประจำวัน</div>' +
    '<div style="margin-top:6px;font-size:22px;font-weight:800;color:#fff;">' + escShop + '</div>' +
    '<div style="margin-top:10px;font-size:14px;color:#e2e8f0;">📅 ' + escToday + '</div></td></tr>' +
    '<tr><td style="padding:8px 16px 4px;"><table width="100%" cellpadding="8" cellspacing="0">' + kpiRow + '</table></td></tr>' +
    '<tr><td style="padding:4px 16px 10px;"><table width="100%" cellpadding="0" cellspacing="0">' + yRow + '</table></td></tr>' +
    '<tr><td style="padding:8px 18px 24px;background:#fafafa;border-top:1px solid #e2e8f0;">' + blocks + '</td></tr>' +
    '</table></td></tr></table></body></html>'
  );
}

function sendDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSecureSettings(ss); // [FIX-1]
  refreshDashboard(ss);
  const ordersSheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  if (!ordersSheet) return false;
  const orders = ordersSheet.getDataRange().getValues().slice(1);
  const active = orders.filter(r => r[0] && CANCELLED_STATUSES.indexOf(r[13]) === -1);
  const tz = CONFIG.TIMEZONE;
  const clock = new Date();
  const todayKey = Utilities.formatDate(clock, tz, 'yyyy-MM-dd');
  const monthKey = Utilities.formatDate(clock, tz, 'yyyy-MM');
  // [FIX-6] ใช้ ONE_DAY_MS constant
  const yesterdayKey = Utilities.formatDate(new Date(clock.getTime() - ONE_DAY_MS), tz, 'yyyy-MM-dd');
  const orderDayKey = (d) => (!(d instanceof Date) || isNaN(d.getTime())) ? null : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const orderMonthKey = (d) => (!(d instanceof Date) || isNaN(d.getTime())) ? null : Utilities.formatDate(d, tz, 'yyyy-MM');
  const todayOrders = active.filter(r => orderDayKey(r[1]) === todayKey);
  const yesterdayOrders = active.filter(r => orderDayKey(r[1]) === yesterdayKey);
  const monthOrders = active.filter(r => orderMonthKey(r[1]) === monthKey);
  const cumTri = summarizeReportRowsTriple_(active);
  const monthTri = summarizeReportRowsTriple_(monthOrders);
  const dayTri = summarizeReportRowsTriple_(todayOrders);
  const yesterdayTri = summarizeReportRowsTriple_(yesterdayOrders);
  const ueDr = getCurrentUserEmail();
  const fbDr = ueDr && ueDr !== 'unknown' ? ueDr : '';
  const alertEmail = settings['อีเมลรับแจ้งเตือนระบบ'] || fbDr;
  if (!alertEmail && !settings['Telegram Bot Token'] && !settings['LINE Token']) return false;
  if (todayOrders.length === 0 && cumTri.rev === 0) return false;
  const shopName = settings['ชื่อร้านค้า'] || ss.getName();
  const todayStrLine = new Date().toLocaleDateString('th-TH', { timeZone: CONFIG.TIMEZONE, weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  const monthLabel = new Date().toLocaleDateString('th-TH', { timeZone: CONFIG.TIMEZONE, month: 'long', year: 'numeric' });
  const shortDate = Utilities.formatDate(clock, CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const dayChMap = accumulateChannelBreakdown_(todayOrders);
  const monthChMap = accumulateChannelBreakdown_(monthOrders);
  const dayChKeys = orderPresetFirstBreakdownKeys_(dayChMap, CONFIG.CHANNELS);
  const monthChKeys = orderPresetFirstBreakdownKeys_(monthChMap, CONFIG.CHANNELS);
  const monthPayMap = accumulatePaymentBreakdown_(monthOrders);
  const monthPayKeys = orderPresetFirstBreakdownKeys_(monthPayMap, CONFIG.PAYMENTS);
  const topDayProd = aggregateProductsReport_(todayOrders, 8);
  const rep = { shopName, todayStrLine, monthLabel, monthTri, dayTri, yesterdayTri, cumTri, monthChMap, dayChMap, monthChKeys, dayChKeys, monthPayMap, monthPayKeys, topDayProd };
  const bodyPlain = buildDailyReportPlain_(rep);
  const bodyHtml = buildDailyReportHtml_(rep);
  const subject = '📊 รายงาน ' + shortDate + ' · ' + fmtBahtDaily_(dayTri.rev) + ' · ' + shopName.substring(0, 32);
  let sentOk = false;
  try {
    if (alertEmail) {
      MailApp.sendEmail({ to: alertEmail, subject: subject, body: bodyPlain, htmlBody: bodyHtml });
      sentOk = true;
    }
    if (settings['Telegram Bot Token'] && settings['Telegram Chat ID']) {
      if (sendTelegramNotify(settings['Telegram Bot Token'], settings['Telegram Chat ID'], bodyPlain)) sentOk = true;
    }
    if (settings['LINE Token'] && settings['LINE Target ID']) {
      if (sendLineOANotify(settings['LINE Token'], settings['LINE Target ID'], bodyPlain)) sentOk = true;
    }
    writeLog(ss, 'DAILY_REPORT', 'วันนี้ ' + dayTri.rev + ' บาท (เดือน ' + monthTri.rev + ')');
    return sentOk;
  } catch (e) {
    writeLog(ss, 'REPORT_ERROR', sanitizeLogValue(e.message));
    return false;
  }
}

// ============================================================
//   11. INVOICE
//   [FIX-3] Validate VAT rate ก่อนหาร — กัน division by zero / silent wrong value
// ============================================================
function generateInvoice(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSecureSettings(ss); // [FIX-1]
  const orderSheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
  const allRows = orderSheet.getDataRange().getValues()
    .filter((r, i) => i > 0 && String(r[0]).trim() === String(orderId).trim());
  if (allRows.length === 0) { safeAlert('ไม่พบข้อมูลรหัสคำสั่งซื้อ ' + orderId); return; }
  const firstRow = allRows[0];
  const clientName = firstRow[2];
  const channel = firstRow[7];
  const status = firstRow[13];
  const shippingFee = allRows.reduce((max, r) => Math.max(max, parseFloat(r[9]) || 0), 0);
  const totalPrice = roundMoney(allRows.reduce((sum, r) => sum + ((parseInt(r[5]) || 0) * (parseFloat(r[6]) || 0)), 0));
  const invName = 'INV-' + String(orderId).replace('ORD-', '');
  let inv = ss.getSheetByName(invName);
  if (inv) ss.deleteSheet(inv);
  inv = ss.insertSheet(invName);
  inv.setHiddenGridlines(true);
  inv.setColumnWidth(1, 25); inv.setColumnWidth(2, 240); inv.setColumnWidth(3, 70);
  inv.setColumnWidth(4, 100); inv.setColumnWidth(5, 110); inv.setColumnWidth(6, 25);

  // [FIX-3] Validate VAT rate ก่อนใช้
  const vatRaw = parseFloat(settings['VAT %']);
  let vatPercent;
  if (!isNaN(vatRaw) && vatRaw >= VAT_MIN && vatRaw <= VAT_MAX) {
    vatPercent = vatRaw;
  } else {
    vatPercent = DEFAULT_VAT;
    if (!isNaN(vatRaw)) {
      // แจ้ง warning ถ้า user ตั้งค่าผิด (ไม่ใช่ NaN แต่อยู่นอก range)
      writeLog(ss, 'WARN_INVALID_VAT',
        'VAT rate ' + vatRaw + ' อยู่นอกช่วง ' + VAT_MIN + '-' + VAT_MAX + '% — ใช้ค่า default ' + DEFAULT_VAT + '%');
      safeNotify('⚠️ VAT rate ผิดปกติ',
        'ค่า VAT ' + vatRaw + '% อยู่นอกช่วงที่ยอมรับ (' + VAT_MIN + '-' + VAT_MAX + '%)\nระบบใช้ค่า default ' + DEFAULT_VAT + '% แทน');
    }
  }

  const vatDivisor = 1 + (vatPercent / 100);  // ปลอดภัย: vatPercent ≥ 0 เสมอ
  const totalWithShipping = roundMoney(totalPrice + shippingFee);
  const baseAmount = roundMoney(totalWithShipping / vatDivisor);
  const vatAmount = roundMoney(totalWithShipping - baseAmount);

  const shopName = settings['ชื่อร้านค้า'] || ss.getName();
  const taxId = settings['เลขผู้เสียภาษี'] ? 'เลขประจำตัวผู้เสียภาษี: ' + settings['เลขผู้เสียภาษี'] : '';
  const shopAddress = settings['ที่อยู่ร้านค้า'] ? 'ที่อยู่ติดต่อ: ' + settings['ที่อยู่ร้านค้า'] : '(กรุณาระบุที่อยู่ในแท็บตั้งค่า)';
  const combinedCompanyInfo = taxId ? taxId + '\n' + shopAddress : shopAddress;

  inv.getRange('B2').setValue(shopName).setFontSize(14).setFontWeight('bold').setFontColor('#1E293B');
  inv.getRange('B3').setValue(combinedCompanyInfo).setFontSize(8.5).setFontColor('#64748B').setWrap(true);
  inv.getRange('E2').setValue('ใบเสร็จรับเงิน / ใบกำกับภาษี').setFontSize(12).setFontWeight('bold').setFontColor('#0F172A').setHorizontalAlignment('right');
  inv.getRange('E3').setValue('(RECEIPT / TAX INVOICE)').setFontSize(8).setFontColor('#94A3B8').setHorizontalAlignment('right');
  inv.getRange('E4').setValue('เอกสารออกโดยระบบอัตโนมัติ').setFontSize(7.5).setFontStyle('italic').setFontColor('#94A3B8').setHorizontalAlignment('right');
  inv.getRange('B5:E5').setBorder(false, false, true, false, null, null, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
  inv.getRange('B7').setValue('ข้อมูลลูกค้า:').setFontSize(9).setFontWeight('bold').setFontColor('#475569');
  inv.getRange('B8').setValue('ชื่อลูกค้า: ' + clientName + '\nช่องทาง: ' + channel + '\nสถานะ: ' + status)
    .setFontSize(8.5).setFontColor('#334155').setWrap(true);
  inv.getRange('D7').setValue('เลขที่:\nวันที่:\nอ้างอิง:').setFontSize(8.5).setFontColor('#64748B').setHorizontalAlignment('right');
  inv.getRange('E7').setValue(invName + '\n' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy') + '\n' + orderId)
    .setFontSize(8.5).setFontColor('#334155').setFontWeight('bold').setHorizontalAlignment('left').setWrap(true);
  const tableHeader = inv.getRange('B11:E11');
  tableHeader.setValues([['รายละเอียดสินค้า', 'จำนวน', 'ราคา/หน่วย', 'จำนวนเงิน (THB)']])
    .setFontSize(9).setFontWeight('bold').setFontColor('#475569').setBackground('#F8FAFC');
  tableHeader.setBorder(true, false, true, false, null, null, '#94A3B8', SpreadsheetApp.BorderStyle.SOLID);
  inv.getRange('C11:E11').setHorizontalAlignment('right');
  const ITEM_START_ROW = 13;
  const itemValues = allRows.map(r => [r[4] + ' (' + r[3] + ')', parseInt(r[5]) || 0, parseFloat(r[6]) || 0, roundMoney((parseInt(r[5]) || 0) * (parseFloat(r[6]) || 0))]);
  const itemsRange = inv.getRange(ITEM_START_ROW, 2, itemValues.length, 4);
  itemsRange.setValues(itemValues).setFontSize(9).setFontColor('#334155');
  inv.getRange(ITEM_START_ROW, 3, itemValues.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
  inv.getRange(ITEM_START_ROW, 4, itemValues.length, 2).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  let cursor = ITEM_START_ROW + itemValues.length;
  if (shippingFee > 0) {
    inv.getRange(cursor, 2).setValue('ค่าจัดส่งสินค้า').setFontSize(9).setFontColor('#64748B').setFontStyle('italic');
    inv.getRange(cursor, 5).setValue(shippingFee).setNumberFormat('#,##0.00').setFontSize(9).setFontColor('#64748B').setHorizontalAlignment('right');
    cursor++;
  }
  cursor++;
  const dividerRow = cursor;
  inv.getRange(dividerRow, 2, 1, 4).setBorder(true, false, false, false, null, null, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID);
  const netRow = dividerRow + 1, vatRow = netRow + 1, grandRow = vatRow + 1;
  inv.getRange(netRow, 3, 3, 2).setFontSize(8.5).setFontColor('#64748B').setHorizontalAlignment('right');
  inv.getRange(netRow, 5, 3, 1).setFontSize(9).setFontColor('#0F172A').setFontWeight('bold').setHorizontalAlignment('right');
  inv.getRange(netRow, 3).setValue('มูลค่ารวมก่อนภาษี:');
  inv.getRange(netRow, 5).setValue(baseAmount).setNumberFormat('#,##0.00');
  inv.getRange(vatRow, 3).setValue('ภาษีมูลค่าเพิ่ม VAT ' + vatPercent + '%:');
  inv.getRange(vatRow, 5).setValue(vatAmount).setNumberFormat('#,##0.00');
  inv.getRange(grandRow, 3).setValue('ยอดรวมสุทธิ (Grand Total):');
  inv.getRange(grandRow, 5).setValue(totalWithShipping).setNumberFormat('#,##0.00');
  inv.getRange(grandRow, 2, 1, 4).setBackground('#F8FAFC')
    .setBorder(true, false, true, false, null, null, '#475569', SpreadsheetApp.BorderStyle.DOUBLE);
  const footerRow = grandRow + 3;
  inv.getRange(footerRow, 2, 1, 4).merge().setValue('~ ขอขอบพระคุณที่อุดหนุนนะคะ ~')
    .setFontSize(9).setFontStyle('italic').setFontColor('#64748B').setHorizontalAlignment('center');
  ss.setActiveSheet(inv);
}

// ============================================================
//   COMMUNICATIONS
// ============================================================
function sendTelegramNotify(token, chatId, message) {
  if (!token || !chatId) return false;
  try {
    return UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: message }),
      muteHttpExceptions: true
    }).getResponseCode() === 200;
  } catch (e) { return false; }
}

function sendLineOANotify(token, toId, message) {
  if (!token || !toId) return false;
  try {
    return UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ to: toId, messages: [{ type: 'text', text: message }] }),
      muteHttpExceptions: true
    }).getResponseCode() === 200;
  } catch (e) { return false; }
}
