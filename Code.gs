/**
 * Nil Transaction BO Detailed Report — Google Apps Script API
 *
 * Required sheets:
 *   1. Office Master
 *   2. Daily Nil Transactions
 *   3. BO Booking Performance
 *   4. BO Booking Performance Staging
 *
 * Deploy as Web app: Execute as Me; access as permitted by your organization.
 * Use the deployed /exec URL in APPS_SCRIPT_URL in the standalone HTML.
 *
 * GET actions:
 *   ?action=health
 *   ?action=getMaster
 *   ?action=getDaily&offset=0&limit=500  (paginated; max limit 1000)
 *   ?action=getBooking
 *   ?action=status
 *
 * POST actions:
 *   {"action":"beginSync","type":"master"}
 *   {"action":"saveBatch","type":"master","rows":[...]}
 *   {"action":"beginSync","type":"daily","months":["2026-10"]}
 *   {"action":"saveBatch","type":"daily","rows":[...]}
 *   {"action":"beginSync","type":"booking"}
 *   {"action":"saveBatch","type":"booking","rows":[...]}
 *   {"action":"commitBooking"}
 */

const MASTER_SHEET = 'Office Master';
const DAILY_SHEET = 'Daily Nil Transactions';
const BOOKING_SHEET = 'BO Booking Performance';
const BOOKING_STAGING_SHEET = 'BO Booking Performance Staging';
const MASTER_HEADERS = [
  'office_id','office_name','office_type_code','sol_id','pli_id',
  'valid_from','valid_to','delivery_office_flag','office_class',
  'office_status','division_name','sub_division_name','ho_name','so_name'
];
const DAILY_HEADERS = [
  'office_id','office_name','transac_date','total_bo_offices',
  'mail_nil_offices','mo_nil_offices','cbs_nil_offices','pli_nil_offices',
  'stamp_nil_offices','ippb_nil_offices','uploaded_at'
];
const BOOKING_KEYS = [
  'circle_id','circle_name','region_id','region_name',
  'division_id','division_name',
  'sub_division_id','sub_division_name',
  'office_id','office_name','booking_date','product_name',
  'walk_in_customer','registered_bulk','on_postal_service',
  'govt_service','walk_in_non_commercial','non_commercial',
  'commercial','e_commerce','pidpi','student_mail','others',
  'walk_in_commercial','walk_in_e_commerce',
  'bulk_non_commercial','bulk_commercial','bulk_e_commerce',
  'bo_rebooking','free_active_service','inter_fpo'
];
const BOOKING_HEADERS = [
  'Circle ID','Circle Name','Region ID','Region Name',
  'Division ID','Division Name',
  'Sub Division ID','Sub Division Name',
  'Office ID','Office Name','Booking Date','Product Name',
  'Walk-in Customer','Registered Bulk','On Postal Service',
  'Govt Service','Walk-in Non-Commercial','Non-Commercial',
  'Commercial','E-commerce','PIDPI','Student Mail','Others',
  'Walk-in Commercial','Walk-in E-commerce',
  'Bulk Non-Commercial','Bulk Commercial','Bulk E-commerce',
  'BO Rebooking','Free Active Service','Inter FPO'
];

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = params.action || 'health';
    let result;

    if (action === 'health') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      result = {
        ok: true,
        service: 'Nil Transaction BO Detailed Report API',
        version: '4.1-paginated',
        spreadsheetId: ss.getId(),
        spreadsheetName: ss.getName(),
        sheets: [MASTER_SHEET, DAILY_SHEET, BOOKING_SHEET],
        timestamp: new Date().toISOString()
      };
    } else if (action === 'getMaster' || action === 'downloadMaster') {
      result = {
        ok: true,
        action: action,
        rows: readRowsAsObjects_(MASTER_SHEET, MASTER_HEADERS),
        count: Math.max(0, getRequiredSheet_(MASTER_SHEET).getLastRow() - 1),
        timestamp: new Date().toISOString()
      };
    } else if (action === 'getDaily') {
      // Return a page, not the entire sheet. This avoids large JSON responses
      // timing out and lets the HTML client verify that every row was received.
      const sheet = getRequiredSheet_(DAILY_SHEET);
      const totalCount = Math.max(0, sheet.getLastRow() - 1);
      const requestedOffset = Number(params.offset);
      const requestedLimit = Number(params.limit);
      const offset = Number.isFinite(requestedOffset)
        ? Math.max(0, Math.floor(requestedOffset)) : 0;
      const limit = Number.isFinite(requestedLimit)
        ? Math.min(1000, Math.max(1, Math.floor(requestedLimit))) : 500;
      const count = Math.max(0, Math.min(limit, totalCount - offset));
      let rows = [];

      if (count > 0) {
        const values = sheet.getRange(offset + 2, 1, count, DAILY_HEADERS.length).getDisplayValues();
        rows = values.map(function(row) {
          const obj = {};
          DAILY_HEADERS.forEach(function(header, i) {
            obj[header] = row[i] === undefined ? '' : row[i];
          });
          return obj;
        });
      }

      result = {
        ok: true,
        action: action,
        rows: rows,
        count: rows.length,
        totalCount: totalCount,
        offset: offset,
        limit: limit,
        hasMore: offset + rows.length < totalCount,
        timestamp: new Date().toISOString()
      };
    } else if (action === 'getBooking' || action === 'downloadBooking') {
      const sheet = getRequiredSheet_(BOOKING_SHEET);
      const lastRow = sheet.getLastRow();
      const rows = [];
      if (lastRow > 1) {
        const values = sheet.getRange(2, 1, lastRow - 1, BOOKING_HEADERS.length).getDisplayValues();
        values.forEach(function(row) {
          const obj = {};
          BOOKING_KEYS.forEach(function(key, i) { obj[key] = row[i] || ''; });
          rows.push(obj);
        });
      }
      result = {ok:true, action:action, rows:rows, count:rows.length, timestamp:new Date().toISOString()};
    } else if (action === 'status' || action === 'verify') {
      result = getStatus_();
    } else {
      result = {ok:false, error:'Unsupported GET action: ' + action};
    }
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ok:false, error:String(err && err.message || err)});
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Missing POST JSON body.');
    const request = JSON.parse(e.postData.contents);
    let result;
    if (request.action === 'beginSync') result = beginSync_(request.type, request.months);
    else if (request.action === 'saveBatch') result = saveBatch_(request.type, normalizeRows_(request.rows));
    else if (request.action === 'commitBooking') result = commitBooking_();
    else throw new Error('Unsupported action: ' + String(request.action || ''));
    return jsonOutput_({ok:true, result:result});
  } catch (err) {
    return jsonOutput_({ok:false, error:String(err && err.message || err)});
  }
}

/** Run once manually from the Apps Script editor. */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, MASTER_SHEET, MASTER_HEADERS);
  ensureSheet_(ss, DAILY_SHEET, DAILY_HEADERS);
  ensureSheet_(ss, BOOKING_SHEET, BOOKING_HEADERS);
  const staging = ensureSheet_(ss, BOOKING_STAGING_SHEET, BOOKING_HEADERS);
  try { staging.hideSheet(); } catch (e) {}
  return 'Google Sheets storage is ready. Booking data sheet created: ' + BOOKING_SHEET;
}

/** Wrappers for google.script.run if the HTML is served by Apps Script. */
function beginSync(type, months) { return beginSync_(type, months); }
function saveBatch(type, rows) { return saveBatch_(type, normalizeRows_(rows)); }
function commitBooking() { return commitBooking_(); }

function normalizeRows_(rows) {
  if (Array.isArray(rows)) return rows;
  if (typeof rows === 'string') {
    const text = rows.trim();
    if (!text) return [];
    try { rows = JSON.parse(text); }
    catch (err) { throw new Error('rows was text but not valid JSON. Send rows as an array.'); }
  }
  if (rows && typeof rows === 'object') {
    if (Array.isArray(rows.rows)) return rows.rows;
    if (Array.isArray(rows.data)) return rows.data;
  }
  if (rows === undefined || rows === null) throw new Error('Missing rows array. Pass rows:[...] to saveBatch.');
  throw new Error('Invalid rows payload. Expected an array, JSON array string, or object containing rows/data array.');
}

function beginSync_(type, months) {
  if (type === 'booking') {
    const staging = getRequiredSheet_(BOOKING_STAGING_SHEET);
    const last = staging.getLastRow();
    if (last > 1) staging.getRange(2, 1, last - 1, BOOKING_HEADERS.length).clearContent();
    return {sheet:staging.getName(), clearedRows:Math.max(0,last-1), scope:'new booking snapshot staged; live sheet unchanged until commit'};
  }

  const sheet = getTargetSheet_(type);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return {sheet:sheet.getName(), clearedRows:0, months:months || []};
    if (type === 'master') {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      stampSync_('master');
      return {sheet:sheet.getName(), clearedRows:lastRow - 1, scope:'all master rows'};
    }

    const monthSet = new Set((Array.isArray(months) ? months : [])
      .map(function(m) { return String(m).slice(0,7); }).filter(Boolean));
    if (!monthSet.size) return {sheet:sheet.getName(), clearedRows:0, scope:'no month supplied; retained existing daily data'};

    const dates = sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues();
    const rowsToDelete = [];
    dates.forEach(function(row, i) {
      const date = normalizeSheetDate_(row[0]);
      if (date && monthSet.has(date.slice(0,7))) rowsToDelete.push(i + 2);
    });
    const blocks = [];
    rowsToDelete.forEach(function(rowNum) {
      const last = blocks[blocks.length - 1];
      if (last && rowNum === last.start + last.count) last.count++;
      else blocks.push({start:rowNum, count:1});
    });
    blocks.reverse().forEach(function(block) { sheet.deleteRows(block.start, block.count); });
    stampSync_('daily');
    return {sheet:sheet.getName(), clearedRows:rowsToDelete.length, months:Array.from(monthSet), scope:'uploaded months only'};
  } finally {
    lock.releaseLock();
  }
}

function saveBatch_(type, rows) {
  if (!Array.isArray(rows)) rows = normalizeRows_(rows);
  if (!rows.length) return {saved:0};
  if (rows.length > 500) throw new Error('Batch too large. Send no more than 500 rows per request.');

  const isMaster = type === 'master';
  const isBooking = type === 'booking';
  const sheet = isBooking ? getRequiredSheet_(BOOKING_STAGING_SHEET) : getTargetSheet_(type);
  const width = isMaster ? MASTER_HEADERS.length : (isBooking ? BOOKING_HEADERS.length : DAILY_HEADERS.length);
  const now = new Date();

  const values = rows.map(function(r) {
    r = r || {};
    if (isMaster) return [
      safeCell_(r.office_id), safeCell_(r.office_name), safeCell_(r.office_type_code),
      safeCell_(r.sol_id), safeCell_(r.pli_id), safeCell_(r.valid_from), safeCell_(r.valid_to),
      safeCell_(r.delivery_office_flag), safeCell_(r.office_class), safeCell_(r.office_status),
      safeCell_(r.division_name), safeCell_(r.sub_division_name), safeCell_(r.ho_name), safeCell_(r.so_name)
    ];
    if (isBooking) return BOOKING_KEYS.map(function(key) { return safeCell_(r[key]); });
    return [
      safeCell_(r.office_id), safeCell_(r.office_name), safeCell_(r.transac_date),
      safeCell_(r.total_bo_offices), safeCell_(r.mail_nil_offices), safeCell_(r.mo_nil_offices),
      safeCell_(r.cbs_nil_offices), safeCell_(r.pli_nil_offices), safeCell_(r.stamp_nil_offices),
      safeCell_(r.ippb_nil_offices), now
    ];
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const startRow = Math.max(2, sheet.getLastRow() + 1);
    sheet.getRange(startRow, 1, values.length, width).setValues(values);
    if (!isMaster && !isBooking) {
      sheet.getRange(startRow, 3, values.length, 1).setNumberFormat('@');
      sheet.getRange(startRow, 11, values.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }
    if (!isBooking) stampSync_(type);
    SpreadsheetApp.flush();
    return {sheet:sheet.getName(), saved:values.length, startRow:startRow};
  } finally {
    lock.releaseLock();
  }
}

function commitBooking_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const staging = getRequiredSheet_(BOOKING_STAGING_SHEET);
    const live = getRequiredSheet_(BOOKING_SHEET);
    const stagedCount = Math.max(0, staging.getLastRow() - 1);
    const liveLast = live.getLastRow();
    if (liveLast > 1) live.getRange(2, 1, liveLast - 1, BOOKING_HEADERS.length).clearContent();
    if (stagedCount > 0) {
      const values = staging.getRange(2, 1, stagedCount, BOOKING_HEADERS.length).getValues();
      live.getRange(2, 1, values.length, BOOKING_HEADERS.length).setValues(values);
    }
    SpreadsheetApp.flush();
    PropertiesService.getScriptProperties().setProperty('lastBookingWrite', new Date().toISOString());
    return {sheet:BOOKING_SHEET, committedRows:stagedCount, timestamp:new Date().toISOString()};
  } finally {
    lock.releaseLock();
  }
}

function getStatus_() {
  const masterSheet = getRequiredSheet_(MASTER_SHEET);
  const dailySheet = getRequiredSheet_(DAILY_SHEET);
  const bookingSheet = getRequiredSheet_(BOOKING_SHEET);
  const masterRows = Math.max(0, masterSheet.getLastRow() - 1);
  const dailyRows = Math.max(0, dailySheet.getLastRow() - 1);
  const bookingRows = Math.max(0, bookingSheet.getLastRow() - 1);
  const monthCounts = {};
  let latestTransactionDate = '';

  if (dailyRows) {
    const dates = dailySheet.getRange(2, 3, dailyRows, 1).getDisplayValues();
    dates.forEach(function(r) {
      const d = normalizeSheetDate_(r[0]);
      if (!d) return;
      const month = d.slice(0,7);
      monthCounts[month] = (monthCounts[month] || 0) + 1;
      if (d > latestTransactionDate) latestTransactionDate = d;
    });
  }

  const props = PropertiesService.getScriptProperties();
  return {
    ok:true,
    service:'Nil Transaction BO Detailed Report API',
    version:'4.1-paginated',
    spreadsheetName:SpreadsheetApp.getActiveSpreadsheet().getName(),
    masterRows:masterRows,
    dailyRows:dailyRows,
    bookingRows:bookingRows,
    monthCounts:monthCounts,
    latestTransactionDate:latestTransactionDate,
    lastMasterWrite:props.getProperty('lastMasterWrite') || '',
    lastDailyWrite:props.getProperty('lastDailyWrite') || '',
    lastBookingWrite:props.getProperty('lastBookingWrite') || '',
    timestamp:new Date().toISOString()
  };
}

function readRowsAsObjects_(sheetName, expectedHeaders) {
  const sheet = getRequiredSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const width = expectedHeaders.length;
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  const matches = expectedHeaders.every(function(h, i) {
    return String(headers[i] || '').trim() === h;
  });
  if (!matches) throw new Error('Header mismatch in "' + sheetName + '". Run setupSheets() and check column order.');
  const values = sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues();
  return values.map(function(row) {
    const obj = {};
    expectedHeaders.forEach(function(h, i) { obj[h] = row[i] === undefined ? '' : row[i]; });
    return obj;
  });
}

function normalizeSheetDate_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(value).trim().split('T')[0].split(' ')[0].replace(/\//g, '-');
  const p = s.split('-');
  if (p.length === 3 && p[0].length === 4) return p[0] + '-' + p[1].padStart(2,'0') + '-' + p[2].padStart(2,'0');
  if (p.length === 3 && p[2].length === 4) return p[2] + '-' + p[1].padStart(2,'0') + '-' + p[0].padStart(2,'0');
  return s;
}

function getTargetSheet_(type) {
  if (type !== 'master' && type !== 'daily' && type !== 'booking') {
    throw new Error('Invalid type. Expected "master", "daily", or "booking".');
  }
  return getRequiredSheet_(type === 'master' ? MASTER_SHEET : type === 'daily' ? DAILY_SHEET : BOOKING_SHEET);
}

function getRequiredSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" not found. Run setupSheets() first.');
  return sheet;
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const existing = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    const same = headers.every(function(h, i) { return String(existing[i] || '').trim() === h; });
    if (!same) throw new Error('Header mismatch in "' + name + '". Expected: ' + headers.join(', '));
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#dbeafe').setWrap(true);
  return sheet;
}

function safeCell_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string' && /^[=+@]/.test(value)) return "'" + value;
  return value;
}

function stampSync_(type) {
  const now = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(type === 'master' ? 'lastMasterWrite' : 'lastDailyWrite', now);
}

function jsonOutput_(object) {
  return ContentService.createTextOutput(JSON.stringify(object)).setMimeType(ContentService.MimeType.JSON);
}
