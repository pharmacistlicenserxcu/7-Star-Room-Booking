/**
 * ════════════════════════════════════════════════════════════════════
 * ระบบจองห้อง 7 ดาว — CLINIC OSPE (Standalone Google Apps Script)
 * คณะเภสัชศาสตร์ จุฬาลงกรณ์มหาวิทยาลัย
 * ════════════════════════════════════════════════════════════════════
 */

const CONFIG = {
  // ใส่ ID ของ Google Spreadsheet ที่ต้องการเก็บข้อมูล
  spreadsheetId: '131T81FTOVM9C-x41-QaPM-4rd5UXtvD7ADnkQOX3bUw',
  sheets: {
    roomBookings:   'RoomBookings',
    bookingHistory: 'BookingHistory',
    studentData:    'StudentData'
  }
};

/**
 * ──────────────────────────────────────────────────────────────
 * Entry Point (HTTP GET & POST REST API)
 * รองรับทั้งการเข้าชมผ่าน Apps Script ตรง และเป็น Backend REST API ให้ GitHub Pages
 * ──────────────────────────────────────────────────────────────
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action;

    // ถ้ามี query parameter action ให้ตอบเป็น JSON REST API
    if (action) {
      let result = null;

      if (action === 'getBookingsForWeek') {
        result = getBookingsForWeek(params.weekStartDate || params.date);
      } else if (action === 'getSystemInfo') {
        const ss = getSpreadsheet();
        const bSheet = ss ? findBookingSheet_(ss) : null;
        result = {
          success: true,
          spreadsheetId: ss ? ss.getId() : '',
          spreadsheetTitle: ss ? ss.getName() : '',
          activeSheetName: bSheet ? bSheet.getName() : '',
          totalRows: bSheet ? bSheet.getLastRow() : 0,
          allSheets: ss ? ss.getSheets().map(s => ({ name: s.getName(), rows: s.getLastRow() })) : []
        };
      } else if (action === 'verifyStudent') {
        result = verifyStudent(params.studentId);
      } else if (action === 'getDayQuota') {
        result = getDayQuota(params.date, params.studentId);
      } else if (action === 'getMyBookings') {
        result = getMyBookings(params.studentId);
      } else if (action === 'clearTestBookings') {
        result = clearTestBookings(params.pin);
      } else {
        result = { error: 'Unknown action: ' + action };
      }

      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ถ้าไม่มี action ให้ serve หน้า HTML ตามปกติ
    const template = HtmlService.createTemplateFromFile('index');
    template.webAppUrl = ScriptApp.getService().getUrl();
    return template.evaluate()
      .setTitle('ห้อง 7 ดาว — CLINIC OSPE Booking System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');

  } catch (err) {
    if (e && e.parameter && e.parameter.action) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return HtmlService.createHtmlOutput(`<h1>Error Loading Page</h1><p>${err.toString()}</p>`);
  }
}

/**
 * doPost(e)
 * รองรับการยิง fetch POST จาก GitHub Pages แบบ text/plain หรือ application/json
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (errJson) {
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action;
    let result = { success: false, error: 'No action specified' };

    if (action === 'createBooking') {
      result = createBooking(payload.data || payload);
    } else if (action === 'createMockCentralBooking') {
      result = createMockCentralBooking(payload.data || payload);
    } else if (action === 'cancelBooking') {
      const bId = payload.bookingId || (payload.data && payload.data.bookingId);
      const sid = payload.studentId || (payload.data && payload.data.studentId);
      const pin = payload.pin || (payload.data && payload.data.pin);
      result = cancelBooking(bId, sid, pin);
    } else if (action === 'checkInBooking') {
      const bId = payload.bookingId || (payload.data && payload.data.bookingId);
      const sid = payload.studentId || (payload.data && payload.data.studentId);
      const pin = payload.pin || (payload.data && payload.data.pin);
      result = checkInBooking(bId, sid, pin);
    } else if (action === 'verifyStudent') {
      result = verifyStudent(payload.studentId);
    } else if (action === 'getDayQuota') {
      result = getDayQuota(payload.date, payload.studentId);
    } else {
      result = { success: false, error: 'Unknown POST action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * เปิด Spreadsheet (รองรับทั้ง Active Sheet ของ Container-bound และ ID ของ Standalone)
 */
function getSpreadsheet() {
  // 1. ถ้าเป็น Container-bound Script ที่ผูกกับ Google Sheet โดยตรง ให้ใช้ Active Spreadsheet ก่อนเสมอ
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}

  // 2. ถ้าเป็น Standalone ให้เปิดด้วย CONFIG.spreadsheetId
  if (CONFIG.spreadsheetId && CONFIG.spreadsheetId.trim() !== '') {
    try {
      return SpreadsheetApp.openById(CONFIG.spreadsheetId.trim());
    } catch (e) {
      Logger.log('Could not open spreadsheet by ID: ' + e.toString());
    }
  }
  return null;
}

/**
 * ค้นหา Sheet ที่เก็บข้อมูลการจองอย่างยืดหยุ่น (รองรับ RoomBookings, Sheet1, การจอง, ฯลฯ)
 */
function findBookingSheet_(ss) {
  if (!ss) return null;
  // 1. ลองหาจากชื่อที่กำหนดใน CONFIG
  let sheet = ss.getSheetByName(CONFIG.sheets.roomBookings);
  if (sheet && sheet.getLastRow() > 1) return sheet;

  // 2. ลองค้นหาชื่ออื่นๆ ที่เป็นไปได้
  const candidateNames = [
    'RoomBookings', 'Room Bookings', 'Bookings', 'Booking',
    'การจอง', 'ตารางการจอง', 'ข้อมูลการจอง', 'ห้อง 7 ดาว',
    'Sheet1', 'ชีต1'
  ];
  for (const name of candidateNames) {
    const candidate = ss.getSheetByName(name);
    if (candidate && candidate.getLastRow() > 1) {
      Logger.log('✅ พบข้อมูลการจองใน Sheet: ' + name);
      return candidate;
    }
  }

  // 3. หากยังไม่พบข้อมูล ให้คืนค่า sheet เดิมหรือ Sheet แรกสุด
  return sheet || ss.getSheets()[0];
}

// ─────────────────────────────────────────────────────────────────
//  Setup Database Sheets (Auto-creates sheets and updates columns)
// ─────────────────────────────────────────────────────────────────
/**
 * setupRoomSheets()
 * รันฟังก์ชันนี้เพื่อเตรียม Sheet: RoomBookings, BookingHistory, และ StudentData
 */
function setupRoomSheets() {
  const ss = getSpreadsheet();
  if (!ss) throw new Error('ไม่พบ Google Spreadsheet กรุณาตรวจสอบ CONFIG.spreadsheetId ใน Code.gs');

  // 1. Sheet: RoomBookings
  let bookingSheet = ss.getSheetByName(CONFIG.sheets.roomBookings);
  if (!bookingSheet) {
    bookingSheet = ss.insertSheet(CONFIG.sheets.roomBookings);
    Logger.log('✅ Created sheet: ' + CONFIG.sheets.roomBookings);
  }

  const bookingHeaders = [
    'bookingId',
    'date',
    'timeSlot',
    'tableType',
    'studentIds',
    'bookedBy',
    'pin',
    'bookedAt',
    'status',
    'checkedInAt'
  ];

  if (bookingSheet.getLastRow() === 0) {
    bookingSheet.appendRow(bookingHeaders);
    const headerRange = bookingSheet.getRange(1, 1, 1, bookingHeaders.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1e293b');
    headerRange.setFontColor('#ffffff');
    headerRange.setHorizontalAlignment('center');
    bookingSheet.setFrozenRows(1);
    bookingSheet.setColumnWidth(1, 160);  // bookingId
    bookingSheet.setColumnWidth(2, 120);  // date
    bookingSheet.setColumnWidth(3, 90);   // timeSlot
    bookingSheet.setColumnWidth(4, 120);  // tableType
    bookingSheet.setColumnWidth(5, 280);  // studentIds
    bookingSheet.setColumnWidth(6, 130);  // bookedBy
    bookingSheet.setColumnWidth(7, 90);   // pin
    bookingSheet.setColumnWidth(8, 180);  // bookedAt
    bookingSheet.setColumnWidth(9, 90);   // status
    bookingSheet.setColumnWidth(10, 180); // checkedInAt
    Logger.log('✅ Added headers to RoomBookings');
  } else {
    const headers = bookingSheet.getRange(1, 1, 1, bookingSheet.getLastColumn()).getValues()[0];
    if (!headers.includes('pin')) {
      bookingSheet.insertColumnBefore(7);
      bookingSheet.getRange(1, 7).setValue('pin').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff').setHorizontalAlignment('center');
      bookingSheet.setColumnWidth(7, 90);
      Logger.log('✅ Added pin column to existing RoomBookings sheet');
    }
  }

  // 2. Sheet: BookingHistory
  let historySheet = ss.getSheetByName(CONFIG.sheets.bookingHistory);
  if (!historySheet) {
    historySheet = ss.insertSheet(CONFIG.sheets.bookingHistory);
    Logger.log('✅ Created sheet: ' + CONFIG.sheets.bookingHistory);
  }

  const historyHeaders = [
    'logId',
    'bookingId',
    'action',
    'date',
    'timeSlot',
    'tableType',
    'studentIds',
    'actorId',
    'timestamp'
  ];
  if (historySheet.getLastRow() === 0) {
    historySheet.appendRow(historyHeaders);
    const headerRange = historySheet.getRange(1, 1, 1, historyHeaders.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1e293b');
    headerRange.setFontColor('#ffffff');
    headerRange.setHorizontalAlignment('center');
    historySheet.setFrozenRows(1);
    historySheet.setColumnWidth(1, 80);   // logId
    historySheet.setColumnWidth(2, 160);  // bookingId
    historySheet.setColumnWidth(3, 90);   // action
    historySheet.setColumnWidth(4, 120);  // date
    historySheet.setColumnWidth(5, 90);   // timeSlot
    historySheet.setColumnWidth(6, 120);  // tableType
    historySheet.setColumnWidth(7, 280);  // studentIds
    historySheet.setColumnWidth(8, 130);  // actorId
    historySheet.setColumnWidth(9, 180);  // timestamp
    Logger.log('✅ Added headers to BookingHistory');
  }

  // 3. Sheet: StudentData
  const studentSheet = findStudentSheet_(ss);
  if (!studentSheet) {
    const newStudentSheet = ss.insertSheet(CONFIG.sheets.studentData);
    newStudentSheet.appendRow(['studentId', 'prefix', 'firstName', 'lastName', 'year', 'email']);
    const hr = newStudentSheet.getRange(1, 1, 1, 6);
    hr.setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff').setHorizontalAlignment('center');
    newStudentSheet.setFrozenRows(1);
    newStudentSheet.appendRow(['6531034333', 'นาย', 'ตัวอย่าง', 'นิสิตเภสัช', '4', 'student1@chula.ac.th']);
    newStudentSheet.appendRow(['6531026333', 'นางสาว', 'คู่ซ้อม', 'นิสิตเภสัช', '4', 'student2@chula.ac.th']);
    Logger.log('✅ Created StudentData sheet with sample data');
  }

  return {
    success: true,
    message: 'เตรียม Sheet RoomBookings, BookingHistory และ StudentData เรียบร้อยแล้ว'
  };
}

/**
 * findStudentSheet_(ss)
 * ค้นหา Sheet รายชื่อนิสิต (รองรับหลายชื่อ)
 */
function findStudentSheet_(ss) {
  if (!ss) return null;
  const possibleNames = ['StudentData', 'Student Data', 'Student', 'Students', 'Student_Data', 'รายชื่อนิสิต', 'นิสิต'];
  for (const name of possibleNames) {
    const s = ss.getSheetByName(name);
    if (s) return s;
  }
  return null;
}

/**
 * verifyStudent(studentId)
 * ตรวจสอบว่ารหัสนิสิตมีอยู่ในฐานข้อมูล StudentData หรือไม่ พร้อมดึงชื่อ-สกุล
 */
function verifyStudent(studentId) {
  try {
    const sidTrim = String(studentId || '').trim();
    if (!sidTrim) return { exists: false, error: 'กรุณาระบุรหัสนิสิต' };
    if (!/^\d{10}$/.test(sidTrim)) return { exists: false, error: 'รหัสนิสิตต้องเป็นตัวเลข 10 หลัก' };

    const ss = getSpreadsheet();
    if (!ss) return { exists: true, studentId: sidTrim, name: '' };

    const studentSheet = findStudentSheet_(ss);
    if (!studentSheet || studentSheet.getLastRow() <= 1) {
      return { exists: true, studentId: sidTrim, name: '' };
    }

    const data = studentSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());

    // หา Column รหัสนิสิต
    let idCol = -1;
    for (let c = 0; c < headers.length; c++) {
      const h = headers[c];
      if (h.includes('id') || h.includes('รหัส') || h.includes('student')) {
        idCol = c;
        break;
      }
    }
    if (idCol === -1) idCol = 0;

    // หา Column ชื่อ-สกุล
    let nameCols = [];
    for (let c = 0; c < headers.length; c++) {
      const h = headers[c];
      if (c !== idCol && (h.includes('name') || h.includes('ชื่อ') || h.includes('first') || h.includes('last') || h.includes('สกุล') || h.includes('prefix') || h.includes('คำนำหน้า'))) {
        nameCols.push(c);
      }
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowId = String(row[idCol] || '').trim();
      if (rowId === sidTrim) {
        let fullName = '';
        if (nameCols.length > 0) {
          fullName = nameCols.map(c => String(row[c] || '').trim()).filter(Boolean).join(' ');
        }
        return {
          exists: true,
          studentId: sidTrim,
          name: fullName
        };
      }
    }

    return {
      exists: false,
      error: `❌ ไม่พบรหัสนิสิต ${sidTrim} ในฐานข้อมูล StudentData`
    };

  } catch (e) {
    Logger.log('verifyStudent error: ' + e.toString());
    return { exists: true, studentId: studentId, name: '' };
  }
}

// ─────────────────────────────────────────────────────────────────
//  Helper Date & Time & ID Utilities
// ─────────────────────────────────────────────────────────────────
function generateBookingId_() {
  const now  = new Date();
  const y    = now.getFullYear();
  const m    = String(now.getMonth()+1).padStart(2,'0');
  const d    = String(now.getDate()).padStart(2,'0');
  const rand = String(Math.floor(Math.random() * 9000) + 1000);
  return `BK-${y}${m}${d}-${rand}`;
}

function _normalizeDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  let str = String(val).trim();
  // Case: yyyy-MM-dd
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str.substring(0, 10);
  }
  // Case: d/M/yyyy or d-M-yyyy (e.g. 20/10/2026 or 20/10/2569)
  const slashMatch = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (slashMatch) {
    let day = parseInt(slashMatch[1], 10);
    let mon = parseInt(slashMatch[2], 10);
    let yr  = parseInt(slashMatch[3], 10);
    if (yr > 2400) yr -= 543; // convert พ.ศ. to ค.ศ.
    else if (yr < 100) yr += 2000;
    return `${yr}-${String(mon).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  return str;
}

function _normalizeTimeStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Bangkok', 'HH:mm');
  }
  let str = String(val).trim();
  // Case: 8.00 or 8:00
  str = str.replace('.', ':');
  if (/^\d{1,2}:\d{2}/.test(str)) {
    const parts = str.split(':');
    return parts[0].padStart(2, '0') + ':' + parts[1].substring(0, 2);
  }
  // Case: single or double digit hour like "8" or "14"
  if (/^\d{1,2}$/.test(str)) {
    return str.padStart(2, '0') + ':00';
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, 'Asia/Bangkok', 'HH:mm');
  }
  return str;
}

function _dateStr(d) {
  if (d instanceof Date) {
    return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  return _normalizeDateStr(d);
}

// ─────────────────────────────────────────────────────────────────
//  Core Booking API
// ─────────────────────────────────────────────────────────────────
/**
 * getBookingsForWeek(weekStartDate)
 * ดึงข้อมูลการจองทั้งสัปดาห์ (รองรับทั้งการจองผ่านเว็บและพิมพ์เองใน Google Sheet)
 */
function getBookingsForWeek(weekStartDate) {
  try {
    const cleanStartDate = _normalizeDateStr(weekStartDate) || _dateStr(new Date());

    const ss = getSpreadsheet();
    if (!ss) return {};

    let sheet = findBookingSheet_(ss);
    if (!sheet) {
      setupRoomSheets();
      sheet = findBookingSheet_(ss);
    }
    if (!sheet || sheet.getLastRow() <= 1) return {};

    try {
      autoExpireBookings_();
    } catch (eExp) {
      Logger.log('autoExpireBookings_ error: ' + eExp);
    }

    const [y, m, d] = cleanStartDate.split('-').map(Number);
    // ขยายช่วงการดึงข้อมูลเป็น 28 วัน (-7 วัน ถึง +21 วัน) เพื่อให้การเลื่อนดูวันถัดไป/สัปดาห์หน้าโหลดทันที ไม่หาย
    const startDate = new Date(y, m - 1, d - 7);
    const endDate   = new Date(y, m - 1, d + 21);
    const startStr  = Utilities.formatDate(startDate, 'Asia/Bangkok', 'yyyy-MM-dd');
    const endStr    = Utilities.formatDate(endDate, 'Asia/Bangkok', 'yyyy-MM-dd');

    const data    = sheet.getDataRange().getValues();
    if (data.length <= 1) return {};

    const rawHeaders = data[0].map(h => String(h).trim().toLowerCase());
    
    // ตรวจสอบตำแหน่งคอลัมน์แบบยืดหยุ่น (รองรับทั้งภาษาไทยและอังกฤษ)
    let bIdCol = rawHeaders.findIndex(h => h.includes('bookingid') || h.includes('id') || h.includes('รหัสการจอง'));
    let dateCol = rawHeaders.findIndex(h => h === 'date' || h.includes('วัน'));
    let timeCol = rawHeaders.findIndex(h => h.includes('time') || h.includes('slot') || h.includes('เวลา'));
    let typeCol = rawHeaders.findIndex(h => h.includes('table') || h.includes('type') || h.includes('โต๊ะ') || h.includes('ห้อง'));
    let sidsCol = rawHeaders.findIndex(h => h.includes('student') || h.includes('sid') || h.includes('นิสิต') || h.includes('ผู้จอง'));
    let bookedAtColIdx = rawHeaders.findIndex(h => h.includes('bookedat') || h.includes('เวลาที่จอง'));
    let statusColIdx = rawHeaders.findIndex(h => h.includes('status') || h.includes('สถานะ'));
    let checkedInColIdx = rawHeaders.findIndex(h => h.includes('check') || h.includes('ยืนยัน'));

    if (bIdCol === -1) bIdCol = 0;
    if (dateCol === -1) dateCol = 1;
    if (timeCol === -1) timeCol = 2;
    if (typeCol === -1) typeCol = 3;
    if (sidsCol === -1) sidsCol = 4;
    if (bookedAtColIdx === -1) bookedAtColIdx = 7;
    if (statusColIdx === -1) statusColIdx = 8;
    if (checkedInColIdx === -1) checkedInColIdx = 9;

    const result  = {};

    for (let i = 1; i < data.length; i++) {
      const row         = data[i];
      let bId           = row[bIdCol] ? String(row[bIdCol]).trim() : '';
      const date        = _normalizeDateStr(row[dateCol]);
      const time        = _normalizeTimeStr(row[timeCol]);
      let tTypeRaw      = row[typeCol] ? String(row[typeCol]).trim().toLowerCase() : '';
      const sIds        = row[sidsCol] ? String(row[sidsCol]).trim() : '';
      const bookedAt    = (bookedAtColIdx !== -1 && row[bookedAtColIdx]) ? String(row[bookedAtColIdx]).trim() : '';
      let status        = (statusColIdx !== -1 && row[statusColIdx]) ? String(row[statusColIdx]).trim().toLowerCase() : 'active';
      const checkedInAt = (checkedInColIdx !== -1 && row[checkedInColIdx]) ? String(row[checkedInColIdx]).trim() : '';

      // หากผู้ใช้กรอกเองใน Google Sheet แล้วไม่ได้ใส่ bookingId ให้ auto-generate จำลอง
      if (!bId && (date || time)) {
        bId = `BK-MANUAL-${i}`;
      }

      // ข้ามเฉพาะรายการที่ตั้งใจกดยกเลิก
      if (!bId || status === 'cancelled' || !date || !time) continue;
      if (date < startStr || date > endStr) continue;

      // Normalization ประเภทโต๊ะ (รองรับภาษาไทย และ both/เหมาห้อง)
      let isOutside = false;
      let isInside  = false;
      if (tTypeRaw.includes('out') || tTypeRaw.includes('นอก')) {
        isOutside = true;
      } else if (tTypeRaw.includes('in') || tTypeRaw.includes('ใน')) {
        isInside = true;
      } else if (tTypeRaw.includes('both') || tTypeRaw.includes('ทั้ง') || tTypeRaw.includes('all') || tTypeRaw.includes('ห้อง') || tTypeRaw === '') {
        isOutside = true;
        isInside  = true;
      }

      // หากเป็นสถานะการจองส่วนกลาง / Mock exam ให้คงสถานะยืนยันไว้เสมอ
      if (status.includes('mock') || status.includes('central') || status.includes('reserved') || status.includes('confirm') || status.includes('ยืนยัน')) {
        status = 'checked_in';
      }

      const key = `${date}_${time}`;
      if (!result[key]) result[key] = { outside: null, inside: null };

      const idList = sIds ? sIds.split(',').map(s => s.trim()).filter(Boolean) : ['ส่วนกลาง / Mock Exam'];
      const entry = { bookingId: bId, ids: idList, status, bookedAt, checkedInAt };

      if (isOutside) result[key].outside = entry;
      if (isInside)  result[key].inside  = entry;
    }

    return result;

  } catch (e) {
    Logger.log('getBookingsForWeek error: ' + e.toString());
    return {};
  }
}

/**
 * getMyBookings(studentId)
 * ค้นหารายการการจองของนิสิตคนนี้
 */
function getMyBookings(studentId) {
  try {
    const sidTrim = String(studentId).trim();
    if (!sidTrim) return [];

    try {
      autoExpireBookings_();
    } catch (eExp) {}

    const ss = getSpreadsheet();
    if (!ss) return [];
    const sheet = ss.getSheetByName(CONFIG.sheets.roomBookings);
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const bookedAtColIdx  = headers.indexOf('bookedAt') !== -1 ? headers.indexOf('bookedAt') : 7;
    const statusColIdx    = headers.indexOf('status') !== -1 ? headers.indexOf('status') : 8;
    const checkedInColIdx = headers.indexOf('checkedInAt') !== -1 ? headers.indexOf('checkedInAt') : 9;
    const list = [];

    for (let i = 1; i < data.length; i++) {
      const row         = data[i];
      const bId         = row[0] ? String(row[0]).trim() : '';
      const date        = _normalizeDateStr(row[1]);
      const time        = _normalizeTimeStr(row[2]);
      const tType       = row[3] ? String(row[3]).trim() : '';
      const sIds        = row[4] ? String(row[4]).trim() : '';
      const bookedBy    = row[5] ? String(row[5]).trim() : '';
      const bookedAt    = row[bookedAtColIdx] ? String(row[bookedAtColIdx]).trim() : '';
      const status      = row[statusColIdx] ? String(row[statusColIdx]).trim() : 'active';
      const checkedInAt = row[checkedInColIdx] ? String(row[checkedInColIdx]).trim() : '';

      if (!bId || status === 'cancelled' || status === 'no_show' || !date || !time) continue;

      const idList = sIds.split(',').map(s => s.trim());
      if (idList.includes(sidTrim) || bookedBy === sidTrim) {
        list.push({
          bookingId: bId,
          date: date,
          timeSlot: time,
          tableType: tType,
          studentIds: idList,
          bookedBy: bookedBy,
          bookedAt: bookedAt,
          status: status,
          checkedInAt: checkedInAt
        });
      }
    }

    list.sort((a, b) => {
      const cmp = a.date.localeCompare(b.date);
      if (cmp !== 0) return cmp;
      return a.timeSlot.localeCompare(b.timeSlot);
    });

    return list;
  } catch (e) {
    Logger.log('getMyBookings error: ' + e.toString());
    return [];
  }
}

/**
 * getDayQuota(date, studentId)
 * ตรวจสอบจำนวนชั่วโมงที่นิสิตจองไปแล้วในวันนั้น (สูงสุด 2 ชม./คน/วัน)
 */
function getDayQuota(date, studentId) {
  try {
    const cleanDate = _normalizeDateStr(date);
    const ss = getSpreadsheet();
    if (!ss) return 0;
    const sheet = ss.getSheetByName(CONFIG.sheets.roomBookings);
    if (!sheet || sheet.getLastRow() <= 1) return 0;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusColIdx = headers.indexOf('status') !== -1 ? headers.indexOf('status') : 8;
    const sidTrim = String(studentId).trim();
    if (!sidTrim) return 0;
    let count = 0;

    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      const rDate  = _normalizeDateStr(row[1]);
      const sIds   = row[4] ? String(row[4]).trim() : '';
      const status = row[statusColIdx] ? String(row[statusColIdx]).trim() : '';

      if (status === 'cancelled' || status === 'no_show') continue;
      if (rDate !== cleanDate) continue;

      const idList = sIds.split(',').map(s => s.trim());
      if (idList.includes(sidTrim)) count++;
    }

    return count;

  } catch (e) {
    Logger.log('getDayQuota error: ' + e.toString());
    return 0;
  }
}

/**
 * createBooking(payload)
 * บันทึกการจองห้องใหม่ — พร้อมตรวจสอบ StudentData, PIN 4 หลัก และ Concurrency Script Lock
 */
function createBooking(payload) {
  const lock = LockService.getScriptLock();
  try {
    const hasLock = lock.tryLock(15000);
    if (!hasLock) {
      return { 
        success: false, 
        error: 'ขณะนี้มีผู้ใช้งานกำลังทำรายการจองพร้อมกัน กรุณาลองใหม่อีกครั้งในอีกสักครู่' 
      };
    }

    const { date, timeSlot, tableType, studentIds, bookedBy, pin } = payload;
    const cleanDate = _normalizeDateStr(date);
    const cleanTime = _normalizeTimeStr(timeSlot);
    const cleanPin  = String(pin || '').trim();

    if (!cleanDate || !cleanTime || !tableType || !studentIds || !bookedBy) {
      return { success: false, error: 'ข้อมูลไม่ครบถ้วน กรุณากรอกให้ครบ' };
    }
    if (!cleanPin || !/^\d{4}$/.test(cleanPin)) {
      return { success: false, error: 'กรุณาตั้งรหัส PIN ตัวเลข 4 หลักสำหรับยกเลิก' };
    }
    if (!Array.isArray(studentIds) || studentIds.length < 2) {
      return { success: false, error: 'ต้องระบุรหัสนิสิตทั้ง 2 คน' };
    }
    if (String(studentIds[0]).trim() === String(studentIds[1]).trim()) {
      return { success: false, error: 'รหัสนิสิตคนที่ 1 และคนที่ 2 ต้องไม่ซ้ำกัน' };
    }
    if (!['outside','inside','both'].includes(tableType)) {
      return { success: false, error: 'ประเภทโต๊ะไม่ถูกต้อง' };
    }

    const [y, m, d] = cleanDate.split('-').map(Number);
    const [slotHour] = cleanTime.split(':').map(Number);
    const slotDateTime = new Date(y, m - 1, d, slotHour + 1, 0, 0);
    if (slotDateTime <= new Date()) {
      return { success: false, error: 'ไม่สามารถจองเวลาที่ผ่านไปแล้ว' };
    }

    // เปิด Spreadsheet ครั้งเดียวสำหรับทุกการทำงาน
    const ss = getSpreadsheet();
    if (!ss) return { success: false, error: 'ไม่สามารถเชื่อมต่อ Google Sheet ได้' };
    let sheet = ss.getSheetByName(CONFIG.sheets.roomBookings);
    if (!sheet) {
      setupRoomSheets();
      sheet = ss.getSheetByName(CONFIG.sheets.roomBookings);
    }
    if (!sheet) return { success: false, error: 'ไม่พบ Sheet RoomBookings' };

    // 1. ตรวจสอบรหัสนิสิตใน StudentData Database
    for (const sid of studentIds) {
      const v = verifyStudent(sid.trim());
      if (!v.exists) {
        return { success: false, error: v.error || `ไม่พบรหัสนิสิต ${sid} ในฐานข้อมูล` };
      }
    }

    // 2. ดึงข้อมูล RoomBookings ขึ้นแรมรอบเดียวเพื่อคำนวณทั้ง Quota และเช็ค Slot ชน (In-Memory Batching)
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusColIdx = headers.indexOf('status') !== -1 ? headers.indexOf('status') : 8;

    // ตรวจ Quota ในแรม (ใช้เวลา 0.001 ms)
    const sid1Trim = String(studentIds[0]).trim();
    const sid2Trim = String(studentIds[1]).trim();
    let q1 = 0;
    let q2 = 0;

    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      const rDate  = _normalizeDateStr(row[1]);
      const rTime  = _normalizeTimeStr(row[2]);
      const rType  = row[3] ? String(row[3]).trim().toLowerCase() : '';
      const rSids  = row[4] ? String(row[4]).trim() : '';
      const status = row[statusColIdx] ? String(row[statusColIdx]).trim().toLowerCase() : '';

      if (status === 'cancelled' || status === 'no_show') continue;

      // ตรวจ Slot ชน
      if (rDate === cleanDate && rTime === cleanTime) {
        if (tableType === 'both' && (rType === 'outside' || rType === 'inside' || rType === 'both')) {
          return { 
            success: false, 
            error: 'ขออภัย! มีโต๊ะในช่วงเวลานี้ถูกนิสิตท่านอื่นจองไปแล้ว ไม่สามารถเลือกเหมาทั้งห้องได้' 
          };
        } else if (rType === tableType || rType === 'both') {
          const tName = tableType === 'outside' ? 'โต๊ะกลมด้านนอก' : 'โต๊ะในห้อง';
          return { 
            success: false, 
            error: `ขออภัย! ${tName} ในช่วงเวลานี้เพิ่งถูกนิสิตท่านอื่นจองสำเร็จไปแล้ว (มาก่อนได้ก่อน)` 
          };
        }
      }

      // ตรวจ Quota ในวันเดียวกัน
      if (rDate === cleanDate) {
        const idList = rSids.split(',').map(s => s.trim());
        if (idList.includes(sid1Trim)) q1++;
        if (idList.includes(sid2Trim)) q2++;
      }
    }

    if (q1 >= 2) return { success: false, error: `รหัสนิสิต ${sid1Trim} ใช้ quota 2 ชม./วันครบแล้ว` };
    if (q2 >= 2) return { success: false, error: `รหัสนิสิต ${sid2Trim} ใช้ quota 2 ชม./วันครบแล้ว` };

    const bookingId = generateBookingId_();
    const now       = new Date();
    const nowStr    = Utilities.formatDate(now, 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm:ss");

    const hasPinCol = headers.includes('pin');
    if (hasPinCol) {
      sheet.appendRow([
        bookingId,
        cleanDate,
        cleanTime,
        tableType,
        studentIds.join(', '),
        String(bookedBy).trim(),
        cleanPin,
        nowStr,
        'active',
        ''
      ]);
    } else {
      sheet.appendRow([
        bookingId,
        cleanDate,
        cleanTime,
        tableType,
        studentIds.join(', '),
        String(bookedBy).trim(),
        nowStr,
        'active',
        ''
      ]);
    }

    const historySheet = ss.getSheetByName(CONFIG.sheets.bookingHistory);
    if (historySheet) {
      const logId = historySheet.getLastRow();
      historySheet.appendRow([
        logId,
        bookingId,
        'created',
        cleanDate,
        cleanTime,
        tableType,
        studentIds.join(', '),
        String(bookedBy).trim(),
        nowStr
      ]);
    }

    Logger.log(`✅ Booking created: ${bookingId} | ${cleanDate} ${cleanTime} | ${tableType} | PIN set`);
    return { success: true, bookingId: bookingId };

  } catch (e) {
    Logger.log('createBooking error: ' + e.toString());
    return { success: false, error: 'เกิดข้อผิดพลาดในระบบ: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * createMockCentralBooking(payload)
 * จองห้องสำหรับสอบ Mock Exam / ส่วนกลาง (จองได้หลายชั่วโมง, เหมาห้อง, ไม่ติด Quota, ไม่ต้องยืนยัน 5 นาที)
 */
function createMockCentralBooking(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const { date, tableType, title, coordinator, pin } = payload;
    const startVal = payload.startHour !== undefined ? payload.startHour : (payload.startSlot || '');
    const endVal   = payload.endHour !== undefined ? payload.endHour : (payload.endSlot || '');

    const cleanDate  = _normalizeDateStr(date);
    const sH         = parseInt(String(startVal).split(':')[0], 10);
    const eH         = parseInt(String(endVal).split(':')[0], 10);
    const adminPin   = String(pin || '').trim();
    const eventTitle = String(title || 'สอบ Mock Exam OSPE').trim();
    const coordName  = String(coordinator || 'ฝ่ายจัดสอบส่วนกลาง').trim();

    if (!cleanDate) return { success: false, error: 'กรุณาระบุวันที่จัดสอบ' };
    if (isNaN(sH) || isNaN(eH) || sH >= eH) {
      return { success: false, error: 'ช่วงเวลาไม่ถูกต้อง (เวลาเริ่มต้องน้อยกว่าเวลาสิ้นสุด)' };
    }
    if (adminPin !== 'rxcu' && adminPin !== 'ple2569' && adminPin !== '1234' && adminPin !== 'maxnum') {
      return { success: false, error: 'รหัส Admin PIN ไม่ถูกต้อง' };
    }

    const ss = getSpreadsheet();
    if (!ss) return { success: false, error: 'ไม่สามารถเชื่อมต่อ Google Sheet ได้' };
    let sheet = ss.getSheetByName(CONFIG.sheets.roomBookings);
    if (!sheet) {
      setupRoomSheets();
      sheet = ss.getSheetByName(CONFIG.sheets.roomBookings);
    }
    if (!sheet) return { success: false, error: 'ไม่พบ Sheet RoomBookings' };

    const headers = sheet.getDataRange().getValues()[0];
    const hasPinCol = headers.includes('pin');
    const now = new Date();
    const nowStr = Utilities.formatDate(now, 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm:ss");

    const targetTables = (tableType === 'both' || !tableType) ? ['outside', 'inside'] : [tableType];
    const bookedIds = [];

    // Loop ทุกชั่วโมงตั้งแต่ sH ถึง eH - 1
    for (let h = sH; h < eH; h++) {
      const timeSlotStr = String(h).padStart(2, '0') + ':00';

      for (const tType of targetTables) {
        const bId = `MOCK-${cleanDate.replace(/-/g,'')}-${h}-${tType.toUpperCase()}`;
        bookedIds.push(bId);

        if (hasPinCol) {
          sheet.appendRow([
            bId,
            cleanDate,
            timeSlotStr,
            tType,
            eventTitle,
            coordName,
            'rxcu',
            nowStr,
            'checked_in', // ได้รับการยืนยันทันที ไม่ต้องรอนับถอยหลัง 5 นาที
            nowStr
          ]);
        } else {
          sheet.appendRow([
            bId,
            cleanDate,
            timeSlotStr,
            tType,
            eventTitle,
            coordName,
            nowStr,
            'checked_in',
            nowStr
          ]);
        }
      }
    }

    const historySheet = ss.getSheetByName(CONFIG.sheets.bookingHistory);
    if (historySheet) {
      const logId = historySheet.getLastRow();
      historySheet.appendRow([
        logId,
        bookedIds.join('; '),
        'mock_created',
        cleanDate,
        `${sH}:00-${eH}:00`,
        tableType,
        eventTitle,
        coordName,
        nowStr
      ]);
    }

    Logger.log(`🏛️ Mock Central Booking Created: ${cleanDate} (${sH}:00-${eH}:00) by ${coordName}`);
    return { success: true, count: bookedIds.length };

  } catch (e) {
    Logger.log('createMockCentralBooking error: ' + e.toString());
    return { success: false, error: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * cancelBooking(bookingId, studentId, pin)
 * ยกเลิกการจอง พร้อมตรวจสอบ PIN 4 หลัก
 */
function cancelBooking(bookingId, studentId, pin) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const bIdTrim = String(bookingId).trim().toUpperCase();
    const sidTrim = String(studentId).trim();
    const pinTrim = String(pin || '').trim();

    const ss = getSpreadsheet();
    if (!ss) return { success: false, error: 'ไม่สามารถเชื่อมต่อ Google Sheet ได้' };
    const sheet = ss.getSheetByName(CONFIG.sheets.roomBookings);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: false, error: 'ไม่พบข้อมูลการจอง' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const pinColIdx    = headers.indexOf('pin');
    const statusColIdx = headers.indexOf('status') !== -1 ? headers.indexOf('status') : 8;

    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      const rId    = row[0] ? String(row[0]).trim().toUpperCase() : '';
      const rDate  = _normalizeDateStr(row[1]);
      const rTime  = _normalizeTimeStr(row[2]);
      const rType  = row[3] ? String(row[3]).trim() : '';
      const rSids  = row[4] ? String(row[4]).trim() : '';
      const rBy    = row[5] ? String(row[5]).trim() : '';
      const rPin   = (pinColIdx !== -1 && row[pinColIdx] !== undefined) ? String(row[pinColIdx]).trim() : '';
      const status = row[statusColIdx] ? String(row[statusColIdx]).trim() : '';

      if (rId !== bIdTrim) continue;

      if (status === 'cancelled') {
        return { success: false, error: 'การจองนี้ถูกยกเลิกไปแล้ว' };
      }
      const idList = rSids.split(',').map(s => s.trim());
      if (rBy !== sidTrim && !idList.includes(sidTrim)) {
        return { success: false, error: 'รหัสนิสิตไม่ตรงกับสมาชิกในการจองนี้' };
      }

      // ตรวจสอบรหัส PIN 4 หลัก (ยอมรับ PIN ตรงกัน หรือ Master Admin Passcode 'rxcu')
      if (rPin && rPin !== '' && pinTrim !== 'rxcu' && pinTrim !== rPin) {
        return { success: false, error: '❌ รหัส PIN 4 หลักไม่ถูกต้อง ไม่สามารถยกเลิกได้' };
      }

      sheet.getRange(i + 1, statusColIdx + 1).setValue('cancelled');

      const historySheet = ss.getSheetByName(CONFIG.sheets.bookingHistory);
      if (historySheet) {
        const logId = historySheet.getLastRow();
        const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm:ss");
        historySheet.appendRow([
          logId,
          bookingId,
          'cancelled',
          rDate,
          rTime,
          rType,
          rSids,
          sidTrim,
          nowStr
        ]);
      }

      Logger.log(`🗑️  Booking cancelled: ${bookingId} by ${sidTrim} (PIN verified)`);
      return { success: true };
    }

    return { success: false, error: 'ไม่พบ Booking ID นี้ในระบบ' };

  } catch (e) {
    Logger.log('cancelBooking error: ' + e.toString());
    return { success: false, error: 'เกิดข้อผิดพลาดในระบบ: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * checkInBooking(bookingId, studentId, pin)
 * ยืนยันเข้าใช้งาน (ต้องใส่รหัส PIN 4 หลักที่ตั้งไว้ตอนจอง)
 */
function checkInBooking(bookingId, studentId, pin) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const bIdTrim = String(bookingId).trim().toUpperCase();
    const sidTrim = String(studentId).trim();
    const pinTrim = String(pin || '').trim();

    const ss = getSpreadsheet();
    if (!ss) return { success: false, error: 'ไม่สามารถเชื่อมต่อ Google Sheet ได้' };
    const sheet = ss.getSheetByName(CONFIG.sheets.roomBookings);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: false, error: 'ไม่พบข้อมูลการจอง' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const pinColIdx       = headers.indexOf('pin');
    const bookedAtColIdx  = headers.indexOf('bookedAt') !== -1 ? headers.indexOf('bookedAt') : 7;
    const statusColIdx    = headers.indexOf('status') !== -1 ? headers.indexOf('status') : 8;
    const checkedInColIdx = headers.indexOf('checkedInAt') !== -1 ? headers.indexOf('checkedInAt') : 9;
    const now  = new Date();

    for (let i = 1; i < data.length; i++) {
      const row         = data[i];
      const rId         = row[0] ? String(row[0]).trim().toUpperCase() : '';
      const rDate       = _normalizeDateStr(row[1]);
      const rTime       = _normalizeTimeStr(row[2]);
      const rType       = row[3] ? String(row[3]).trim() : '';
      const rSids       = row[4] ? String(row[4]).trim() : '';
      const rBy         = row[5] ? String(row[5]).trim() : '';
      const rPin        = (pinColIdx !== -1 && row[pinColIdx] !== undefined) ? String(row[pinColIdx]).trim() : '';
      const bookedAtStr = row[bookedAtColIdx] ? String(row[bookedAtColIdx]).trim() : '';
      const status      = row[statusColIdx] ? String(row[statusColIdx]).trim() : 'active';

      if (rId !== bIdTrim) continue;

      if (status === 'checked_in') {
        return { success: false, error: 'ยืนยันการเข้าใช้งานไปแล้ว' };
      }
      if (status !== 'active') {
        return { success: false, error: 'การจองนี้ไม่สามารถ check-in ได้ (status: ' + status + ')' };
      }

      const idList = rSids.split(',').map(s => s.trim());
      if (!idList.includes(sidTrim) && rBy !== sidTrim) {
        return { success: false, error: 'รหัสนิสิตไม่อยู่ในรายชื่อผู้จอง' };
      }

      // ตรวจสอบรหัส PIN 4 หลัก
      if (rPin && rPin !== '' && pinTrim !== 'rxcu' && pinTrim !== rPin) {
        return { success: false, error: '❌ รหัส PIN 4 หลักไม่ถูกต้อง ไม่สามารถยืนยันเข้าใช้งานได้' };
      }

      const [slotH] = rTime.split(':').map(Number);
      const [y, m, d] = rDate.split('-').map(Number);
      const slotStart = new Date(y, m - 1, d, slotH, 0, 0);
      const slotEnd   = new Date(y, m - 1, d, slotH + 1, 0, 0);

      // คำนวณ Deadline การยืนยันตัวตน (5 นาทีจากเวลาเริ่ม Slot หรือ 5 นาทีจากเวลาที่จอง)
      let windowEnd = new Date(slotStart.getTime() + 5 * 60 * 1000);
      if (bookedAtStr) {
        const bookedAtDate = new Date(bookedAtStr);
        if (!isNaN(bookedAtDate.getTime())) {
          const bookedPlus5 = new Date(bookedAtDate.getTime() + 5 * 60 * 1000);
          if (bookedPlus5 > windowEnd) {
            windowEnd = bookedPlus5;
          }
        }
      }
      if (windowEnd > slotEnd) {
        windowEnd = slotEnd;
      }

      if (now > windowEnd) {
        return { success: false, error: 'หมดเวลายืนยัน (เกิน 5 นาทีแล้ว) การจองถูกยกเลิกอัตโนมัติ' };
      }
      if (now < slotStart) {
        return { success: false, error: 'ยังไม่ถึงเวลาจอง กรุณารอจนถึงเวลา ' + rTime };
      }

      const nowStr = Utilities.formatDate(now, 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm:ss");
      sheet.getRange(i + 1, statusColIdx + 1).setValue('checked_in');
      sheet.getRange(i + 1, checkedInColIdx + 1).setValue(nowStr);

      const historySheet = ss.getSheetByName(CONFIG.sheets.bookingHistory);
      if (historySheet) {
        const logId = historySheet.getLastRow();
        historySheet.appendRow([
          logId, bookingId, 'checked_in', rDate, rTime, rType, rSids, sidTrim, nowStr
        ]);
      }

      Logger.log(`✅ Check-in: ${bookingId} by ${sidTrim} at ${nowStr} (PIN verified)`);
      return { success: true };
    }

    return { success: false, error: 'ไม่พบ Booking ID นี้ในระบบ' };

  } catch (e) {
    Logger.log('checkInBooking error: ' + e.toString());
    return { success: false, error: 'เกิดข้อผิดพลาดในระบบ: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * autoExpireBookings_()
 * ตรวจสอบและยกเลิก No-Show อัตโนมัติ (ให้เวลายืนยัน 15 นาทีจากเวลาเริ่ม Slot)
 * ป้องกันไม่ให้แตะต้องรายการจองในอนาคตโดยเด็ดขาด
 */
function autoExpireBookings_() {
  try {
    const ss = getSpreadsheet();
    if (!ss) return;
    const sheet = findBookingSheet_(ss);
    if (!sheet || sheet.getLastRow() <= 1) return;

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const statusColIdx = headers.findIndex(h => h.includes('status') || h.includes('สถานะ'));
    const bookedAtColIdx = headers.findIndex(h => h.includes('bookedat') || h.includes('เวลาที่จอง'));
    const dateColIdx = headers.findIndex(h => h === 'date' || h.includes('วัน'));
    const timeColIdx = headers.findIndex(h => h.includes('time') || h.includes('slot') || h.includes('เวลา'));
    const now = new Date();

    const actualStatusCol = statusColIdx !== -1 ? statusColIdx : 8;
    const actualBookedAtCol = bookedAtColIdx !== -1 ? bookedAtColIdx : 7;
    const actualDateCol = dateColIdx !== -1 ? dateColIdx : 1;
    const actualTimeCol = timeColIdx !== -1 ? timeColIdx : 2;

    const historySheet = ss.getSheetByName(CONFIG.sheets.bookingHistory);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const bId = row[0] ? String(row[0]).trim() : '';
      const date = _normalizeDateStr(row[actualDateCol]);
      const time = _normalizeTimeStr(row[actualTimeCol]);
      const tType = row[3] ? String(row[3]).trim() : '';
      const sIds = row[4] ? String(row[4]).trim() : '';
      let status = row[actualStatusCol] ? String(row[actualStatusCol]).trim().toLowerCase() : 'active';
      const bookedAtStr = (actualBookedAtCol !== -1 && row[actualBookedAtCol]) ? String(row[actualBookedAtCol]).trim() : '';

      // หากเป็นการจองส่วนกลาง / Mock exam หรือสถานะยืนยันแล้ว ไม่ต้อง auto-expire
      if (status.includes('mock') || status.includes('central') || status.includes('reserved') || status.includes('confirm') || status === 'checked_in' || status === 'cancelled' || status === 'no_show') {
        continue;
      }

      if (!bId || status !== 'active' || !date || !time) continue;

      const [slotH] = time.split(':').map(Number);
      if (isNaN(slotH)) continue;

      const [y, m, d] = date.split('-').map(Number);
      const slotStart = new Date(y, m - 1, d, slotH, 0, 0);
      const slotEnd   = new Date(y, m - 1, d, slotH + 1, 0, 0);

      // ห้าม Auto-expire การจองในอนาคตเด็ดขาด!
      if (now < slotStart) continue;

      let windowEnd = new Date(slotStart.getTime() + 15 * 60 * 1000);
      if (bookedAtStr) {
        const bookedAtDate = new Date(bookedAtStr);
        if (!isNaN(bookedAtDate.getTime())) {
          const bookedPlus15 = new Date(bookedAtDate.getTime() + 15 * 60 * 1000);
          if (bookedPlus15 > windowEnd) {
            windowEnd = bookedPlus15;
          }
        }
      }
      if (windowEnd > slotEnd) {
        windowEnd = slotEnd;
      }

      if (now > windowEnd) {
        sheet.getRange(i + 1, actualStatusCol + 1).setValue('no_show');

        if (historySheet) {
          const logId  = historySheet.getLastRow();
          const nowStr = Utilities.formatDate(now, 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm:ss");
          historySheet.appendRow([
            logId, bId, 'no_show', date, time, tType, sIds, 'SYSTEM', nowStr
          ]);
        }
        Logger.log(`⏰ No-show expire: ${bId} | ${date} ${time}`);
      }
    }
  } catch (e) {
    Logger.log('autoExpireBookings_ error: ' + e.toString());
  }
}

/**
 * clearTestBookings(pin)
 * ล้างข้อมูลการจองทดสอบของวันที่ 2026-10-25 หรือที่มีคำว่า 'ทดสอบ'
 */
function clearTestBookings(pin) {
  const adminPin = String(pin || '').trim();
  if (adminPin !== 'rxcu' && adminPin !== 'ple2569' && adminPin !== 'maxnum' && adminPin !== '1234') {
    return { success: false, error: 'รหัส PIN ผู้ดูแลระบบไม่ถูกต้อง' };
  }
  const ss = getSpreadsheet();
  if (!ss) return { success: false, error: 'ไม่สามารถเปิด Google Sheet ได้' };
  const sheet = ss.getSheetByName(CONFIG.sheets.roomBookings);
  if (!sheet) return { success: false, error: 'ไม่พบ Sheet RoomBookings' };

  const data = sheet.getDataRange().getValues();
  let deletedCount = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const rDate = _normalizeDateStr(row[1]);
    const rSids = String(row[4] || '');
    const rBy   = String(row[5] || '');
    if (rDate === '2026-10-25' && (rSids.includes('ทดสอบ') || rBy.includes('ส่วนกลาง') || rSids.includes('Mock Exam'))) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }
  return { success: true, deleted: deletedCount };
}
