/**
 * ทำความสะอาดและ normalize URL ของ Google Sheets
 * รองรับหลายรูปแบบ:
 *   - https://docs.google.com/spreadsheets/d/ID
 *   - https://docs.google.com/spreadsheets/d/ID/edit
 *   - https://docs.google.com/spreadsheets/d/ID/edit?gid=123#gid=123
 *   - ID ตรงๆ (ไม่มี URL)
 *
 * @param {string} url - URL หรือ ID ของ Google Sheets
 * @return {string} URL ที่ normalize แล้ว
 */
function normalizeSheetUrl(url) {
  if (!url) return url;

  url = url.trim();

  // ถ้าเป็น ID ตรงๆ (ไม่มี / หรือ http) ให้สร้าง URL เต็ม
  if (!url.includes('/') && !url.includes('http')) {
    return 'https://docs.google.com/spreadsheets/d/' + url + '/edit';
  }

  // ถ้า URL ไม่มี /edit ให้เติม /edit ต่อท้าย
  // เช่น https://docs.google.com/spreadsheets/d/ID → .../d/ID/edit
  if (url.match(/\/d\/[a-zA-Z0-9_-]+\/?$/) && !url.includes('/edit')) {
    url = url.replace(/\/?$/, '/edit');
  }

  return url;
}

/**
 * ดึง Spreadsheet ID จาก URL
 * รองรับทั้ง URL แบบ /d/ID/edit, /d/ID, หรือ ID ตรงๆ
 *
 * @param {string} url - URL ของ Google Sheets
 * @return {string} Spreadsheet ID
 */
function extractSpreadsheetId(url) {
  // ลอง match จาก URL pattern ก่อน
  var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) {
    return match[1];
  }

  // ถ้าไม่ match อาจเป็น ID ตรงๆ
  var idMatch = url.trim().match(/^[a-zA-Z0-9_-]+$/);
  if (idMatch) {
    return idMatch[0];
  }

  throw new Error('ไม่สามารถดึง Spreadsheet ID จาก URL: ' + url);
}

/**
 * ดึง GID (Sheet ID) จาก URL
 *
 * @param {string} url - URL ของ Google Sheets
 * @return {number|null} GID หรือ null ถ้าไม่มี
 */
function extractGid(url) {
  var match = url.match(/[#&?]gid=(\d+)/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * เปิด Sheet จาก URL โดยระบุ sheet ที่ถูกต้องตาม GID
 * ใช้ openById แทน openByUrl เพื่อหลีกเลี่ยงปัญหาสิทธิ์
 *
 * @param {string} sheetUrl - URL ของ Google Sheets
 * @return {GoogleAppsScript.Spreadsheet.Sheet} Sheet object
 */
function openSheetByUrl(sheetUrl) {
  // Normalize URL ก่อน — รองรับ URL แบบสั้น (ไม่มี /edit) หรือ ID ตรงๆ
  sheetUrl = normalizeSheetUrl(sheetUrl);
  Logger.log('URL หลัง normalize: ' + sheetUrl);

  var spreadsheetId = extractSpreadsheetId(sheetUrl);

  // ตรวจสอบสิทธิ์เข้าถึงไฟล์ผ่าน DriveApp ก่อน
  try {
    var file = DriveApp.getFileById(spreadsheetId);
    Logger.log('พบไฟล์: ' + file.getName() + ' (แก้ไขได้: ' + file.isEditable() + ')');
  } catch (driveError) {
    throw new Error(
      'ไม่มีสิทธิ์เข้าถึงไฟล์ (ID: ' + spreadsheetId + ') — ' +
      'กรุณาแชร์ไฟล์ให้กับ script owner หรือ service account ด้วยสิทธิ์ Editor — ' +
      'DriveApp error: ' + driveError.toString()
    );
  }

  // เปิด Spreadsheet ด้วย openById (เสถียรกว่า openByUrl)
  var spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (ssError) {
    throw new Error(
      'เปิด Spreadsheet ไม่ได้ (ID: ' + spreadsheetId + ') — ' +
      'ตรวจสอบว่า script owner มีสิทธิ์ Editor บนไฟล์นี้ — ' +
      'SpreadsheetApp error: ' + ssError.toString()
    );
  }

  // เลือก sheet ตาม GID (ถ้ามี) แทนการใช้ getActiveSheet
  var gid = extractGid(sheetUrl);
  if (gid !== null) {
    var sheets = spreadsheet.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === gid) {
        Logger.log('เลือก sheet: ' + sheets[i].getName() + ' (GID: ' + gid + ')');
        return sheets[i];
      }
    }
    Logger.log('ไม่พบ sheet ที่มี GID: ' + gid + ' — ใช้ sheet แรกแทน');
  }

  return spreadsheet.getSheets()[0];
}

/**
 * เขียนค่าลง cell โดยข้าม merge ถ้าเกิด error
 * (กรณี sheet มีการป้องกันบาง cell)
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} rangeA1 - เช่น 'E43:F43'
 * @param {string} value
 * @param {boolean} shouldMerge - ต้อง merge ก่อนเขียนหรือไม่
 */
function safeSetValue(sheet, rangeA1, value, shouldMerge) {
  try {
    var range = sheet.getRange(rangeA1);
    if (shouldMerge) {
      range.merge();
    }
    range.setValue(value);
  } catch (e) {
    Logger.log('ไม่สามารถเขียน ' + rangeA1 + ': ' + e.toString());
    // ลองเขียนโดยไม่ merge
    if (shouldMerge) {
      try {
        var firstCell = rangeA1.split(':')[0];
        sheet.getRange(firstCell).setValue(value);
        Logger.log('เขียน ' + firstCell + ' สำเร็จ (ข้าม merge)');
      } catch (e2) {
        Logger.log('ไม่สามารถเขียน ' + firstCell + ' ได้เลย: ' + e2.toString());
      }
    }
  }
}

/**
 * อัพเดทข้อมูลผู้อนุมัติลงใน Google Sheets
 * เรียกจาก AppSheet Automation
 *
 * แก้ไขปัญหา:
 * 1. ใช้ openById แทน openByUrl (เสถียรกว่าเมื่อเรียกจาก Automation)
 * 2. ระบุ sheet ตาม GID แทน getActiveSheet (ไม่มี active sheet ใน Automation)
 * 3. ตรวจสอบสิทธิ์ผ่าน DriveApp ก่อนเปิดไฟล์
 * 4. เขียนค่าแบบ safe เพื่อรองรับ protected ranges
 *
 * @param {string} sheetUrl - URL ของ Google Sheets
 * @param {string} approver1Name - ชื่อผู้อนุมัติคนที่ 1
 * @param {string} approver1Date - วันเวลาอนุมัติคนที่ 1
 * @param {string} approver2Name - ชื่อผู้อนุมัติคนที่ 2
 * @param {string} approver2Date - วันเวลาอนุมัติคนที่ 2
 * @param {string} approver3Name - ชื่อผู้อนุมัติคนที่ 3
 * @param {string} approver3Date - วันเวลาอนุมัติคนที่ 3
 * @param {string} approver4Name - ชื่อผู้อนุมัติคนที่ 4
 * @param {string} approver4Date - วันเวลาอนุมัติคนที่ 4
 * @param {string} approver5Name - ชื่อผู้อนุมัติคนที่ 5
 * @param {string} approver5Date - วันเวลาอนุมัติคนที่ 5
 * @return {string} ผลลัพธ์การทำงาน
 */
function updateApprovalFromAppSheet(
  sheetUrl,
  approver1Name, approver1Date,
  approver2Name, approver2Date,
  approver3Name, approver3Date,
  approver4Name, approver4Date,
  approver5Name, approver5Date
) {
  try {
    // ตรวจสอบ sheetUrl
    if (!sheetUrl) {
      throw new Error('sheetUrl is required');
    }

    Logger.log('เริ่มอัพเดท: ' + sheetUrl);

    // เปิด sheet ด้วยวิธีที่เสถียรกว่า
    var sheet = openSheetByUrl(sheetUrl);

    // ข้อมูลผู้อนุมัติ (ชื่อ) - แถว 43
    safeSetValue(sheet, 'E43:F43', approver5Name || '', true);
    safeSetValue(sheet, 'G43:H43', approver4Name || '', true);
    safeSetValue(sheet, 'I43',     approver3Name || '', false);
    safeSetValue(sheet, 'J43',     approver2Name || '', false);
    safeSetValue(sheet, 'K43',     approver1Name || '', false);

    // ข้อมูลวันเวลา - แถว 44
    safeSetValue(sheet, 'E44:F44', approver5Date || '', true);
    safeSetValue(sheet, 'G44:H44', approver4Date || '', true);
    safeSetValue(sheet, 'I44',     approver3Date || '', false);
    safeSetValue(sheet, 'J44',     approver2Date || '', false);
    safeSetValue(sheet, 'K44',     approver1Date || '', false);

    Logger.log('อัพเดทข้อมูลสำเร็จ: ' + sheetUrl);
    return 'Success';

  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return 'Error: ' + error.toString();
  }
}

/**
 * ฟังก์ชันทดสอบ
 */
function testUpdateFromAppSheet() {
  var result = updateApprovalFromAppSheet(
    'https://docs.google.com/spreadsheets/d/12Y8PouhDXjq36wviIZrqn5ULdZBA02vGSuymHBKNht8/edit?gid=2119072802#gid=2119072802',
    'Jansawang Satitpong', '2026-01-14 15:44:21',
    'Tansiri Natnicha', '2026-01-15 08:49:18',
    'Kumya Somrudee', '2026-01-15 08:45:08',
    'Saiek Tirachai', '2026-01-15 14:48:57',
    'Masuda Junichi', '2026-01-16 07:55:38'
  );

  Logger.log('Result: ' + result);
}

/**
 * ฟังก์ชันทดสอบแบบมีผู้อนุมัติไม่ครบ 5 คน
 */
function testUpdatePartialApprovers() {
  var result = updateApprovalFromAppSheet(
    'https://docs.google.com/spreadsheets/d/12Y8PouhDXjq36wviIZrqn5ULdZBA02vGSuymHBKNht8/edit?gid=2119072802#gid=2119072802',
    'Jansawang Satitpong', '2026-01-14 15:44:21',
    'Tansiri Natnicha', '2026-01-15 08:49:18',
    'Kumya Somrudee', '2026-01-15 08:45:08',
    '', '',
    '', ''
  );

  Logger.log('Result: ' + result);
}

/**
 * ฟังก์ชันทดสอบ URL แบบสั้น (ไม่มี /edit?gid=...)
 */
function testUpdateShortUrl() {
  var result = updateApprovalFromAppSheet(
    'https://docs.google.com/spreadsheets/d/12Y8PouhDXjq36wviIZrqn5ULdZBA02vGSuymHBKNht8',
    'Jansawang Satitpong', '2026-01-14 15:44:21',
    'Tansiri Natnicha', '2026-01-15 08:49:18',
    'Kumya Somrudee', '2026-01-15 08:45:08',
    '', '',
    '', ''
  );

  Logger.log('Result (short URL): ' + result);
}

/**
 * ฟังก์ชันทดสอบสิทธิ์เข้าถึง Spreadsheet
 * รันฟังก์ชันนี้ก่อนเพื่อตรวจสอบว่า script owner มีสิทธิ์เข้าถึงไฟล์หรือไม่
 */
function testPermission() {
  var testUrl = 'https://docs.google.com/spreadsheets/d/12Y8PouhDXjq36wviIZrqn5ULdZBA02vGSuymHBKNht8/edit?gid=2119072802#gid=2119072802';
  var spreadsheetId = extractSpreadsheetId(testUrl);

  Logger.log('Spreadsheet ID: ' + spreadsheetId);

  try {
    var file = DriveApp.getFileById(spreadsheetId);
    Logger.log('ชื่อไฟล์: ' + file.getName());
    Logger.log('สามารถแก้ไขได้: ' + file.isEditable());
    try {
      Logger.log('เจ้าของ: ' + file.getOwner().getEmail());
      Logger.log('สิทธิ์ปัจจุบัน: ' + file.getAccess(Session.getEffectiveUser()));
    } catch (permErr) {
      Logger.log('ไม่สามารถดึงข้อมูลผู้ใช้ได้ (ปกติเมื่อรันจาก Automation): ' + permErr.toString());
    }

    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheets = ss.getSheets();
    Logger.log('จำนวน sheets: ' + sheets.length);
    for (var i = 0; i < sheets.length; i++) {
      Logger.log('  Sheet: ' + sheets[i].getName() + ' (GID: ' + sheets[i].getSheetId() + ')');
    }

    Logger.log('✓ สิทธิ์เข้าถึงปกติ');
  } catch (e) {
    Logger.log('✗ ไม่มีสิทธิ์: ' + e.toString());
    Logger.log('');
    Logger.log('วิธีแก้ไข:');
    Logger.log('1. เปิด Google Sheets ที่ต้องการ');
    Logger.log('2. คลิก "Share" (แชร์)');
    Logger.log('3. เพิ่ม email ของ script owner (ดูจาก Project Settings > Owner)');
    Logger.log('4. ให้สิทธิ์ "Editor"');
  }
}
