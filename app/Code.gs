/**
 * ER.ใบลา nongmee - รองรับลาป่วย/กิจ/พักร้อน/คลอด ตามสิทธิ์ประเภทบุคลากร
 * + แจ้งเตือนเข้ากลุ่ม LINE (ผ่าน Messaging API เนื่องจาก LINE Notify ปิดบริการแล้ว)
 */

// ====== ตั้งค่าเบื้องต้น ======
const SYSTEM_TIMEZONE = 'Asia/Bangkok';
let _spreadsheetInstance = null;
let _spreadsheetTimezoneChecked = false;

function getSS() {
  // Spreadsheet ชุดใหม่ของ ER.ใบลา nongmee (ไม่ผูกกับฐานระบบเดิม)
  if (!_spreadsheetInstance) _spreadsheetInstance = SpreadsheetApp.openById('1ej5if62XHy8bKxBBSqNXzc4N3Rr1c7SYqHqfb5ePX0c');
  if (!_spreadsheetTimezoneChecked) {
    _spreadsheetTimezoneChecked = true;
    try {
      if (_spreadsheetInstance.getSpreadsheetTimeZone() !== SYSTEM_TIMEZONE) {
        _spreadsheetInstance.setSpreadsheetTimeZone(SYSTEM_TIMEZONE);
      }
    } catch (err) {
      // Continue with the fixed application timezone even if this account cannot update sheet settings.
    }
  }
  return _spreadsheetInstance;
}

function acquireWriteLock_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return null;
  return lock;
}


// จุดเดียวที่ทุกฟังก์ชันเรียกใช้ชีต — เช็ค+ซ่อมโครงสร้างอัตโนมัติก่อนเสมอ (ครั้งแรกของแต่ละคำขอเท่านั้น เร็ว)
// ทำให้ไม่ต้องไปกดรัน initializeNewSystem เองอีกเลย ระบบดูแลตัวเองได้
let _schemaEnsured = false;
function getSheetByName(name) {
  if (!_schemaEnsured) {
    _schemaEnsured = 'running';
    try {
      ensureSchema_();
      _schemaEnsured = true;
    } catch (err) {
      _schemaEnsured = false;
      throw err;
    }
  }
  return getSS().getSheetByName(name);
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ER.ใบลา nongmee')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ====================================================================
// ระบบซ่อมตัวเองอัตโนมัติ — เรียกจาก getSheetByName() ทุกครั้งที่เริ่มคำขอใหม่
// ตรวจทุกชีต/คอลัมน์ที่จำเป็น สร้าง/ซ่อมให้ถ้าขาด ไม่ต้องเข้าไปกดรันเองในหน้า Apps Script อีกเลย
// ====================================================================
function ensureSchema_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'schema_health_v20260729_4';
  if (cache.get(cacheKey)) return;
  ensureLeaveRecordsSchema_();
  ensureStaffColumns_();
  ensureFiscalYearSeedSchema_();
  ensureLeaveQuotaSeeded_();
  ensureLeaveTypeConfigSchema_();
  ensureSettingsAuditSchema_();
  ensureHolidaysSeeded_();
  ensureSettingsDefaults_();
  ensureDailyTrigger_();
  cache.put(cacheKey, 'ok', 3600);
}

// ตั้งเวลาแจ้งเตือน LINE อัตโนมัติทุกวัน 07:00 น. ให้เอง ถ้ายังไม่เคยตั้งไว้ (ไม่ต้องเข้าไปกดรัน createDailyTrigger เองอีกแล้ว)
// ห่อด้วย try-catch เพราะการสร้าง trigger ต้องใช้สิทธิ์เพิ่มเติม (script.scriptapp) ซึ่งบาง deployment อาจยังไม่ได้อนุญาตไว้
// ถ้าล้มเหลวตรงนี้ จะไม่กระทบการทำงานหลักของเว็บเลย แค่ต้องไปกดรัน createDailyTrigger เองแทนหนึ่งครั้ง
function ensureDailyTrigger_() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const exists = triggers.some(t => t.getHandlerFunction() === 'runDailyLineNotification');
    if (!exists) {
      const settings = getSettings_();
      const hour = Math.min(23, Math.max(0, Number(settings.ReminderHour) || 7));
      ScriptApp.newTrigger('runDailyLineNotification').timeBased().everyDays(1).atHour(hour).create();
    }
  } catch (err) {
    // เงียบไว้ ไม่ให้กระทบหน้าเว็บหลัก — ถ้าเกิดปัญหานี้ต้องไปกดรัน createDailyTrigger เองจาก Apps Script Editor
  }
}

// ซ่อมชีต LeaveRecords: ถ้าเป็นโครงสร้างเก่า (มีคอลัมน์ Period) หรือมีข้อมูลปนกันมั่ว
// จะเก็บของเดิมไว้เป็นสำรอง (ไม่ลบข้อมูลเด็ดขาด) แล้วสร้างชีตใหม่ที่สะอาด พร้อมย้ายข้อมูลที่กู้คืนได้เข้าไปให้
function ensureLeaveRecordsSchema_() {
  const ss = getSS();
  const sheet = ss.getSheetByName('LeaveRecords');

  if (!sheet) {
    const sh = ss.insertSheet('LeaveRecords');
    sh.getRange(1, 1, 1, 10).setValues([[
      'ID', 'StaffName', 'LeaveType', 'StartDate', 'EndDate', 'TotalDays', 'Reason', 'RecordedBy', 'Timestamp', 'HalfDayPeriod'
    ]]);
    return;
  }

  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const looksNew = header.indexOf('LeaveType') !== -1 && header.indexOf('HalfDayPeriod') !== -1;
  if (looksNew) return; // โครงสร้างถูกต้องอยู่แล้ว ไม่ต้องทำอะไร (กรณีปกติ เร็ว)

  // ต้องซ่อม: อ่านข้อมูลทั้งหมด แยกแถวเก่า/ใหม่ แล้วสร้างชีตสะอาดใหม่
  const data = sheet.getDataRange().getValues();
  const tz = SYSTEM_TIMEZONE;
  const cleanRows = [];
  const reviewNotes = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // แถวว่าง

    const col3 = row[2];
    if (LEAVE_TYPES.indexOf(col3) !== -1) {
      // แถวรูปแบบใหม่ (คอลัมน์ตรงตาม schema ใหม่อยู่แล้ว) — คัดลอกตรงๆ
      const clean = row.slice(0, 10);
      while (clean.length < 10) clean.push('');
      // แปลงวันที่ให้เป็น text เผื่อเป็น Date object
      [3, 4].forEach(idx => {
        if (Object.prototype.toString.call(clean[idx]) === '[object Date]') {
          clean[idx] = Utilities.formatDate(clean[idx], tz, 'yyyy-MM-dd');
        }
      });
      cleanRows.push(clean);
    } else if (row[3] === 'morning' || row[3] === 'afternoon') {
      // แถวรูปแบบเก่าสุด (ระบบครึ่งวันก่อนมีประเภทลา) — เดาประเภทลาจากคำในเหตุผล แล้วทำเครื่องหมายให้ตรวจสอบ
      const reasonText = String(row[6] || '');
      const guessedType = /ป่วย|หมอ|รักษา|คลินิก|โรงพยาบาล/.test(reasonText) ? 'ลาป่วย' : 'ลากิจส่วนตัว';
      let dateVal = row[2];
      if (Object.prototype.toString.call(dateVal) === '[object Date]') {
        dateVal = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
      }
      dateVal = String(dateVal);
      const period = row[3] === 'afternoon' ? 'บ่าย' : 'เช้า';
      cleanRows.push([
        row[0], row[1], guessedType, dateVal, dateVal, 0.5,
        '[ย้ายจากระบบเดิมอัตโนมัติ — ประเภทลาเป็นการเดา กรุณาตรวจสอบ/แก้ไข] ' + reasonText,
        row[4], row[5], period
      ]);
      reviewNotes.push(row[1] + ' วันที่ ' + dateVal);
    }
    // แถวรูปแบบอื่นที่ไม่เข้าเงื่อนไขทั้งสอง จะถูกข้าม (ยังอยู่ในชีตสำรองให้ตรวจสอบเองได้เสมอ)
  }

  // เก็บของเดิมไว้เป็นสำรอง ไม่ลบข้อมูลเด็ดขาด
  const archiveName = 'LeaveRecords_สำรองก่อนซ่อม_' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
  sheet.setName(archiveName);

  const fresh = ss.insertSheet('LeaveRecords');
  fresh.getRange(1, 1, 1, 10).setValues([[
    'ID', 'StaffName', 'LeaveType', 'StartDate', 'EndDate', 'TotalDays', 'Reason', 'RecordedBy', 'Timestamp', 'HalfDayPeriod'
  ]]);
  if (cleanRows.length) fresh.getRange(2, 1, cleanRows.length, 10).setValues(cleanRows);

  if (reviewNotes.length) {
    Logger.log('ย้ายรายการเก่าแบบเดา ' + reviewNotes.length + ' รายการ กรุณาตรวจสอบในหน้าเว็บ: ' + reviewNotes.join(', '));
  }
}

// เพิ่มคอลัมน์ในตาราง Staff (กลุ่มงาน / ประเภทบุคลากร / PIN ส่วนตัว) ถ้ายังไม่มี
function ensureStaffColumns_() {
  const ss = getSS();
  const staffSheet = ss.getSheetByName('Staff');
  if (!staffSheet) return; // ชีต Staff ควรมีอยู่แล้วตั้งแต่แรก ถ้าไม่มีจริงๆ ต้องสร้างเองเพราะไม่รู้โครงสร้างคอลัมน์ A-C ที่ตั้งไว้เดิม
  const header = staffSheet.getRange(1, 1, 1, Math.max(14, staffSheet.getLastColumn())).getValues()[0];
  if (header[3] !== 'JobGroup') staffSheet.getRange(1, 4).setValue('JobGroup');
  if (header[4] !== 'PersonnelType') staffSheet.getRange(1, 5).setValue('PersonnelType');
  if (header[5] !== 'PersonalPin') staffSheet.getRange(1, 6).setValue('PersonalPin');
  if (header[6] !== 'Position') staffSheet.getRange(1, 7).setValue('Position');
  if (header[7] !== 'Phone') staffSheet.getRange(1, 8).setValue('Phone');
  if (header[8] !== 'Address') staffSheet.getRange(1, 9).setValue('Address');
  if (header[9] !== 'SignatureFileId') staffSheet.getRange(1, 10).setValue('SignatureFileId');
  if (header[10] !== 'VacationCarryDays') staffSheet.getRange(1, 11).setValue('VacationCarryDays');
  if (header[11] !== 'LastLeaveSeedJson') staffSheet.getRange(1, 12).setValue('LastLeaveSeedJson');
  if (header[12] !== 'LeaveFormName') staffSheet.getRange(1, 13).setValue('LeaveFormName');
  if (header[13] !== 'Gender') staffSheet.getRange(1, 14).setValue('Gender');
  const dataRowCount = Math.max(staffSheet.getMaxRows() - 1, 1);
  staffSheet.getRange(2, 6, dataRowCount, 1).setNumberFormat('@');
  staffSheet.getRange(2, 8, dataRowCount, 1).setNumberFormat('@');
  const lastRow = staffSheet.getLastRow();
  if (lastRow > 1) {
    const rows = staffSheet.getRange(2, 2, lastRow - 1, 13).getValues();
    const inferred = rows.map(row => [normalizeGender_(row[12]) || inferGenderFromName_(row[11] || row[0])]);
    staffSheet.getRange(2, 14, inferred.length, 1).setValues(inferred);
  }
}

function normalizeGender_(value) {
  const text = String(value || '').trim().toLowerCase();
  if (text === 'ชาย' || text === 'male' || text === 'm') return 'ชาย';
  if (text === 'หญิง' || text === 'female' || text === 'f') return 'หญิง';
  return '';
}

function inferGenderFromName_(value) {
  const name = String(value || '').trim();
  if (/^(?:นาย|เด็กชาย|ด\.ช\.)/.test(name)) return 'ชาย';
  if (/^(?:ว่าที่ร้อยตรีหญิง|นางสาว|นาง|เด็กหญิง|น\.ส\.|ด\.ญ\.)/.test(name)) return 'หญิง';
  return '';
}

function familyLeaveTypeForGender_(gender) {
  gender = normalizeGender_(gender);
  if (gender === 'ชาย') return 'ลาไปช่วยเหลือภริยาที่คลอดบุตร';
  if (gender === 'หญิง') return 'ลาคลอดบุตร';
  return '';
}

function normalizeTextIdentifier_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

function writePlainText_(range, value) {
  range.setNumberFormat('@');
  range.setValue(normalizeTextIdentifier_(value));
}

function formatSystemDateTH_(isoDate) {
  const match = String(isoDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return String(isoDate || '');
  return match[3] + '/' + match[2] + '/' + (Number(match[1]) + 543);
}

function ensureFiscalYearSeedSchema_() {
  const ss = getSS();
  let sheet = ss.getSheetByName('FiscalYearSeeds');
  if (!sheet) sheet = ss.insertSheet('FiscalYearSeeds');
  const header = sheet.getRange(1, 1, 1, 6).getValues()[0];
  const expected = ['FiscalYearBE', 'StaffName', 'VacationCarryDays', 'UpdatedBy', 'UpdatedAt', 'Note'];
  let needsHeader = false;
  for (let i = 0; i < expected.length; i++) {
    if (header[i] !== expected[i]) needsHeader = true;
  }
  if (needsHeader) sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
}

// สร้างและตรวจตารางสิทธิ์มาตรฐานให้ตรงกับเอกสารอ้างอิงฉบับปัจจุบัน
function ensureLeaveQuotaSeeded_() {
  syncLeaveQuotaReference_();
}

function ensureLeaveTypeConfigSchema_() {
  const ss = getSS();
  let sheet = ss.getSheetByName('LeaveTypeConfig');
  if (!sheet) sheet = ss.insertSheet('LeaveTypeConfig');
  const headers = ['LeaveType', 'Active', 'AllowHalfDay', 'CountingMode', 'PaidDefault', 'FormKind', 'UpdatedAt', 'UpdatedBy'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const existing = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String) : [];
  const defaults = [
    ['ลาป่วย', true, true, 'working_exclude_holidays', 'ตามสิทธิ์', 'general', '', 'ระบบ'],
    ['ลากิจส่วนตัว', true, true, 'working_exclude_holidays', 'ตามสิทธิ์', 'general', '', 'ระบบ'],
    ['ลาพักผ่อน', true, true, 'working_exclude_holidays', 'ตามสิทธิ์', 'vacation', '', 'ระบบ'],
    ['ลาคลอดบุตร', true, false, 'calendar', 'ตามสิทธิ์', 'general', '', 'ระบบ'],
    ['ลาไปช่วยเหลือภริยาที่คลอดบุตร', true, false, 'working_exclude_holidays', 'ตามสิทธิ์', 'general', '', 'ระบบ']
  ];
  defaults.forEach(row => { if (existing.indexOf(row[0]) === -1) sheet.appendRow(row); });
}

function ensureSettingsAuditSchema_() {
  const ss = getSS();
  let sheet = ss.getSheetByName('SettingsAudit');
  if (!sheet) sheet = ss.insertSheet('SettingsAudit');
  sheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Category', 'Action', 'Summary', 'UpdatedBy']]);
}

function syncLeaveQuotaReference_() {
  const ss = getSS();
  let sheet = ss.getSheetByName('LeaveQuota');
  if (!sheet) sheet = ss.insertSheet('LeaveQuota');
  const header = ['PersonnelType', 'LeaveType', 'AnnualQuotaDays', 'PaidStatus', 'AccumMax', 'Note'];
  const data = sheet.getDataRange().getValues();
  const officialKeys = {};
  LEAVE_QUOTA_SEED_ROWS.forEach(row => {
    officialKeys[String(row[0]) + '\u0000' + String(row[1])] = true;
  });

  // เก็บเฉพาะรายการพิเศษที่ไม่ได้อยู่ในตารางมาตรฐาน ส่วนแถวมาตรฐานใช้ค่าจากเอกสารเสมอ
  const customRows = data.slice(1).filter(row => {
    const personnelType = String(row[0] || '').trim();
    const leaveType = String(row[1] || '').trim();
    if (!personnelType || !leaveType) return false;
    return !officialKeys[personnelType + '\u0000' + leaveType];
  }).map(row => row.slice(0, 6));
  const expectedRows = LEAVE_QUOTA_SEED_ROWS.concat(customRows);
  const currentRows = data.slice(1).filter(row => row.slice(0, 6).some(value => value !== '')).map(row => row.slice(0, 6));
  const headerMatches = header.every((value, index) => data[0] && data[0][index] === value);
  const rowsMatch = JSON.stringify(currentRows) === JSON.stringify(expectedRows);
  if (headerMatches && rowsMatch) return { changed: false, standardCount: LEAVE_QUOTA_SEED_ROWS.length, customCount: customRows.length };

  sheet.getRange(1, 1, 1, 6).setValues([header]);
  if (sheet.getMaxRows() > 1) sheet.getRange(2, 1, sheet.getMaxRows() - 1, 6).clearContent();
  if (expectedRows.length) sheet.getRange(2, 1, expectedRows.length, 6).setValues(expectedRows);
  _quotaMemoryCache = null;
  CacheService.getScriptCache().remove('quota_cache');
  return { changed: true, standardCount: LEAVE_QUOTA_SEED_ROWS.length, customCount: customRows.length };
}

// ====== ซิงก์ตารางสิทธิ์กับเอกสารทางการล่าสุด (เรียกจาก Editor ได้เมื่ออัปเดตโค้ด) ======
function upgradeLeaveQuotaData() {
  const result = syncLeaveQuotaReference_();
  return result.changed
    ? 'ปรับตารางสิทธิ์การลาตามเอกสารอ้างอิงเรียบร้อยแล้ว ' + result.standardCount +
      ' รายการ และเก็บรายการพิเศษไว้อีก ' + result.customCount + ' รายการ'
    : 'ตารางสิทธิ์การลาตรงกับเอกสารอ้างอิงอยู่แล้ว ไม่ต้องแก้ไข';
}

// สร้าง+ใส่วันหยุดนักขัตฤกษ์ปี 2569 ถ้ายังไม่มีชีต (ไม่ทับถ้ามีอยู่แล้ว)
function ensureHolidaysSeeded_() {
  const ss = getSS();
  if (ss.getSheetByName('Holidays')) return;
  const h = ss.insertSheet('Holidays');
  h.getRange(1, 1, 1, 2).setValues([['Date', 'Name']]);
  const holidayRows = [
    ['2026-01-01', 'วันขึ้นปีใหม่'], ['2026-01-02', 'วันหยุดราชการเพิ่มเป็นกรณีพิเศษ'],
    ['2026-04-13', 'วันสงกรานต์'], ['2026-04-14', 'วันสงกรานต์'], ['2026-04-15', 'วันสงกรานต์'],
    ['2026-05-01', 'วันแรงงานแห่งชาติ'], ['2026-05-04', 'วันฉัตรมงคล'], ['2026-05-13', 'วันพืชมงคล'],
    ['2026-06-01', 'ชดเชยวันวิสาขบูชา'], ['2026-06-03', 'วันเฉลิมพระชนมพรรษาสมเด็จพระนางเจ้าฯ พระบรมราชินี'],
    ['2026-07-28', 'วันเฉลิมพระชนมพรรษาพระบาทสมเด็จพระเจ้าอยู่หัว'], ['2026-07-29', 'วันอาสาฬหบูชา'], ['2026-07-30', 'วันเข้าพรรษา'],
    ['2026-08-12', 'วันแม่แห่งชาติ'], ['2026-10-13', 'วันคล้ายวันสวรรคต ร.9'], ['2026-10-23', 'วันปิยมหาราช'],
    ['2026-12-05', 'วันชาติ และวันพ่อแห่งชาติ'], ['2026-12-07', 'ชดเชยวันชาติ/วันพ่อแห่งชาติ'],
    ['2026-12-10', 'วันรัฐธรรมนูญ'], ['2026-12-31', 'วันสิ้นปี']
  ];
  h.getRange(2, 1, holidayRows.length, 2).setValues(holidayRows);
}

const DEFAULT_ANNOUNCEMENT_TITLE = 'ประกาศหน่วยงาน: แนวปฏิบัติเกี่ยวกับการลาครึ่งวัน และการขอตัวออกก่อนเวลางาน';
const DEFAULT_ANNOUNCEMENT_CONTENT = [
  'แจ้ง เจ้าหน้าที่ทุกท่าน',
  '',
  'เพื่อให้การบริหารจัดการอัตรากำลังหน้างานในการดูแลคนไข้เป็นไปด้วยความเรียบร้อย และตรวจสอบเวลาปฏิบัติงานได้อย่างถูกต้อง ขอแจ้งแนวปฏิบัติให้ทราบและถือปฏิบัติร่วมกันอย่างเคร่งครัด ดังนี้',
  '',
  '1. การลาครึ่งวันตามระบบปกติ (ลา 0.5 วัน)',
  '• เจ้าหน้าที่ต้องแจ้งหัวหน้างานให้ทราบทุกครั้งก่อนการลา',
  '• ต้องส่งใบลาในระบบ Smart Office ทุกคน โดยระบุจำนวนวันลาเป็น 0.5 วัน และเลือกช่วงเวลาให้ถูกต้อง (ครึ่งเช้า หรือ ครึ่งบ่าย)',
  '• หมายเหตุสำหรับเจ้าหน้าที่รายวัน: การลากิจส่วนตัว (รวมถึงการลาครึ่งวัน) จะไม่ได้รับค่าจ้างในวันลา โดยจะหักเงินตามสัดส่วนที่ลาจริง',
  '',
  '2. เงื่อนไขการอนุโลม (กรณีมีเหตุจำเป็นต้องออกก่อนเวลา โดยไม่ได้ลาครึ่งวัน)',
  'ในกรณีที่เจ้าหน้าที่มาปฏิบัติงานแล้ว แต่มีเหตุจำเป็นต้องขอตัวออกไปทำธุระชั่วคราว หน่วยงานมีมาตรการอนุโลมให้ภายใต้เงื่อนไขดังต่อไปนี้:',
  '• ตัวอย่างการอนุโลม: มาทำงานช่วงเช้าแล้วจำเป็นต้องออกก่อนเวลาตอน 11.00 น. หรือมาทำงานช่วงบ่ายแล้วจำเป็นต้องออกก่อนเวลาตอน 14.30 น.',
  '• กรณีพิเศษ (ไม่นับเป็นวันลา):',
  '  - ขอออกไปทำธุระตอน 11.00 น. และช่วงบ่ายกลับมาปฏิบัติงานตามปกติ',
  '  - หรือขอออกไปทำธุระช่วงบ่าย และกลับเข้ามาปฏิบัติงานต่อประมาณ 14.30 น.',
  '  - ทั้งสองกรณีนี้ไม่นับเป็นวันลาครึ่งวัน',
  '• ข้อปฏิบัติ: เจ้าหน้าที่ต้องแจ้งขออนุญาตหัวหน้างานหรือผู้ควบคุมเวรก่อนออกจากงานทุกครั้ง เพื่อบริหารจัดการอัตรากำลังปฏิบัติงานหน้างาน และให้สแกนใบหน้าเข้า-ออกตามเวลามาทำงานและเวลาเลิกงานตามปกติ (ไม่ต้องสแกนเข้า-ออกระหว่างที่ไปทำธุระ)',
  '',
  '3. มาตรการการตรวจสอบข้อมูลการสแกนใบหน้า',
  '• หัวหน้างานจะทำการตรวจสอบข้อมูลการสแกนใบหน้าทุกเช้าของวันถัดไป',
  '• กรณีลืมสแกนหน้าออก (แต่มาปฏิบัติงานจริง): เจ้าหน้าที่ต้องแจ้งเหตุผลให้หัวหน้างานทราบ โดยหัวหน้างานจะสอบถามหรือทวนสอบกับเจ้าหน้าที่หน้างานที่อยู่ร่วมเวร หรือตรวจสอบจากชื่อในบันทึกเคสคนไข้ในช่วงเวลานั้น เพื่อยืนยันการปฏิบัติงานจริง',
  '• กรณีตรวจพบว่าไม่ได้ปฏิบัติงานอยู่หน้างานจริง (และไม่มีข้อมูลยืนยันได้): เจ้าหน้าที่ท่านดังกล่าวจะต้องส่งใบลาครึ่งวัน (0.5 วัน) ย้อนหลังในระบบ Smart Office เพื่อหักวันลาหรือหักค่าจ้างตามระเบียบต่อไป',
  '',
  'จึงแจ้งมาเพื่อทราบและถือปฏิบัติร่วมกันตั้งแต่วันนี้เป็นต้นไป'
].join('\n');

// เพิ่มค่า Settings ที่จำเป็น ถ้ายังไม่มี (ไม่แตะค่าที่ผู้ใช้ตั้งไว้แล้ว)
function ensureSettingsDefaults_() {
  const ss = getSS();
  const settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) return; // Settings ต้องมีอยู่แล้วตั้งแต่แรก (มี AdminPin) ถ้าไม่มีจริงๆ ต้องให้แอดมินสร้างเองเพราะต้องตั้ง PIN เอง
  const data = settingsSheet.getDataRange().getValues();
  const existingKeys = data.map(r => r[0]);
  const toAdd = [['LineChannelAccessToken', ''], ['LineGroupId', ''], ['ReminderDaysBefore', 1],
    ['DocTemplateId_SickPersonalMaternity', ''], ['DocTemplateId_Sick', ''], ['DocTemplateId_Personal', ''],
    ['DocTemplateId_Maternity', ''], ['DocTemplateId_Vacation', ''],
    ['LeaveFormSupervisorName', ''], ['LeaveFormSupervisorPosition', ''], ['LeaveFormSupervisorSignatureFileId', ''],
    ['LeaveFormApproverName', ''], ['LeaveFormApproverPosition', ''], ['LeaveFormApproverSignatureFileId', ''],
    ['AnnouncementTitle', DEFAULT_ANNOUNCEMENT_TITLE], ['AnnouncementContent', DEFAULT_ANNOUNCEMENT_CONTENT],
    ['AnnouncementUpdatedAt', ''], ['OrganizationName', ''], ['OrganizationAddress', ''],
    ['DepartmentName', ''], ['DivisionName', ''],
    ['DocumentLocation', ''], ['DocumentRecipient', 'ผู้อำนวยการ'],
    ['DirectorFooter', 'ผู้อำนวยการโรงพยาบาลขุนตาล'], ['DefaultFiscalYearBE', ''],
    ['ReminderHour', 7], ['LeaveFormFontFamily', 'Sarabun'], ['LeaveFormFontSizePt', 12.5],
    ['LeaveFormMarginTopMm', 20], ['LeaveFormMarginRightMm', 20], ['LeaveFormMarginBottomMm', 15],
    ['LeaveFormMarginLeftMm', 25], ['LeaveFormSignatureHeightMm', 18],
    ['LeaveFormRespectText', 'ขอแสดงความนับถือ'], ['LeaveFormOrderText', 'คำสั่ง'],
    ['LeaveRightsNote', 'การลาเกินสิทธิ์ จะถูกหักค่าจ้างตามจำนวนวันที่ลา ทั้งนี้ ให้เป็นไปตามระเบียบที่เกี่ยวข้อง'],
    ['LeaveRightsDocumentUrl', '']];
  toAdd.forEach(pair => {
    if (existingKeys.indexOf(pair[0]) === -1) settingsSheet.appendRow(pair);
  });
}


// ====================================================================
// ตัวเลือกอ้างอิง (ให้ตรงกันทั้งฝั่งเซิร์ฟเวอร์และหน้าเว็บ)
// ====================================================================
const PERSONNEL_TYPES = [
  'ข้าราชการพลเรือน',
  'ลูกจ้างประจำ',
  'พนักงานราชการ',
  'พนักงานกระทรวงฯ (พกส.)',
  'ลูกจ้างชั่วคราว (รายเดือน)',
  'ลูกจ้างชั่วคราว (รายวัน)',
  'จ้างเหมาบริการ (TOR)'
];
const LEAVE_TYPES = ['ลาป่วย', 'ลากิจส่วนตัว', 'ลาพักผ่อน', 'ลาคลอดบุตร', 'ลาไปช่วยเหลือภริยาที่คลอดบุตร'];
// เว้นว่างไว้สำหรับหน่วยงานใหม่ ไม่ยกชื่อกลุ่มงานของระบบต้นแบบติดมาด้วย
const JOB_GROUPS = [];

// ====================================================================
// ตารางสิทธิ์การลาอ้างอิง (แหล่งข้อมูลเดียว ใช้ทั้งตอนสร้างชีตใหม่และตอนอัปเดตข้อมูลที่มีอยู่แล้ว)
// อ้างอิงจากเอกสาร "สิทธิการลาของเจ้าหน้าที่ที่ได้รับค่าตอบแทนระหว่างลา" ฉบับล่าสุด
// หมายเหตุ: จ้างเหมาบริการ (TOR) ไม่อยู่ในเอกสารนี้ (ไม่มีสิทธิ์ลาแบบได้ค่าตอบแทนตามระเบียบ)
// ====================================================================
const LEAVE_QUOTA_SEED_ROWS = [
  // ลาพักผ่อน
  ['ข้าราชการพลเรือน', 'ลาพักผ่อน', 10, 'ได้เงิน', 20, 'อายุงานไม่เกิน 10 ปี สะสมได้ 10 วัน / อายุงาน 10 ปีขึ้นไป สะสมได้ 20 วัน'],
  ['ลูกจ้างประจำ', 'ลาพักผ่อน', 10, 'ได้เงิน', 20, 'อายุงานไม่เกิน 10 ปี สะสมได้ 10 วัน / อายุงาน 10 ปีขึ้นไป สะสมได้ 20 วัน'],
  ['พนักงานราชการ', 'ลาพักผ่อน', 10, 'ได้เงิน', 5, 'เกิดสิทธิเมื่อครบ 6 เดือน ปีแรกทอนสิทธิตามส่วน อายุงาน 1 ปีขึ้นไปสะสมได้ 5 วัน'],
  ['พนักงานกระทรวงฯ (พกส.)', 'ลาพักผ่อน', 10, 'ได้เงิน', 5, 'เกิดสิทธิเมื่อครบ 6 เดือน อายุงาน 1 ปีขึ้นไปสะสมได้ 5 วัน'],
  ['ลูกจ้างชั่วคราว (รายเดือน)', 'ลาพักผ่อน', 10, 'ได้เงิน', '', 'เกิดสิทธิเมื่อครบ 6 เดือน'],
  ['ลูกจ้างชั่วคราว (รายวัน)', 'ลาพักผ่อน', 10, 'ได้เงิน', '', 'เกิดสิทธิเมื่อครบ 6 เดือน'],
  ['จ้างเหมาบริการ (TOR)', 'ลาพักผ่อน', 0, 'ไม่มีสิทธิ์', '', ''],

  // ลากิจส่วนตัว
  ['ข้าราชการพลเรือน', 'ลากิจส่วนตัว', 45, 'ได้เงิน', '', 'ปีแรก 15 วันทำการ / รวมกับลาป่วยห้ามเกิน 23 วันทำการ มิฉะนั้นหมดสิทธิ์เลื่อนเงินเดือน'],
  ['ลูกจ้างประจำ', 'ลากิจส่วนตัว', 45, 'ได้เงิน', '', 'ปีแรก 15 วันทำการ'],
  ['พนักงานราชการ', 'ลากิจส่วนตัว', 10, 'ได้เงิน', '', 'ปีแรกทอนสิทธิตามส่วน'],
  ['พนักงานกระทรวงฯ (พกส.)', 'ลากิจส่วนตัว', 15, 'ได้เงิน', '', 'ปีแรก 6 วันทำการ'],
  ['ลูกจ้างชั่วคราว (รายเดือน)', 'ลากิจส่วนตัว', 0, 'ไม่มีสิทธิ์', '', 'ไม่มีสิทธิ์ลากิจตามระเบียบ'],
  ['ลูกจ้างชั่วคราว (รายวัน)', 'ลากิจส่วนตัว', 0, 'ไม่มีสิทธิ์', '', 'ไม่มีสิทธิ์ลากิจตามระเบียบ (No Work No Pay)'],
  ['จ้างเหมาบริการ (TOR)', 'ลากิจส่วนตัว', 0, 'ไม่มีสิทธิ์', '', ''],

  // ลาป่วย
  ['ข้าราชการพลเรือน', 'ลาป่วย', 60, 'ได้เงิน', '', 'ห้ามเกิน 23 วันทำการเมื่อรวมกับลากิจ มิฉะนั้นหมดสิทธิ์เลื่อนเงินเดือน'],
  ['ลูกจ้างประจำ', 'ลาป่วย', 60, 'ได้เงิน', '', ''],
  ['พนักงานราชการ', 'ลาป่วย', 30, 'ได้เงิน', '', ''],
  ['พนักงานกระทรวงฯ (พกส.)', 'ลาป่วย', 45, 'ได้เงิน', '', ''],
  ['ลูกจ้างชั่วคราว (รายเดือน)', 'ลาป่วย', 15, 'ได้เงิน', '', 'เกิดสิทธิเมื่อครบ 6 เดือน / ปีแรก 8 วันทำการ'],
  ['ลูกจ้างชั่วคราว (รายวัน)', 'ลาป่วย', 15, 'ได้เงิน', '', 'เกิดสิทธิเมื่อครบ 6 เดือน / ปีแรก 8 วันทำการ'],
  ['จ้างเหมาบริการ (TOR)', 'ลาป่วย', 0, 'ไม่มีสิทธิ์', '', ''],

  // ลาคลอดบุตร
  ['ข้าราชการพลเรือน', 'ลาคลอดบุตร', 90, 'ได้เงิน', '', ''],
  ['ลูกจ้างประจำ', 'ลาคลอดบุตร', 90, 'ได้เงิน', '', ''],
  ['พนักงานราชการ', 'ลาคลอดบุตร', 90, 'ได้เงิน', '', 'ค่าตอบแทนระหว่างลาไม่เกิน 45 วัน + เงินสงเคราะห์จากประกันสังคม'],
  ['พนักงานกระทรวงฯ (พกส.)', 'ลาคลอดบุตร', 90, 'ได้เงิน', '', 'ค่าตอบแทนระหว่างลาไม่เกิน 45 วัน + เงินสงเคราะห์จากประกันสังคม'],
  ['ลูกจ้างชั่วคราว (รายเดือน)', 'ลาคลอดบุตร', 90, 'ได้เงิน', '', 'เกิดสิทธิเมื่อครบ 7 เดือน / ค่าจ้างระหว่างลาไม่เกิน 45 วัน'],
  ['ลูกจ้างชั่วคราว (รายวัน)', 'ลาคลอดบุตร', 0, 'ไม่มีสิทธิ์', '', ''],
  ['จ้างเหมาบริการ (TOR)', 'ลาคลอดบุตร', 0, 'ไม่มีสิทธิ์', '', ''],

  // ลาไปช่วยเหลือภริยาที่คลอดบุตร (ระบบเลือกให้เจ้าหน้าที่ชายจากคำนำหน้าชื่อ)
  ['ข้าราชการพลเรือน', 'ลาไปช่วยเหลือภริยาที่คลอดบุตร', 15, 'ได้เงิน', '', '15 วันทำการ'],
  ['ลูกจ้างประจำ', 'ลาไปช่วยเหลือภริยาที่คลอดบุตร', 15, 'ได้เงิน', '', '15 วันทำการ'],
  ['พนักงานราชการ', 'ลาไปช่วยเหลือภริยาที่คลอดบุตร', 0, 'ไม่มีสิทธิ์', '', ''],
  ['พนักงานกระทรวงฯ (พกส.)', 'ลาไปช่วยเหลือภริยาที่คลอดบุตร', 15, 'ได้เงิน', '', '15 วันทำการ'],
  ['ลูกจ้างชั่วคราว (รายเดือน)', 'ลาไปช่วยเหลือภริยาที่คลอดบุตร', 0, 'ไม่มีสิทธิ์', '', ''],
  ['ลูกจ้างชั่วคราว (รายวัน)', 'ลาไปช่วยเหลือภริยาที่คลอดบุตร', 0, 'ไม่มีสิทธิ์', '', ''],
  ['จ้างเหมาบริการ (TOR)', 'ลาไปช่วยเหลือภริยาที่คลอดบุตร', 0, 'ไม่มีสิทธิ์', '', '']
];

// ====================================================================
// ONE-TIME SETUP — รันครั้งเดียวจาก Apps Script Editor (เลือกฟังก์ชันนี้แล้วกด Run)
// จะ: ย้าย LeaveRecords เก่าไปเก็บสำรอง, สร้างชีตใหม่ทั้งหมด, ใส่ข้อมูลสิทธิ์การลาให้อัตโนมัติ
// ====================================================================
/**
 * สร้างฐานข้อมูลใหม่แบบสะอาดสำหรับ ER.ใบลา nongmee
 * ไม่อ่านหรือย้ายข้อมูลจาก Spreadsheet ของระบบต้นแบบ
 * ให้สร้างสเปรดชีตใหม่ ผูก Apps Script กับชีตนั้น แล้วรันฟังก์ชันนี้เพียงครั้งเดียว
 */
function initializeFreshBearLeaveSystem(confirmationToken) {
  if (confirmationToken !== 'CREATE_EMPTY_BEAR_SYSTEM') {
    throw new Error('ฟังก์ชันนี้ใช้สร้างฐานใหม่และล้างข้อมูลได้ จึงปิดการรันโดยไม่ส่งรหัสยืนยัน');
  }
  const ss = getSS();
  const sheetSpecs = [
    ['Staff', ['ID', 'Name', 'Active', 'JobGroup', 'PersonnelType', 'PersonalPin', 'Position', 'Phone', 'Address', 'SignatureFileId', 'VacationCarryDays', 'LastLeaveSeedJson', 'LeaveFormName', 'Gender']],
    ['LeaveRecords', ['ID', 'StaffName', 'LeaveType', 'StartDate', 'EndDate', 'TotalDays', 'Reason', 'RecordedBy', 'Timestamp', 'HalfDayPeriod']],
    ['Settings', ['Key', 'Value']],
    ['Holidays', ['Date', 'Name']]
  ];
  sheetSpecs.forEach(spec => {
    let sheet = ss.getSheetByName(spec[0]);
    if (!sheet) sheet = ss.insertSheet(spec[0]);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, spec[1].length).setValues([spec[1]]);
  });

  const settingsSheet = ss.getSheetByName('Settings');
  settingsSheet.getRange(2, 1, 3, 2).setValues([
    ['AdminPin', ''],
    ['SupervisorPin', ''],
    ['LeaveFormOutputFolderId', '1VBJxboQaHuWR8fOL1PqwRrHxgo_yY1BV']
  ]);
  ensureFiscalYearSeedSchema_();
  ensureLeaveTypeConfigSchema_();
  ensureSettingsAuditSchema_();
  syncLeaveQuotaReference_();
  ensureSettingsDefaults_();
  ensureStaffColumns_();
  _schemaEnsured = false;
  return 'สร้างฐานระบบ ER.ใบลา nongmee เรียบร้อยแล้ว กรุณาตั้ง AdminPin, ข้อมูลหน่วยงาน และข้อมูล ผอ. ก่อนนำไปใช้งานจริง';
}

function initializeNewSystem() {
  const ss = getSS();

  // 1) ย้าย LeaveRecords เก่า (ระบบครึ่งวัน) ไปเก็บสำรอง ไม่ให้ข้อมูลหาย
  const oldRecords = ss.getSheetByName('LeaveRecords');
  if (oldRecords) {
    const headerRow = oldRecords.getRange(1, 1, 1, oldRecords.getLastColumn()).getValues()[0];
    const isOldStructure = headerRow.indexOf('Period') !== -1; // โครงสร้างเก่ามีคอลัมน์ Period
    if (isOldStructure && !ss.getSheetByName('LeaveRecords_เก่า_สำรอง')) {
      oldRecords.setName('LeaveRecords_เก่า_สำรอง');
    }
  }

  // 2) สร้าง LeaveRecords ใหม่ (ถ้ายังไม่มี)
  if (!ss.getSheetByName('LeaveRecords')) {
    const sh = ss.insertSheet('LeaveRecords');
    sh.getRange(1, 1, 1, 10).setValues([[
      'ID', 'StaffName', 'LeaveType', 'StartDate', 'EndDate', 'TotalDays', 'Reason', 'RecordedBy', 'Timestamp', 'HalfDayPeriod'
    ]]);
  } else {
    // ชีตมีอยู่แล้ว เช็คว่ามีคอลัมน์ HalfDayPeriod (คอลัมน์ J) หรือยัง ถ้ายังไม่มีให้เพิ่ม header ให้
    const lr = ss.getSheetByName('LeaveRecords');
    const lrHeader = lr.getRange(1, 1, 1, Math.max(10, lr.getLastColumn())).getValues()[0];
    if (lrHeader[9] !== 'HalfDayPeriod') lr.getRange(1, 10).setValue('HalfDayPeriod');
  }

  // 3) เพิ่มคอลัมน์ในตาราง Staff (กลุ่มงาน / ประเภทบุคลากร / PIN ส่วนตัว) ถ้ายังไม่มี
  const staffSheet = ss.getSheetByName('Staff');
  if (staffSheet) {
    const header = staffSheet.getRange(1, 1, 1, Math.max(6, staffSheet.getLastColumn())).getValues()[0];
    if (header[3] !== 'JobGroup') staffSheet.getRange(1, 4).setValue('JobGroup');
    if (header[4] !== 'PersonnelType') staffSheet.getRange(1, 5).setValue('PersonnelType');
    if (header[5] !== 'PersonalPin') staffSheet.getRange(1, 6).setValue('PersonalPin');
  }

  // 4) สร้างชีต LeaveQuota พร้อมข้อมูลสิทธิ์การลาอ้างอิง (ถ้ายังไม่มี)
  if (!ss.getSheetByName('LeaveQuota')) {
    const q = ss.insertSheet('LeaveQuota');
    q.getRange(1, 1, 1, 6).setValues([[
      'PersonnelType', 'LeaveType', 'AnnualQuotaDays', 'PaidStatus', 'AccumMax', 'Note'
    ]]);
    q.getRange(2, 1, LEAVE_QUOTA_SEED_ROWS.length, 6).setValues(LEAVE_QUOTA_SEED_ROWS);
  }

  // 5) เพิ่มค่า Settings สำหรับ LINE (ถ้ายังไม่มี) โดยไม่แตะของเดิม
  const settingsSheet = ss.getSheetByName('Settings');
  if (settingsSheet) {
    const data = settingsSheet.getDataRange().getValues();
    const existingKeys = data.map(r => r[0]);
    const toAdd = [
      ['LineChannelAccessToken', ''],
      ['LineGroupId', ''],
      ['ReminderDaysBefore', 1]
    ];
    toAdd.forEach(pair => {
      if (existingKeys.indexOf(pair[0]) === -1) settingsSheet.appendRow(pair);
    });
  }

  // 6) สร้างชีต Holidays พร้อมวันหยุดราชการปี 2569 (ถ้ายังไม่มี) — แอดมินแก้ไข/เพิ่มเองได้ทีหลัง
  if (!ss.getSheetByName('Holidays')) {
    const h = ss.insertSheet('Holidays');
    h.getRange(1, 1, 1, 2).setValues([['Date', 'Name']]);
    const holidayRows = [
      ['2026-01-01', 'วันขึ้นปีใหม่'],
      ['2026-01-02', 'วันหยุดราชการเพิ่มเป็นกรณีพิเศษ'],
      ['2026-04-13', 'วันสงกรานต์'],
      ['2026-04-14', 'วันสงกรานต์'],
      ['2026-04-15', 'วันสงกรานต์'],
      ['2026-05-01', 'วันแรงงานแห่งชาติ'],
      ['2026-05-04', 'วันฉัตรมงคล'],
      ['2026-05-13', 'วันพืชมงคล'],
      ['2026-06-01', 'ชดเชยวันวิสาขบูชา'],
      ['2026-06-03', 'วันเฉลิมพระชนมพรรษาสมเด็จพระนางเจ้าฯ พระบรมราชินี'],
      ['2026-07-28', 'วันเฉลิมพระชนมพรรษาพระบาทสมเด็จพระเจ้าอยู่หัว'],
      ['2026-07-29', 'วันอาสาฬหบูชา'],
      ['2026-07-30', 'วันเข้าพรรษา'],
      ['2026-08-12', 'วันแม่แห่งชาติ'],
      ['2026-10-13', 'วันคล้ายวันสวรรคต ร.9'],
      ['2026-10-23', 'วันปิยมหาราช'],
      ['2026-12-05', 'วันชาติ และวันพ่อแห่งชาติ'],
      ['2026-12-07', 'ชดเชยวันชาติ/วันพ่อแห่งชาติ'],
      ['2026-12-10', 'วันรัฐธรรมนูญ'],
      ['2026-12-31', 'วันสิ้นปี']
    ];
    h.getRange(2, 1, holidayRows.length, 2).setValues(holidayRows);
  }

  return 'ตั้งค่าระบบใหม่เรียบร้อยแล้ว ลองรีเฟรชหน้าเว็บดูได้เลยค่ะ';
}

// ====================================================================
// ซ่อมชีต LeaveQuota เฉพาะกรณีว่างเปล่า (เช่น เคยสร้างชีตไว้ตั้งแต่รอบแรกแต่ข้อมูลไม่เข้า)
// รันจาก Apps Script Editor ได้ทุกเมื่อ ปลอดภัย — ถ้ามีข้อมูลอยู่แล้วจะไม่ทำอะไรเลย ไม่ทับของเดิม
// ====================================================================
function repairLeaveQuotaTable() {
  const ss = getSS();
  let q = ss.getSheetByName('LeaveQuota');
  if (!q) q = ss.insertSheet('LeaveQuota');

  const lastRow = q.getLastRow();
  if (lastRow >= 2) {
    return 'ชีต LeaveQuota มีข้อมูลอยู่แล้ว (' + (lastRow - 1) + ' แถว) ไม่ต้องซ่อม ไม่มีการเปลี่ยนแปลงใดๆ (ถ้าต้องการอัปเดตตัวเลขให้ตรงเอกสารใหม่ ให้รัน upgradeLeaveQuotaData แทน)';
  }

  q.getRange(1, 1, 1, 6).setValues([['PersonnelType', 'LeaveType', 'AnnualQuotaDays', 'PaidStatus', 'AccumMax', 'Note']]);
  q.getRange(2, 1, LEAVE_QUOTA_SEED_ROWS.length, 6).setValues(LEAVE_QUOTA_SEED_ROWS);
  CacheService.getScriptCache().remove('quota_cache');
  return 'ซ่อมชีต LeaveQuota สำเร็จ ใส่ข้อมูลสิทธิ์การลาครบ ' + LEAVE_QUOTA_SEED_ROWS.length + ' แถวแล้ว ลองรีเฟรชหน้าเว็บดูได้เลยค่ะ';
}

// ตั้งเวลาแจ้งเตือนอัตโนมัติทุกวัน (รันครั้งเดียวจาก Editor เพื่อสร้าง Trigger)
// ตั้งเวลาแจ้งเตือนอัตโนมัติทุกวัน 07:00 น. — เรียกได้ทั้งจาก Apps Script Editor (ไม่ต้องมี pin)
// และจากหน้าเว็บ (ต้องมี PIN แอดมิน) ถ้าเรียกจาก Editor โดยตรง pin จะเป็น undefined จึงข้ามการเช็คไป
function createDailyTrigger(pin) {
  if (pin !== undefined && !checkPin(pin, 'admin')) return '❌ PIN ไม่ถูกต้อง';
  try {
    const hour = Math.min(23, Math.max(0, Number(getSettings_().ReminderHour) || 7));
    createDailyTrigger_(hour);
    return '✅ ตั้งเวลาแจ้งเตือนอัตโนมัติทุกวันประมาณ ' + String(hour).padStart(2, '0') + ':00 น. เรียบร้อยแล้ว';
  } catch (err) {
    return '❌ ตั้งเวลาไม่สำเร็จ: ' + err.message + ' — ถ้าขึ้น permission/authorization ต้องไปกดรัน createDailyTrigger จาก Apps Script Editor โดยตรงแทน (จะมีหน้าต่างขออนุญาตสิทธิ์เพิ่มเติมโผล่ขึ้นมา กดอนุญาตได้เลย)';
  }
}

// เช็คว่ามีการตั้งเวลาแจ้งเตือนอัตโนมัติไว้จริงหรือไม่
function checkDailyTriggerStatus(pin) {
  if (pin !== undefined && !checkPin(pin, 'admin')) return '❌ PIN ไม่ถูกต้อง';
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const found = triggers.filter(t => t.getHandlerFunction() === 'runDailyLineNotification');
    if (found.length === 0) {
      return '❌ ยังไม่พบการตั้งเวลาแจ้งเตือนอัตโนมัติเลย กดปุ่ม "ตั้ง/ซ่อมเวลาใหม่" เพื่อสร้างใหม่';
    }
    if (found.length > 1) {
      return '⚠️ พบการตั้งเวลาซ้ำกัน ' + found.length + ' อัน — กดปุ่ม "ตั้ง/ซ่อมเวลาใหม่" อีกครั้งเพื่อล้างของซ้ำแล้วสร้างใหม่ให้เหลืออันเดียว';
    }
    const tz = SYSTEM_TIMEZONE;
    const hour = Math.min(23, Math.max(0, Number(getSettings_().ReminderHour) || 7));
    return '✅ พบการตั้งเวลาไว้แล้ว 1 อัน (ทำงานทุกวันเวลาประมาณ ' + String(hour).padStart(2, '0') + ':00 น. ตาม time zone ของโปรเจกต์: ' + tz + ')';
  } catch (err) {
    return '❌ เช็คสถานะไม่สำเร็จ: ' + err.message;
  }
}

// เรียกส่งแจ้งเตือนรายวันทันที (ทดสอบจากหน้าเว็บ) — คืนข้อความสรุปผลให้เห็นชัดว่าส่งอะไรไปบ้าง
function runDailyLineNotificationManual(pin) {
  if (!checkPin(pin, 'admin')) return '❌ PIN ไม่ถูกต้อง';
  try {
    const result = runDailyLineNotification_(true);
    if (!result.hasContent) {
      return 'ℹ️ รันสำเร็จ แต่ไม่มีข้อมูลให้แจ้งเตือนตอนนี้ (ไม่มีใครลาวันนี้ หรือเริ่มลาตามจำนวนวันที่ตั้งค่าไว้พอดี) — ไม่ใช่ error';
    }
    return result.success
      ? '✅ ส่งข้อความแจ้งเตือนเข้า LINE แล้ว (ลองเช็คในกลุ่มได้เลย)'
      : '❌ พบข้อมูลที่ต้องแจ้งเตือน แต่ส่ง LINE ไม่สำเร็จ: ' + result.message;
  } catch (err) {
    return '❌ เกิดข้อผิดพลาด: ' + err.message;
  }
}

function createDailyTrigger_(hour) {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runDailyLineNotification') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyLineNotification').timeBased().everyDays(1).atHour(hour).create();
}

// ====================================================================
// CACHE
// ====================================================================
const CACHE_TTL_SECONDS = 120;
let _settingsMemoryCache = null;
let _staffMemoryCache = null;
let _quotaMemoryCache = null;
let _leaveTypeMemoryCache = null;

function getSettings_() {
  if (_settingsMemoryCache) return _settingsMemoryCache;
  const cache = CacheService.getScriptCache();
  const cached = cache.get('settings_cache');
  if (cached) {
    _settingsMemoryCache = JSON.parse(cached);
    return _settingsMemoryCache;
  }
  const sheet = getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) settings[data[i][0]] = data[i][1];
  cache.put('settings_cache', JSON.stringify(settings), CACHE_TTL_SECONDS);
  _settingsMemoryCache = settings;
  return _settingsMemoryCache;
}
function clearSettingsCache_() {
  _settingsMemoryCache = null;
  CacheService.getScriptCache().remove('settings_cache');
}

function upsertSettings_(updates) {
  const sheet = getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  Object.keys(updates).forEach(key => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === key) { sheet.getRange(i + 1, 2).setValue(updates[key]); found = true; break; }
    }
    if (!found) sheet.appendRow([key, updates[key]]);
  });
  clearSettingsCache_();
}

function auditSettingChange_(category, action, summary, updatedBy) {
  getSheetByName('SettingsAudit').appendRow([new Date(), category || '', action || '', summary || '', updatedBy || 'แอดมิน']);
}

function getRecentSettingsAudit_(limit) {
  const data = getSheetByName('SettingsAudit').getDataRange().getValues();
  const tz = SYSTEM_TIMEZONE;
  return data.slice(1).reverse().slice(0, Number(limit) || 20).map(r => ({
    timestamp: Object.prototype.toString.call(r[0]) === '[object Date]' ? Utilities.formatDate(r[0], tz, 'dd/MM/yyyy HH:mm') : String(r[0] || ''),
    category: String(r[1] || ''), action: String(r[2] || ''), summary: String(r[3] || ''), updatedBy: String(r[4] || '')
  }));
}

function getAnnouncement() {
  const settings = getSettings_();
  return {
    title: String(settings.AnnouncementTitle || DEFAULT_ANNOUNCEMENT_TITLE),
    content: String(settings.AnnouncementContent || DEFAULT_ANNOUNCEMENT_CONTENT),
    updatedAt: String(settings.AnnouncementUpdatedAt || '')
  };
}

function updateAnnouncement(title, content, pin) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  title = String(title || '').trim();
  content = String(content || '').trim();
  if (!title || !content) return { success: false, message: 'กรุณากรอกหัวข้อและเนื้อหาประกาศให้ครบ' };

  const updatedAt = Utilities.formatDate(new Date(), SYSTEM_TIMEZONE, 'dd/MM/yyyy HH:mm น.');
  const updates = { AnnouncementTitle: title, AnnouncementContent: content, AnnouncementUpdatedAt: updatedAt };
  const sheet = getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  Object.keys(updates).forEach(key => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(updates[key]);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([key, updates[key]]);
  });
  clearSettingsCache_();
  return { success: true, message: 'บันทึกประกาศสำเร็จ', title: title, content: content, updatedAt: updatedAt };
}

function parseLastLeaveSeed_(value) {
  if (!value) return {};
  try {
    const parsed = JSON.parse(String(value));
    return parsed && typeof parsed === 'object' ? normalizeLastLeaveSeed_(parsed) : {};
  } catch (err) {
    return {};
  }
}

function normalizeLastLeaveSeed_(seed) {
  seed = seed || {};
  ['sick', 'personal', 'vacation', 'maternity'].forEach(key => {
    const item = seed[key] || {};
    seed[key] = {
      startDate: normalizeDateString_(item.startDate || ''),
      endDate: normalizeDateString_(item.endDate || ''),
      totalDays: Number(item.totalDays) || 0
    };
    if (!seed[key].startDate && !seed[key].endDate && !seed[key].totalDays) delete seed[key];
  });
  return seed;
}

function lastLeaveSeedForType_(staff, leaveType) {
  const seed = staff.lastLeaveSeed || {};
  let key = '';
  if (leaveType === '\u0e25\u0e32\u0e1b\u0e48\u0e27\u0e22') key = 'sick';
  else if (leaveType === '\u0e25\u0e32\u0e01\u0e34\u0e08\u0e2a\u0e48\u0e27\u0e19\u0e15\u0e31\u0e27') key = 'personal';
  else if (leaveType === '\u0e25\u0e32\u0e1e\u0e31\u0e01\u0e1c\u0e48\u0e2d\u0e19') key = 'vacation';
  else if (leaveType === '\u0e25\u0e32\u0e04\u0e25\u0e2d\u0e14\u0e1a\u0e38\u0e15\u0e23') key = 'maternity';
  const item = key ? seed[key] : null;
  if (!item || !item.startDate || !item.endDate || !item.totalDays) return null;
  return { start: item.startDate, end: item.endDate, days: Number(item.totalDays) || 0, seeded: true };
}

function getCurrentFiscalYearBE_() {
  const tz = SYSTEM_TIMEZONE;
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return fiscalYearBEFromDate_(today);
}

function getDocumentSettings_() {
  const s = getSettings_();
  return {
    organizationName: s.OrganizationName || '', organizationAddress: s.OrganizationAddress || '',
    departmentName: s.DepartmentName || '', divisionName: s.DivisionName || '',
    documentLocation: s.DocumentLocation || s.OrganizationName || '',
    documentRecipient: s.DocumentRecipient || 'ผู้อำนวยการ',
    directorFooter: (!s.DirectorFooter || s.DirectorFooter === 'ผู้อำนวยการ') ? 'ผู้อำนวยการโรงพยาบาลขุนตาล' : s.DirectorFooter,
    fontFamily: s.LeaveFormFontFamily || 'Sarabun', fontSizePt: Number(s.LeaveFormFontSizePt) || 12.5,
    marginTopMm: Number(s.LeaveFormMarginTopMm) || 20, marginRightMm: Number(s.LeaveFormMarginRightMm) || 20,
    marginBottomMm: Number(s.LeaveFormMarginBottomMm) || 15, marginLeftMm: Number(s.LeaveFormMarginLeftMm) || 25,
    signatureHeightMm: Number(s.LeaveFormSignatureHeightMm) || 18,
    respectText: s.LeaveFormRespectText || 'ขอแสดงความนับถือ', orderText: s.LeaveFormOrderText || 'คำสั่ง',
    leaveRightsNote: s.LeaveRightsNote || 'การลาเกินสิทธิ์ จะถูกหักค่าจ้างตามจำนวนวันที่ลา ทั้งนี้ ให้เป็นไปตามระเบียบที่เกี่ยวข้อง',
    leaveRightsDocumentUrl: s.LeaveRightsDocumentUrl || ''
  };
}

function fiscalYearBEFromDate_(dateStr) {
  const normalized = normalizeDateString_(dateStr);
  const year = Number(normalized.substring(0, 4));
  const month = Number(normalized.substring(5, 7));
  return (month >= 10 ? year + 1 : year) + 543;
}

function clearFiscalSeedCache_() {
  CacheService.getScriptCache().remove('fiscal_seed_cache_v1');
}

function getFiscalSeedRows_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('fiscal_seed_cache_v1');
  if (cached) return JSON.parse(cached);
  const sheet = getSheetByName('FiscalYearSeeds');
  const data = sheet.getDataRange().getValues();
  const tz = SYSTEM_TIMEZONE;
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || !data[i][1]) continue;
    let updatedAt = data[i][4] || '';
    if (Object.prototype.toString.call(updatedAt) === '[object Date]') {
      updatedAt = Utilities.formatDate(updatedAt, tz, 'yyyy-MM-dd HH:mm');
    }
    rows.push({
      fiscalYearBE: Number(data[i][0]) || 0,
      staffName: data[i][1],
      vacationCarryDays: Number(data[i][2]) || 0,
      updatedBy: data[i][3] || '',
      updatedAt: updatedAt,
      note: data[i][5] || ''
    });
  }
  cache.put('fiscal_seed_cache_v1', JSON.stringify(rows), CACHE_TTL_SECONDS);
  return rows;
}

function getVacationCarryDaysFor_(staff, fiscalYearBE) {
  if (!staff || !staff.name) return 0;
  if (!getVacationAccumulationEligibility_(staff).eligible) return 0;
  const rows = getFiscalSeedRows_();
  const found = rows.find(r => r.fiscalYearBE === Number(fiscalYearBE) && r.staffName === staff.name);
  return found ? Number(found.vacationCarryDays) || 0 : 0;
}

function getVacationCarrySeedInfo_(staffName, fiscalYearBE) {
  const rows = getFiscalSeedRows_();
  return rows.find(r => r.fiscalYearBE === Number(fiscalYearBE) && r.staffName === staffName) || null;
}

function getVacationAccumulationEligibility_(staff, quotaTable) {
  if (!staff || !staff.personnelType) {
    return { configured: false, eligible: false, accumMax: 0, reason: 'ยังไม่ได้กำหนดประเภทบุคลากร' };
  }
  const quota = (quotaTable || getLeaveQuotaTable_())
    .find(row => row.personnelType === staff.personnelType && row.leaveType === 'ลาพักผ่อน') || null;
  if (!quota) {
    return { configured: false, eligible: false, accumMax: 0, reason: 'ยังไม่พบข้อมูลสิทธิ์ลาพักผ่อน' };
  }
  const accumMax = Number(quota.accumMax);
  const eligible = Number(quota.annualQuotaDays) > 0 && Number.isFinite(accumMax) && accumMax > 0;
  return {
    configured: true,
    eligible: eligible,
    accumMax: eligible ? accumMax : 0,
    reason: eligible ? ('สะสมได้สูงสุด ' + accumMax + ' วัน') : 'ประเภทบุคลากรนี้ไม่มีสิทธิ์สะสมวันลาพักผ่อน'
  };
}

function getFiscalYearSetupStatus(fiscalYearBE, prefetchedStaff, prefetchedQuotaTable) {
  const year = Number(fiscalYearBE) || getCurrentFiscalYearBE_();
  const seeds = getFiscalSeedRows_();
  const seededNames = {};
  seeds.forEach(row => {
    if (Number(row.fiscalYearBE) === year) seededNames[String(row.staffName || '').trim()] = true;
  });
  const activeStaff = (prefetchedStaff || getStaffRaw_()).filter(staff => staff.active);
  const quotaTable = prefetchedQuotaTable || getLeaveQuotaTable_();
  const eligibleStaff = activeStaff.filter(staff => getVacationAccumulationEligibility_(staff, quotaTable).eligible);
  const missingStaffNames = eligibleStaff
    .map(staff => String(staff.name || '').trim())
    .filter(name => name && !seededNames[name]);
  const unconfiguredStaffNames = activeStaff
    .filter(staff => !getVacationAccumulationEligibility_(staff, quotaTable).configured)
    .map(staff => String(staff.name || '').trim());
  return {
    fiscalYearBE: year,
    complete: missingStaffNames.length === 0 && unconfiguredStaffNames.length === 0,
    missingCount: missingStaffNames.length,
    missingStaffNames: missingStaffNames,
    eligibleCount: eligibleStaff.length,
    noAccumulationCount: activeStaff.length - eligibleStaff.length - unconfiguredStaffNames.length,
    unconfiguredStaffNames: unconfiguredStaffNames
  };
}

function upsertVacationCarrySeed_(staffName, fiscalYearBE, vacationCarryDays, updatedBy, note) {
  const staff = getStaffRaw_().find(row => row.name === staffName);
  const eligibility = getVacationAccumulationEligibility_(staff);
  if (!eligibility.eligible) return;
  const sheet = getSheetByName('FiscalYearSeeds');
  const data = sheet.getDataRange().getValues();
  const fy = Number(fiscalYearBE) || getCurrentFiscalYearBE_();
  const safeCarryDays = Math.min(Math.max(0, Number(vacationCarryDays) || 0), eligibility.accumMax);
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === fy && data[i][1] === staffName) {
      sheet.getRange(i + 1, 3, 1, 4).setValues([[safeCarryDays, updatedBy || '', new Date(), note || '']]);
      clearFiscalSeedCache_();
      return;
    }
  }
  sheet.appendRow([fy, staffName, safeCarryDays, updatedBy || '', new Date(), note || '']);
  clearFiscalSeedCache_();
}

function getStaffRaw_() {
  if (_staffMemoryCache) return _staffMemoryCache;
  const cache = CacheService.getScriptCache();
  const cached = cache.get('staff_cache_v9');
  if (cached) {
    _staffMemoryCache = JSON.parse(cached);
    return _staffMemoryCache;
  }
  const sheet = getSheetByName('Staff');
  const staffRange = sheet.getDataRange();
  const data = staffRange.getValues();
  const displayData = staffRange.getDisplayValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    rows.push({
      id: data[i][0],
      name: data[i][1],
      active: (data[i][2] === true || data[i][2] === 'TRUE'),
      jobGroup: data[i][3] || '',
      personnelType: data[i][4] || '',
      personalPin: normalizeTextIdentifier_(displayData[i][5]),
      position: data[i][6] || '',
      phone: normalizeTextIdentifier_(displayData[i][7]),
      address: data[i][8] || '',
      signatureFileId: data[i][9] || '',
      vacationCarryDays: Number(data[i][10]) || 0,
      lastLeaveSeed: parseLastLeaveSeed_(data[i][11] || ''),
      leaveFormName: data[i][12] || String(data[i][1] || '').replace(/^(?:ก\.ภ\.|ผ\.ช\.)\s*/, ''),
      gender: normalizeGender_(data[i][13]) || inferGenderFromName_(data[i][12] || data[i][1])
    });
  }
  cache.put('staff_cache_v9', JSON.stringify(rows), CACHE_TTL_SECONDS);
  _staffMemoryCache = rows;
  return _staffMemoryCache;
}
function clearStaffCache_() {
  _staffMemoryCache = null;
  const cache = CacheService.getScriptCache();
  cache.remove('staff_cache_v4');
  cache.remove('staff_cache_v7');
  cache.remove('staff_cache_v8');
  cache.remove('staff_cache_v9');
}

// ====== วันหยุดนักขัตฤกษ์ (ใช้ตัดออกจากการนับวันทำการ) ======
function getHolidaySet_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('holiday_cache');
  if (cached) return new Set(JSON.parse(cached));
  const sheet = getSheetByName('Holidays');
  const data = sheet.getDataRange().getValues();
  const tz = SYSTEM_TIMEZONE;
  const dates = [];
  for (let i = 1; i < data.length; i++) {
    let d = data[i][0];
    if (Object.prototype.toString.call(d) === '[object Date]') d = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    dates.push(String(d));
  }
  cache.put('holiday_cache', JSON.stringify(dates), CACHE_TTL_SECONDS);
  return new Set(dates);
}
function clearHolidayCache_() { CacheService.getScriptCache().remove('holiday_cache'); }

// ดูรายการวันหยุดของปีที่ระบุ (ไม่ต้องใช้ PIN แค่ดูข้อมูล)
function getHolidaysForYear(yearStr) {
  const sheet = getSheetByName('Holidays');
  const data = sheet.getDataRange().getValues();
  const tz = SYSTEM_TIMEZONE;
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    let d = data[i][0];
    if (Object.prototype.toString.call(d) === '[object Date]') d = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    d = String(d);
    if (d.substring(0, 4) === yearStr) rows.push({ date: d, name: data[i][1] || '' });
  }
  rows.sort((a, b) => a.date < b.date ? -1 : (a.date > b.date ? 1 : 0));
  return rows;
}

function addHoliday(dateStr, name, pin) {
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  if (!dateStr) return { success: false, message: 'กรุณาระบุวันที่' };
  const sheet = getSheetByName('Holidays');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    let d = data[i][0];
    if (Object.prototype.toString.call(d) === '[object Date]') d = Utilities.formatDate(d, SYSTEM_TIMEZONE, 'yyyy-MM-dd');
    if (String(d) === String(dateStr)) return { success: false, message: 'มีวันหยุดนี้อยู่แล้ว' };
  }
  sheet.appendRow([dateStr, name || '']);
  clearHolidayCache_();
  return { success: true, message: 'เพิ่มวันหยุดสำเร็จ' };
}

function deleteHoliday(dateStr, pin) {
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังลบข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  const sheet = getSheetByName('Holidays');
  const data = sheet.getDataRange().getValues();
  const tz = SYSTEM_TIMEZONE;
  for (let i = 1; i < data.length; i++) {
    let d = data[i][0];
    if (Object.prototype.toString.call(d) === '[object Date]') d = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    if (String(d) === String(dateStr)) {
      sheet.deleteRow(i + 1);
      clearHolidayCache_();
      return { success: true, message: 'ลบวันหยุดสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบวันหยุดนี้' };
}

function getLeaveQuotaTable_() {
  if (_quotaMemoryCache) return _quotaMemoryCache;
  const cache = CacheService.getScriptCache();
  const cached = cache.get('quota_cache');
  if (cached) {
    _quotaMemoryCache = JSON.parse(cached);
    return _quotaMemoryCache;
  }
  const sheet = getSheetByName('LeaveQuota');
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    rows.push({
      personnelType: data[i][0],
      leaveType: data[i][1],
      annualQuotaDays: Number(data[i][2]) || 0,
      paidStatus: data[i][3] || '',
      accumMax: data[i][4],
      note: data[i][5] || ''
    });
  }
  cache.put('quota_cache', JSON.stringify(rows), CACHE_TTL_SECONDS);
  _quotaMemoryCache = rows;
  return _quotaMemoryCache;
}

function importHolidayRows(rows, pin) {
  const lock = acquireWriteLock_();
  if (!lock) return { success: false, message: 'ระบบกำลังบันทึกข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  try {
    if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
    if (!Array.isArray(rows) || !rows.length) return { success: false, message: 'ไม่มีรายการวันหยุดให้นำเข้า' };
    if (rows.length > 1000) return { success: false, message: 'นำเข้าได้ครั้งละไม่เกิน 1,000 รายการ' };
    const sheet = getSheetByName('Holidays');
    const data = sheet.getDataRange().getValues();
    const existing = new Set();
    for (let i = 1; i < data.length; i++) {
      let d = data[i][0];
      if (Object.prototype.toString.call(d) === '[object Date]') d = Utilities.formatDate(d, SYSTEM_TIMEZONE, 'yyyy-MM-dd');
      existing.add(String(d));
    }
    const accepted = [], rejected = [], seen = new Set();
    rows.forEach((row, index) => {
      const date = String(row && row.date || '').trim();
      const name = String(row && row.name || '').trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !name) { rejected.push(index + 1); return; }
      const parsed = new Date(date + 'T00:00:00');
      if (isNaN(parsed.getTime()) || Utilities.formatDate(parsed, SYSTEM_TIMEZONE, 'yyyy-MM-dd') !== date) { rejected.push(index + 1); return; }
      if (existing.has(date) || seen.has(date)) return;
      seen.add(date); accepted.push([date, name]);
    });
    if (accepted.length) sheet.getRange(sheet.getLastRow() + 1, 1, accepted.length, 2).setValues(accepted);
    clearHolidayCache_();
    const duplicateCount = rows.length - accepted.length - rejected.length;
    return {
      success: true, added: accepted.length, duplicateCount: duplicateCount, rejectedCount: rejected.length,
      message: 'นำเข้าสำเร็จ ' + accepted.length + ' รายการ' + (duplicateCount ? ' · ข้ามวันที่ซ้ำ ' + duplicateCount + ' รายการ' : '') + (rejected.length ? ' · ข้ามข้อมูลไม่สมบูรณ์ ' + rejected.length + ' รายการ' : '')
    };
  } finally { lock.releaseLock(); }
}
function clearQuotaCache_() {
  _quotaMemoryCache = null;
  CacheService.getScriptCache().remove('quota_cache');
}

function getLeaveTypeConfigs_() {
  if (_leaveTypeMemoryCache) return _leaveTypeMemoryCache;
  const cache = CacheService.getScriptCache();
  const cached = cache.get('leave_type_config_cache_v1');
  if (cached) {
    _leaveTypeMemoryCache = JSON.parse(cached);
    return _leaveTypeMemoryCache;
  }
  const data = getSheetByName('LeaveTypeConfig').getDataRange().getValues();
  const rows = data.slice(1).filter(r => String(r[0] || '').trim()).map(r => ({
    leaveType: String(r[0]).trim(), active: r[1] === true || String(r[1]).toLowerCase() === 'true',
    allowHalfDay: r[2] === true || String(r[2]).toLowerCase() === 'true',
    countingMode: String(r[3] || 'working_exclude_holidays'), paidDefault: String(r[4] || ''),
    formKind: String(r[5] || 'general'), updatedAt: String(r[6] || ''), updatedBy: String(r[7] || '')
  }));
  cache.put('leave_type_config_cache_v1', JSON.stringify(rows), CACHE_TTL_SECONDS);
  _leaveTypeMemoryCache = rows;
  return _leaveTypeMemoryCache;
}
function clearLeaveTypeConfigCache_() {
  _leaveTypeMemoryCache = null;
  CacheService.getScriptCache().remove('leave_type_config_cache_v1');
}
function getActiveLeaveTypes_() { return getLeaveTypeConfigs_().filter(r => r.active).map(r => r.leaveType); }
function getLeaveTypeConfig_(leaveType) { return getLeaveTypeConfigs_().find(r => r.leaveType === leaveType) || null; }
function getQuotaFor_(personnelType, leaveType) {
  const table = getLeaveQuotaTable_();
  return table.find(r => r.personnelType === personnelType && r.leaveType === leaveType) || null;
}

// ====== ตรวจสอบ PIN ======
function checkPin(pin, role) {
  const settings = getSettings_();
  const inputPin = String(pin).trim();
  if (!inputPin) return false;
  if (role === 'admin') {
    const adminPin = String(settings['AdminPin'] || '').trim();
    return !!adminPin && adminPin === inputPin;
  }
  if (role === 'supervisor') {
    const supervisorPin = String(settings['SupervisorPin'] || '').trim();
    const adminPin = String(settings['AdminPin'] || '').trim();
    return (!!supervisorPin && supervisorPin === inputPin) ||
           (!!adminPin && adminPin === inputPin);
  }
  return false;
}

function verifyAdminAccess(pin) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN แอดมินไม่ถูกต้อง' };
  return { success: true, message: 'ยืนยันสิทธิ์แอดมินแล้ว' };
}

// ตรวจสอบ PIN ส่วนตัวของเจ้าหน้าที่คนนั้นๆ (ให้บันทึกลาของตัวเองได้โดยไม่ต้องพึ่งแอดมิน/หัวหน้างาน)
function checkStaffPin_(staffName, pin) {
  const staff = getStaffRaw_().find(s => s.name === staffName && s.active);
  if (!staff || !staff.personalPin) return false;
  return String(staff.personalPin).trim() === String(pin).trim();
}

// ====== ตัวเลือกอ้างอิงสำหรับหน้าเว็บ (dropdown ต่างๆ) ======
function getReferenceData(prefetchedLeaveTypeConfigs, prefetchedQuotaTable) {
  const leaveTypeConfigs = prefetchedLeaveTypeConfigs || getLeaveTypeConfigs_();
  const staffGroups = getStaffRaw_().map(s => String(s.jobGroup || '').trim()).filter(Boolean);
  const jobGroups = Array.from(new Set(JOB_GROUPS.concat(staffGroups)));
  return {
    personnelTypes: PERSONNEL_TYPES,
    leaveTypes: leaveTypeConfigs.filter(r => r.active).map(r => r.leaveType),
    leaveTypeConfigs: leaveTypeConfigs,
    jobGroups: jobGroups,
    quotaTable: prefetchedQuotaTable || getLeaveQuotaTable_()
  };
}

// ====== รายชื่อเจ้าหน้าที่ (ไม่ส่ง PIN ส่วนตัวออกไปฝั่งเว็บเด็ดขาด เพื่อความปลอดภัย) ======
function getActiveStaffList() {
  return getStaffRaw_().filter(s => s.active).map(s => ({
    id: s.id, name: s.name, leaveFormName: s.leaveFormName, jobGroup: s.jobGroup,
    personnelType: s.personnelType, position: s.position, gender: s.gender
  }));
}
function getAllStaffList(pin) {
  if (!checkPin(pin, 'admin')) throw new Error('PIN ไม่ถูกต้อง');
  const currentFiscalYearBE = getCurrentFiscalYearBE_();
  return getStaffRaw_().map(s => ({
    id: s.id, name: s.name, leaveFormName: s.leaveFormName, active: s.active, jobGroup: s.jobGroup,
    personnelType: s.personnelType, gender: s.gender, hasPin: !!s.personalPin,
    position: s.position, phone: s.phone, address: s.address,
    signatureFileId: s.signatureFileId,
    signatureImage: signatureImageDataUrl_(s.signatureFileId || ''),
    fiscalYearBE: currentFiscalYearBE,
    vacationCarryDays: getVacationCarryDaysFor_(s, currentFiscalYearBE)
  }));
}

function getMyStaffProfile(staffName, pin) {
  if (!checkStaffPin_(staffName, pin)) return { success: false, message: 'PIN ไม่ถูกต้อง หรือไม่พบเจ้าหน้าที่' };
  const staff = getStaffRaw_().find(s => s.name === staffName && s.active);
  if (!staff) return { success: false, message: 'ไม่พบข้อมูลเจ้าหน้าที่' };
  const currentFiscalYearBE = getCurrentFiscalYearBE_();
  const seedInfo = getVacationCarrySeedInfo_(staff.name, currentFiscalYearBE);
  const previousSeedInfo = getVacationCarrySeedInfo_(staff.name, currentFiscalYearBE - 1);
  const accumulationEligibility = getVacationAccumulationEligibility_(staff);
  return {
    success: true,
    staff: {
      name: staff.name,
      leaveFormName: staff.leaveFormName || staff.name,
      gender: staff.gender || '',
      position: staff.position || '',
      phone: staff.phone || '',
      address: staff.address || '',
      signatureFileId: staff.signatureFileId || '',
      signatureImage: signatureImageDataUrl_(staff.signatureFileId || ''),
      fiscalYearBE: currentFiscalYearBE,
      vacationCarryDays: seedInfo ? Number(seedInfo.vacationCarryDays) || 0 : '',
      vacationCarryIsSet: !!seedInfo,
      vacationAccumulationEligible: accumulationEligibility.eligible,
      vacationAccumulationConfigured: accumulationEligibility.configured,
      vacationAccumMax: accumulationEligibility.accumMax,
      vacationAccumulationReason: accumulationEligibility.reason,
      previousFiscalYearBE: currentFiscalYearBE - 1,
      previousVacationCarryDays: previousSeedInfo ? Number(previousSeedInfo.vacationCarryDays) || 0 : '',
      previousVacationCarryAvailable: !!previousSeedInfo,
      vacationCarryUpdatedAt: seedInfo ? seedInfo.updatedAt : '',
      vacationCarryUpdatedBy: seedInfo ? seedInfo.updatedBy : '',
      lastLeaveSeed: staff.lastLeaveSeed || {}
    }
  };
}

function clampNumber_(value, min, max, fallback) {
  const n = Number(value);
  return isNaN(n) ? fallback : Math.min(max, Math.max(min, n));
}

function getSystemCenterSettings(pin) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  const s = getSettings_();
  return { success: true, settings: {
    organizationName: s.OrganizationName || '', organizationAddress: s.OrganizationAddress || '',
    departmentName: s.DepartmentName || '', divisionName: s.DivisionName || '',
    documentLocation: s.DocumentLocation || '', documentRecipient: s.DocumentRecipient || '',
    directorFooter: (!s.DirectorFooter || s.DirectorFooter === 'ผู้อำนวยการ') ? 'ผู้อำนวยการโรงพยาบาลขุนตาล' : s.DirectorFooter,
    defaultFiscalYearBE: s.DefaultFiscalYearBE || getCurrentFiscalYearBE_(), reminderHour: Number(s.ReminderHour) || 7,
    reminderDaysBefore: Number(s.ReminderDaysBefore) || 0, fontFamily: s.LeaveFormFontFamily || 'Sarabun',
    fontSizePt: Number(s.LeaveFormFontSizePt) || 12.5, marginTopMm: Number(s.LeaveFormMarginTopMm) || 20,
    marginRightMm: Number(s.LeaveFormMarginRightMm) || 20, marginBottomMm: Number(s.LeaveFormMarginBottomMm) || 15,
    marginLeftMm: Number(s.LeaveFormMarginLeftMm) || 25, signatureHeightMm: Number(s.LeaveFormSignatureHeightMm) || 18,
    respectText: s.LeaveFormRespectText || 'ขอแสดงความนับถือ', orderText: s.LeaveFormOrderText || 'คำสั่ง'
  }, leaveTypes: getLeaveTypeConfigs_(), quotaTable: getLeaveQuotaTable_(), audit: getRecentSettingsAudit_(25) };
}

function updateSystemCenterSettings(payload, pin) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  payload = payload || {};
  const leaveRightsDocumentUrl = String(payload.leaveRightsDocumentUrl || '').trim();
  if (leaveRightsDocumentUrl && !/^https?:\/\/\S+$/i.test(leaveRightsDocumentUrl)) {
    return { success: false, message: 'ลิงก์เอกสารสิทธิการลาต้องขึ้นต้นด้วย http:// หรือ https://' };
  }
  upsertSettings_({
    OrganizationName: String(payload.organizationName || '').trim(), OrganizationAddress: String(payload.organizationAddress || '').trim(),
    DepartmentName: String(payload.departmentName || '').trim(), DivisionName: String(payload.divisionName || '').trim(),
    DocumentLocation: String(payload.documentLocation || '').trim(), DocumentRecipient: String(payload.documentRecipient || '').trim(),
    DirectorFooter: String(payload.directorFooter || '').trim(),
    LeaveRightsNote: String(payload.leaveRightsNote || '').trim(),
    LeaveRightsDocumentUrl: leaveRightsDocumentUrl
  });
  auditSettingChange_('ข้อมูลหน่วยงาน', 'แก้ไข', 'ข้อมูลที่ใช้แสดงในใบลา', 'แอดมิน');
  return { success: true, message: 'บันทึกข้อมูลหน่วยงานเรียบร้อยแล้ว' };
}

function saveLeaveTypeConfig(config, pin) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  config = config || {};
  const name = String(config.leaveType || '').trim(), actor = String(config.updatedBy || '').trim();
  if (!name || !actor) return { success: false, message: 'กรุณากรอกชื่อประเภทลาและชื่อผู้แก้ไข' };
  const mode = ['working_exclude_holidays', 'weekdays_include_holidays', 'calendar'].indexOf(config.countingMode) !== -1 ? config.countingMode : 'working_exclude_holidays';
  const sheet = getSheetByName('LeaveTypeConfig'), data = sheet.getDataRange().getValues();
  const row = [name, !!config.active, !!config.allowHalfDay, mode, String(config.paidDefault || ''), config.formKind === 'vacation' ? 'vacation' : 'general', Utilities.formatDate(new Date(), SYSTEM_TIMEZONE, 'dd/MM/yyyy HH:mm'), actor];
  let found = false;
  for (let i = 1; i < data.length; i++) if (String(data[i][0]).trim() === name) { sheet.getRange(i + 1, 1, 1, row.length).setValues([row]); found = true; break; }
  if (!found) sheet.appendRow(row);
  clearLeaveTypeConfigCache_();
  auditSettingChange_('ประเภทลา', found ? 'แก้ไข' : 'เพิ่ม', name + (config.active ? ' (เปิดใช้)' : ' (ปิดใช้)'), actor);
  return { success: true, message: (found ? 'แก้ไข' : 'เพิ่ม') + 'ประเภทลาเรียบร้อยแล้ว' };
}

function saveLeaveQuotaConfig(config, pin) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  config = config || {};
  const personnelType = String(config.personnelType || '').trim(), leaveType = String(config.leaveType || '').trim(), actor = String(config.updatedBy || '').trim();
  if (!personnelType || !leaveType || !actor) return { success: false, message: 'กรุณากรอกข้อมูลสิทธิ์และชื่อผู้แก้ไขให้ครบ' };
  const row = [personnelType, leaveType, Math.max(0, Number(config.annualQuotaDays) || 0), String(config.paidStatus || ''), config.accumMax === '' ? '' : Math.max(0, Number(config.accumMax) || 0), String(config.note || '')];
  const sheet = getSheetByName('LeaveQuota'), data = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) if (String(data[i][0]) === personnelType && String(data[i][1]) === leaveType) { sheet.getRange(i + 1, 1, 1, 6).setValues([row]); found = true; break; }
  if (!found) sheet.appendRow(row);
  clearQuotaCache_();
  auditSettingChange_('สิทธิ์การลา', found ? 'แก้ไข' : 'เพิ่ม', personnelType + ' / ' + leaveType + ' = ' + row[2] + ' วัน', actor);
  return { success: true, message: 'บันทึกสิทธิ์การลาเรียบร้อยแล้ว' };
}

function updateMyStaffProfile(staffName, position, phone, address, signatureFileId, vacationCarryDays, lastLeaveSeed, fiscalYearBE, leaveFormName, pin) {
  if (pin === undefined && leaveFormName !== undefined) {
    pin = leaveFormName;
    leaveFormName = '';
  }
  if (pin === undefined) {
    pin = fiscalYearBE !== undefined ? fiscalYearBE : lastLeaveSeed;
    fiscalYearBE = getCurrentFiscalYearBE_();
    lastLeaveSeed = {};
  }
  if (pin === undefined && fiscalYearBE !== undefined) {
    pin = fiscalYearBE;
    fiscalYearBE = getCurrentFiscalYearBE_();
  }
  if (!checkStaffPin_(staffName, pin)) return { success: false, message: 'PIN ไม่ถูกต้อง หรือไม่พบเจ้าหน้าที่' };
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  const sheet = getSheetByName('Staff');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === staffName && (data[i][2] === true || data[i][2] === 'TRUE')) {
      const profileRow = [[
        (position || '').trim(),
        normalizeTextIdentifier_(phone),
        (address || '').trim(),
        (signatureFileId || '').trim(),
        Number(vacationCarryDays) || 0,
        JSON.stringify(normalizeLastLeaveSeed_(lastLeaveSeed || {})),
        (leaveFormName || data[i][12] || data[i][1] || '').trim()
      ]];
      sheet.getRange(i + 1, 8).setNumberFormat('@');
      sheet.getRange(i + 1, 7, 1, 7).setValues(profileRow);
      upsertVacationCarrySeed_(staffName, Number(fiscalYearBE) || getCurrentFiscalYearBE_(), vacationCarryDays, staffName + ' (แก้เอง)', 'ตั้งต้นปีงบจากหน้าแก้ไขข้อมูลของฉัน');
      clearStaffCache_();
      return { success: true, message: 'อัปเดตข้อมูลของฉันสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูลเจ้าหน้าที่' };
}

// ====== เพิ่มเจ้าหน้าที่ (ต้องตั้ง PIN ส่วนตัวให้ด้วย เพื่อให้เจ้าตัวบันทึกลาเองได้) ======
function addStaff(name, jobGroup, personnelType, personalPin, position, phone, address, signatureFileId, vacationCarryDays, gender, pin) {
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  if (pin === undefined) { pin = gender; gender = ''; }
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  if (!name || name.trim() === '') return { success: false, message: 'กรุณากรอกชื่อ' };
  if (!personalPin || String(personalPin).trim() === '') {
    return { success: false, message: 'กรุณาตั้ง PIN ส่วนตัวให้เจ้าหน้าที่ท่านนี้ด้วย' };
  }
  gender = normalizeGender_(gender) || inferGenderFromName_(name);
  if (!gender) return { success: false, message: 'กรุณาเลือกเพศ เพื่อให้ระบบกำหนดสิทธิ์ลาคลอดหรือช่วยภริยาได้ถูกต้อง' };

  const sheet = getSheetByName('Staff');
  const newId = 'S' + new Date().getTime();
  sheet.appendRow([newId, name.trim(), true, jobGroup || '', personnelType || '', String(personalPin).trim(),
    (position || '').trim(), (phone || '').trim(), (address || '').trim(), (signatureFileId || '').trim(),
    Number(vacationCarryDays) || 0, '', name.trim(), gender]);
  const addedRow = sheet.getLastRow();
  writePlainText_(sheet.getRange(addedRow, 6), personalPin);
  writePlainText_(sheet.getRange(addedRow, 8), phone);
  upsertVacationCarrySeed_(name.trim(), getCurrentFiscalYearBE_(), vacationCarryDays, 'แอดมิน', 'ตั้งต้นตอนเพิ่มเจ้าหน้าที่');
  clearStaffCache_();
  return { success: true, message: 'เพิ่ม ' + name + ' สำเร็จ' };
}

// ====== แก้ไขข้อมูลเจ้าหน้าที่ที่มีอยู่ (PIN ใหม่ใส่เฉพาะตอนต้องการเปลี่ยน เว้นว่างไว้ = ไม่เปลี่ยน) ======
function updateStaffInfo(staffId, jobGroup, personnelType, newPersonalPin, position, phone, address, signatureFileId, vacationCarryDays, leaveFormName, gender, pin) {
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังแก้ไขข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  if (pin === undefined) { pin = gender; gender = ''; }
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  const sheet = getSheetByName('Staff');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === staffId) {
      sheet.getRange(i + 1, 4, 1, 2).setValues([[
        jobGroup || '', personnelType || ''
      ]]);
      if (newPersonalPin && String(newPersonalPin).trim() !== '') {
        writePlainText_(sheet.getRange(i + 1, 6), newPersonalPin);
      }
      sheet.getRange(i + 1, 8).setNumberFormat('@');
      const normalizedGender = normalizeGender_(gender) || inferGenderFromName_(leaveFormName || data[i][12] || data[i][1]);
      if (!normalizedGender) return { success: false, message: 'กรุณาเลือกเพศ เพื่อให้ระบบกำหนดสิทธิ์ลาได้ถูกต้อง' };
      sheet.getRange(i + 1, 7, 1, 8).setValues([[
        (position || '').trim(),
        normalizeTextIdentifier_(phone),
        (address || '').trim(),
        (signatureFileId || '').trim(),
        Number(vacationCarryDays) || 0,
        data[i][11] || '',
        (leaveFormName || data[i][12] || data[i][1] || '').trim(),
        normalizedGender
      ]]);
      upsertVacationCarrySeed_(data[i][1], getCurrentFiscalYearBE_(), vacationCarryDays, 'แอดมิน', 'แก้ไขจากรายชื่อเจ้าหน้าที่ทั้งหมด');
      clearStaffCache_();
      return { success: true, message: 'อัปเดตข้อมูลสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบเจ้าหน้าที่' };
}

function toggleStaffActive(staffId, pin) {
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังแก้ไขข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  const sheet = getSheetByName('Staff');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === staffId) {
      const current = (data[i][2] === true || data[i][2] === 'TRUE');
      sheet.getRange(i + 1, 3).setValue(!current);
      clearStaffCache_();
      return { success: true, message: 'อัปเดตสถานะสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบเจ้าหน้าที่' };
}

// ====== ลบเจ้าหน้าที่ถาวร (ใช้กับคนที่ลาออกแล้ว) ======
// หมายเหตุ: ลบแค่แถวใน Staff เท่านั้น ประวัติการลาเก่าใน LeaveRecords จะยังอยู่ครบเพื่อการตรวจสอบย้อนหลัง
// (ถ้าอยากดูประวัติของคนที่ลบไปแล้ว ต้องเปิดดูตรงในชีต LeaveRecords โดยตรง เพราะหน้าเว็บจะแสดงเฉพาะเจ้าหน้าที่ที่ยังมีอยู่ในระบบ)
function deleteStaff(staffId, pin) {
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังแก้ไขข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  const sheet = getSheetByName('Staff');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === staffId) {
      const name = data[i][1];
      sheet.deleteRow(i + 1);
      clearStaffCache_();
      return { success: true, message: 'ลบ ' + name + ' ออกจากระบบแล้ว (ประวัติการลาเก่ายังอยู่ในชีต LeaveRecords)' };
    }
  }
  return { success: false, message: 'ไม่พบเจ้าหน้าที่' };
}

// ====== นับจำนวนวันทำการ (จ-ศ และไม่ใช่วันหยุดนักขัตฤกษ์) ระหว่างวันที่ (รวมวันเริ่มและวันสิ้นสุด) ======
function countWorkingDays_(startDate, endDate) {
  const holidays = getHolidaySet_();
  const tz = SYSTEM_TIMEZONE;
  let count = 0;
  const cur = new Date(startDate + 'T00:00:00');
  const end = new Date(endDate + 'T00:00:00');
  while (cur <= end) {
    const day = cur.getDay(); // 0=อาทิตย์, 6=เสาร์
    const dStr = Utilities.formatDate(cur, tz, 'yyyy-MM-dd');
    if (day !== 0 && day !== 6 && !holidays.has(dStr)) count++;
    cur.setDate(cur.getDate() + 1);
  }
  return count;
}

function countLeaveDays_(startDate, endDate, leaveType) {
  const config = getLeaveTypeConfig_(leaveType);
  const mode = config ? config.countingMode : 'working_exclude_holidays';
  if (mode === 'working_exclude_holidays') return countWorkingDays_(startDate, endDate);
  let count = 0;
  const cur = new Date(startDate + 'T00:00:00');
  const end = new Date(endDate + 'T00:00:00');
  while (cur <= end) {
    if (mode === 'calendar' || (cur.getDay() !== 0 && cur.getDay() !== 6)) count++;
    cur.setDate(cur.getDate() + 1);
  }
  return count;
}

function leaveConflictResult_(row, startDate, endDate, conflictType) {
  const existingStart = normalizeDateString_(row[3], SYSTEM_TIMEZONE);
  const existingEnd = normalizeDateString_(row[4], SYSTEM_TIMEZONE);
  const exact = conflictType === 'duplicate';
  return {
    success: false,
    code: exact ? 'DUPLICATE_LEAVE' : 'OVERLAPPING_LEAVE',
    message: exact
      ? 'พบรายการลาเดิมที่ตรงกับข้อมูลนี้ ระบบจึงไม่บันทึกซ้ำ'
      : 'พบรายการลาเดิมที่มีช่วงวันที่ทับซ้อนกัน ระบบจึงยังไม่บันทึกรายการใหม่',
    conflict: {
      conflictType: conflictType,
      recordId: String(row[0] || ''),
      staffName: String(row[1] || ''),
      leaveType: String(row[2] || ''),
      startDate: existingStart,
      endDate: existingEnd,
      totalDays: Number(row[5]) || 0,
      reason: String(row[6] || ''),
      halfDayPeriod: String(row[9] || ''),
      requestedStartDate: startDate,
      requestedEndDate: endDate
    }
  };
}

// ====== บันทึกการลา ======
// ใช้ได้ 3 แบบ: (1) เจ้าตัวใส่ PIN ส่วนตัว (2) หัวหน้างานบันทึกแทน (3) แอดมินบันทึกแทน
// halfDayPeriod: 'เช้า' หรือ 'บ่าย' — มีความหมายเฉพาะกรณี isHalfDay=true เท่านั้น
function recordLeave(staffName, leaveType, startDate, endDate, isHalfDay, halfDayPeriod, reason, pin) {
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  const tz = SYSTEM_TIMEZONE;
  startDate = normalizeDateString_(startDate, tz);
  endDate = normalizeDateString_(endDate, tz);

  const isAdmin = checkPin(pin, 'admin');
  const isSup = checkPin(pin, 'supervisor');
  const isSelf = checkStaffPin_(staffName, pin);
  if (!isAdmin && !isSup && !isSelf) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  if (!staffName || !leaveType || !startDate || !endDate) {
    return { success: false, message: 'กรอกข้อมูลไม่ครบ' };
  }
  const leaveConfig = getLeaveTypeConfig_(leaveType);
  if (!leaveConfig || !leaveConfig.active) return { success: false, message: 'ประเภทลานี้ถูกปิดใช้งานแล้ว' };
  const selectedStaff = getStaffRaw_().find(staff => staff.active && staff.name === staffName);
  if (!selectedStaff) return { success: false, message: 'ไม่พบข้อมูลเจ้าหน้าที่ที่เลือก' };
  const familyLeaveType = familyLeaveTypeForGender_(selectedStaff.gender);
  if ((leaveType === 'ลาคลอดบุตร' || leaveType === 'ลาไปช่วยเหลือภริยาที่คลอดบุตร') && !familyLeaveType) {
    return { success: false, message: 'ยังไม่ได้กำหนดเพศของเจ้าหน้าที่ กรุณาให้แอดมินแก้ไขข้อมูลก่อนบันทึกประเภทลานี้' };
  }
  if (leaveType === 'ลาคลอดบุตร' && familyLeaveType !== leaveType) {
    return { success: false, message: 'เจ้าหน้าที่ชายต้องใช้ประเภท “ลาไปช่วยเหลือภริยาที่คลอดบุตร”' };
  }
  if (leaveType === 'ลาไปช่วยเหลือภริยาที่คลอดบุตร' && familyLeaveType !== leaveType) {
    return { success: false, message: 'ประเภทนี้ใช้สำหรับเจ้าหน้าที่ชาย กรุณาเลือก “ลาคลอดบุตร”' };
  }
  if (leaveType === 'ลาพักผ่อน') {
    const accumulationEligibility = getVacationAccumulationEligibility_(selectedStaff);
    if (!accumulationEligibility.configured) {
      return { success: false, message: 'ยังไม่ได้กำหนดประเภทบุคลากร จึงตรวจสอบสิทธิ์ลาพักผ่อนไม่ได้' };
    }
    if (accumulationEligibility.eligible && !getVacationCarrySeedInfo_(staffName, fiscalYearBEFromDate_(startDate))) {
      return { success: false, message: 'ยังไม่ได้ยืนยันยอดวันลาพักผ่อนสะสมของปีงประมาณนี้ กรุณาตั้งค่าก่อนบันทึกลา' };
    }
  }
  if (isHalfDay && !leaveConfig.allowHalfDay) return { success: false, message: 'ประเภทลานี้ไม่อนุญาตให้ลาครึ่งวัน' };
  if (endDate < startDate) return { success: false, message: 'วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่ม' };

  const sheet = getSheetByName('LeaveRecords');

  // กันบันทึกซ้ำ: ชื่อ+ประเภท+ช่วงวันที่เดียวกันเป๊ะ
  const existing = sheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    const exStart = normalizeDateString_(existing[i][3], tz);
    const exEnd = normalizeDateString_(existing[i][4], tz);
    if (existing[i][1] === staffName && existing[i][2] === leaveType &&
        exStart === startDate && exEnd === endDate) {
      return leaveConflictResult_(existing[i], startDate, endDate, 'duplicate');
    }
  }

  // กันวันที่ลาซ้อนกัน (คนเดียวกัน ประเภทเดียวกัน ช่วงวันที่คาบเกี่ยวกัน) เพื่อไม่ให้นับโควต้าซ้ำ
  for (let i = 1; i < existing.length; i++) {
    if (existing[i][1] === staffName && existing[i][2] === leaveType) {
      const exStart = normalizeDateString_(existing[i][3], tz);
      const exEnd = normalizeDateString_(existing[i][4], tz);
      const overlaps = (startDate <= exEnd) && (endDate >= exStart);
      if (overlaps) {
        return leaveConflictResult_(existing[i], startDate, endDate, 'overlap');
      }
    }
  }

  let totalDays;
  const finalHalfDayPeriod = (startDate === endDate && isHalfDay) ? (halfDayPeriod || 'เช้า') : '';
  if (startDate === endDate && isHalfDay) {
    totalDays = 0.5;
  } else {
    totalDays = countLeaveDays_(startDate, endDate, leaveType);
  }

  const newId = 'L' + new Date().getTime();
  let recordedBy;
  if (isAdmin) recordedBy = 'แอดมิน';
  else if (isSup) recordedBy = 'หัวหน้างาน';
  else recordedBy = staffName + ' (บันทึกเอง)';

  sheet.appendRow([newId, staffName, leaveType, startDate, endDate, totalDays, (reason || '').trim(), recordedBy, new Date(), finalHalfDayPeriod]);

  const periodLabel = finalHalfDayPeriod ? ' (ครึ่ง' + finalHalfDayPeriod + ')' : '';

  return { success: true, message: 'บันทึก' + leaveType + 'ให้ ' + staffName + ' สำเร็จ (' + totalDays + ' วัน' + periodLabel + ')', recordId: newId };
}

function deleteLeaveRecord(recordId, pin) {
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังลบหรือแก้ไขข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  const sheet = getSheetByName('LeaveRecords');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === recordId) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบรายการสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบรายการ' };
}

// ====== แก้ไขรายการลาที่บันทึกผิด (ไม่ต้องลบแล้วบันทึกใหม่) — แอดมินเท่านั้น ======
function updateLeaveRecord(recordId, leaveType, startDate, endDate, isHalfDay, halfDayPeriod, reason, pin) {
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังแก้ไขข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  const tz = SYSTEM_TIMEZONE;
  startDate = normalizeDateString_(startDate, tz);
  endDate = normalizeDateString_(endDate, tz);

  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  if (!leaveType || !startDate || !endDate) return { success: false, message: 'กรอกข้อมูลไม่ครบ' };
  const leaveConfig = getLeaveTypeConfig_(leaveType);
  if (isHalfDay && leaveConfig && !leaveConfig.allowHalfDay) return { success: false, message: 'ประเภทลานี้ไม่อนุญาตให้ลาครึ่งวัน' };
  if (endDate < startDate) return { success: false, message: 'วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่ม' };

  const sheet = getSheetByName('LeaveRecords');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === recordId) {
      const staffName = data[i][1];
      for (let j = 1; j < data.length; j++) {
        if (j === i) continue;
        if (data[j][1] !== staffName || data[j][2] !== leaveType) continue;
        const exStart = normalizeDateString_(data[j][3], tz);
        const exEnd = normalizeDateString_(data[j][4], tz);
        const overlaps = (startDate <= exEnd) && (endDate >= exStart);
        if (overlaps) {
          return { success: false, message: 'ช่วงวันที่นี้ทับกับรายการ ' + leaveType + ' ที่บันทึกไว้แล้ว (' + exStart + (exEnd !== exStart ? ' ถึง ' + exEnd : '') + ') กรุณาตรวจสอบ' };
        }
      }

      let totalDays;
      const finalHalfDayPeriod = (startDate === endDate && isHalfDay) ? (halfDayPeriod || 'เช้า') : '';
      if (startDate === endDate && isHalfDay) totalDays = 0.5;
      else totalDays = countLeaveDays_(startDate, endDate, leaveType);

      sheet.getRange(i + 1, 3, 1, 9).setValues([[
        leaveType,
        startDate,
        endDate,
        totalDays,
        (reason || '').trim(),
        data[i][7],
        data[i][8],
        finalHalfDayPeriod,
        ''
      ]]);
      return { success: true, message: 'แก้ไขรายการสำเร็จ (' + totalDays + ' วัน)' };
    }
  }
  return { success: false, message: 'ไม่พบรายการ' };
}

// ====== สรุปการลารายปี แยกตามประเภทการลา (หน้าจอหลัก) ======
// yearStr = ปีงบประมาณแบบ พ.ศ. (เช่น "2569" หมายถึง 1 ต.ค. 2568 - 30 ก.ย. 2569 ตามระเบียบราชการ)
function getAnnualSummary(yearStr, prefetched) {
  prefetched = prefetched || {};
  const staffList = prefetched.staffList || getActiveStaffList();
  const leaveTypeConfigs = prefetched.leaveTypeConfigs || getLeaveTypeConfigs_();
  const quotaTable = prefetched.quotaTable || getLeaveQuotaTable_();
  const configuredLeaveTypes = leaveTypeConfigs.map(config => config.leaveType);
  const recordsSheet = getSheetByName('LeaveRecords');
  const allRecords = recordsSheet.getDataRange().getValues();

  // แปลงปีงบประมาณ (พ.ศ.) เป็นช่วงวันที่จริงแบบ ค.ศ.: 1 ต.ค. (ปีก่อนหน้า) ถึง 30 ก.ย. (ปีนี้)
  const fiscalYearBE = Number(yearStr);
  const gregorianEndYear = fiscalYearBE - 543;
  const gregorianStartYear = gregorianEndYear - 1;
  const rangeStart = gregorianStartYear + '-10-01';
  const rangeEnd = gregorianEndYear + '-09-30';

  // เตรียมรายการดิบของแต่ละคนที่อยู่ในปีงบประมาณนี้
  const recordsByStaff = {};
  staffList.forEach(s => { recordsByStaff[s.name] = []; });

  for (let i = 1; i < allRecords.length; i++) {
    const staffName = allRecords[i][1];
    if (recordsByStaff[staffName] === undefined) continue;
    const tz = SYSTEM_TIMEZONE;
    const startDate = normalizeDateString_(allRecords[i][3], tz);
    if (startDate < rangeStart || startDate > rangeEnd) continue;

    const endDate = normalizeDateString_(allRecords[i][4], tz);

    recordsByStaff[staffName].push({
      id: allRecords[i][0],
      leaveType: allRecords[i][2],
      startDate: startDate,
      endDate: endDate,
      totalDays: Number(allRecords[i][5]) || 0,
      reason: allRecords[i][6] || '',
      recordedBy: allRecords[i][7],
      halfDayPeriod: allRecords[i][9] || ''
    });
  }

  const result = staffList.map(s => {
    const records = recordsByStaff[s.name].sort((a, b) => a.startDate < b.startDate ? -1 : (a.startDate > b.startDate ? 1 : 0));
    const needsSetup = !s.personnelType; // ยังไม่ได้ตั้งประเภทบุคลากรให้คนนี้ (ต่างจาก "ไม่มีสิทธิ์" ของ TOR/รายวันที่ตั้งค่าแล้ว)

    const summaryTypes = Array.from(new Set(configuredLeaveTypes.concat(records.map(r => r.leaveType))))
      .filter(lt => lt !== 'ลาไปช่วยเหลือภริยาที่คลอดบุตร');
    const byType = summaryTypes.map(lt => {
      const effectiveType = lt === 'ลาคลอดบุตร' ? (familyLeaveTypeForGender_(s.gender) || lt) : lt;
      const quotaInfo = quotaTable.find(row => row.personnelType === s.personnelType && row.leaveType === effectiveType) || null;
      const usedDays = records.filter(r => r.leaveType === effectiveType).reduce((sum, r) => sum + r.totalDays, 0);
      const quotaDays = quotaInfo ? quotaInfo.annualQuotaDays : 0;
      const paidStatus = quotaInfo ? quotaInfo.paidStatus : '';

      let status = 'no_quota';
      if (quotaDays > 0) {
        if (usedDays > quotaDays) status = 'over';
        else if (usedDays >= quotaDays * 0.8) status = 'near';
        else status = 'ok';
      }

      return {
        leaveType: lt,
        displayLeaveType: effectiveType,
        usedDays: usedDays,
        quotaDays: quotaDays,
        remainingDays: quotaDays - usedDays,
        paidStatus: paidStatus,
        status: status,
        note: quotaInfo ? quotaInfo.note : ''
      };
    });

    const hasOver = byType.some(t => t.status === 'over');
    const hasNear = byType.some(t => t.status === 'near');

    return {
      name: s.name,
      jobGroup: s.jobGroup,
      personnelType: s.personnelType,
      needsSetup: needsSetup,
      byType: byType,
      records: records,
      overallStatus: needsSetup ? 'setup' : (hasOver ? 'over' : (hasNear ? 'near' : 'ok'))
    };
  });

  // เรียงให้คนที่ยังไม่ตั้งค่าขึ้นก่อนสุด (ต้องรีบแก้) ตามด้วยคนที่เกินโควต้า
  const priority = { setup: -1, over: 0, near: 1, ok: 2 };
  result.sort((a, b) => priority[a.overallStatus] - priority[b.overallStatus]);

  return { year: yearStr, data: result };
}

function getInitialData(yearStr) {
  const settings = getSettings_();
  const staffRaw = getStaffRaw_();
  const staff = staffRaw.filter(row => row.active).map(row => ({
    id: row.id, name: row.name, leaveFormName: row.leaveFormName, jobGroup: row.jobGroup,
    personnelType: row.personnelType, position: row.position, gender: row.gender
  }));
  const leaveTypeConfigs = getLeaveTypeConfigs_();
  const quotaTable = getLeaveQuotaTable_();
  const fiscalYearBE = getCurrentFiscalYearBE_();
  const prefetched = {
    staffList: staff,
    leaveTypeConfigs: leaveTypeConfigs,
    quotaTable: quotaTable
  };
  return {
    summary: getAnnualSummary(yearStr, prefetched),
    staff: staff,
    reference: getReferenceData(leaveTypeConfigs, quotaTable),
    defaultFiscalYearBE: fiscalYearBE,
    fiscalYearSetup: getFiscalYearSetupStatus(fiscalYearBE, staffRaw, quotaTable),
    announcement: {
      title: String(settings.AnnouncementTitle || DEFAULT_ANNOUNCEMENT_TITLE),
      content: String(settings.AnnouncementContent || DEFAULT_ANNOUNCEMENT_CONTENT),
      updatedAt: String(settings.AnnouncementUpdatedAt || '')
    },
    leaveRights: {
      note: String(settings.LeaveRightsNote || 'การลาเกินสิทธิ์ จะถูกหักค่าจ้างตามจำนวนวันที่ลา ทั้งนี้ ให้เป็นไปตามระเบียบที่เกี่ยวข้อง'),
      documentUrl: String(settings.LeaveRightsDocumentUrl || '')
    }
  };
}

// ====== ข้อมูลปฏิทินรายเดือน (แสดงภาพรวมทั้งเดือน) ======
// yearStr: 'YYYY', monthStr: '1'-'12'
function getCalendarData(yearStr, monthStr) {
  const year = Number(yearStr);
  const month = Number(monthStr); // 1-12
  const daysInMonth = new Date(year, month, 0).getDate();
  const firstDayWeekday = new Date(year, month - 1, 1).getDay(); // 0=อาทิตย์

  const monthPrefix = yearStr + '-' + String(month).padStart(2, '0');
  const firstOfMonth = monthPrefix + '-01';
  const lastOfMonth = monthPrefix + '-' + String(daysInMonth).padStart(2, '0');

  const sheet = getSheetByName('LeaveRecords');
  const data = sheet.getDataRange().getValues();
  const tz = SYSTEM_TIMEZONE;

  // entries[day] = [{staffName, leaveType, isHalfDay, recordId}, ...]
  const entries = {};
  const seenCalendarEntries = {};
  for (let d = 1; d <= daysInMonth; d++) entries[d] = [];

  for (let i = 1; i < data.length; i++) {
    let start = data[i][3], end = data[i][4];
    if (Object.prototype.toString.call(start) === '[object Date]') start = Utilities.formatDate(start, tz, 'yyyy-MM-dd');
    if (Object.prototype.toString.call(end) === '[object Date]') end = Utilities.formatDate(end, tz, 'yyyy-MM-dd');
    start = normalizeDateString_(start, tz);
    end = normalizeDateString_(end, tz);
    if (!start || !end) continue;

    // ข้ามถ้าช่วงลาไม่ทับกับเดือนนี้เลย
    if (end < firstOfMonth || start > lastOfMonth) continue;

    const totalDays = Number(data[i][5]) || 0;
    const isHalfDay = (start === end && totalDays === 0.5);
    const halfDayPeriod = data[i][9] || '';

    // ไล่ทีละวันในช่วงที่ทับกับเดือนนี้
    const rangeStart = start > firstOfMonth ? start : firstOfMonth;
    const rangeEnd = end < lastOfMonth ? end : lastOfMonth;
    const cur = new Date(rangeStart + 'T00:00:00');
    const endD = new Date(rangeEnd + 'T00:00:00');
    while (cur <= endD) {
      const dayOfWeek = cur.getDay(); // 0=อาทิตย์, 6=เสาร์
      if (dayOfWeek !== 0 && dayOfWeek !== 6) {
        const dayNum = cur.getDate();
        // One person and one leave type must appear only once per calendar day.
        // Trim/collapse hidden whitespace because legacy Sheet rows may look identical on screen
        // while containing different spaces or half-day metadata.
        const calendarStaffKey = String(data[i][1] || '').trim().replace(/\s+/g, ' ').toLowerCase();
        const calendarTypeKey = String(data[i][2] || '').trim().replace(/\s+/g, ' ').toLowerCase();
        const dedupeKey = dayNum + '|' + calendarStaffKey + '|' + calendarTypeKey;
        if (!seenCalendarEntries[dedupeKey]) {
          seenCalendarEntries[dedupeKey] = true;
          entries[dayNum].push({
            staffName: data[i][1],
            leaveType: data[i][2],
            isHalfDay: isHalfDay,
            halfDayPeriod: halfDayPeriod,
            recordId: data[i][0]
          });
        }
      }
      cur.setDate(cur.getDate() + 1);
    }
  }

  return {
    year: year, month: month, daysInMonth: daysInMonth, firstDayWeekday: firstDayWeekday, entries: entries
  };
}


// ====================================================================
// LINE MESSAGING API (ทดแทน LINE Notify ที่ปิดบริการไปแล้ว 31 มี.ค. 2568)
// ====================================================================

// รับ Webhook จาก LINE — ใช้ตอนเชิญบอทเข้ากลุ่ม เพื่อดัก Group ID มาเก็บอัตโนมัติ
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const events = body.events || [];
    events.forEach(ev => {
      // ประมวลผลเฉพาะ "ตอนบอทถูกเชิญเข้ากลุ่ม" (join) เท่านั้น
      // ข้อความสนทนาปกติในกลุ่ม (type: 'message') จะถูกข้ามไปเลย ไม่ตอบกลับ ไม่ประมวลผลใดๆทั้งสิ้น
      if (ev.type !== 'join') return;
      if (ev.source && ev.source.type === 'group' && ev.source.groupId) {
        const sheet = getSheetByName('Settings');
        const data = sheet.getDataRange().getValues();
        let found = false;
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === 'LineGroupId') {
            sheet.getRange(i + 1, 2).setValue(ev.source.groupId);
            found = true;
          }
        }
        if (!found) sheet.appendRow(['LineGroupId', ev.source.groupId]);
        clearSettingsCache_();

        // ตอบกลับยืนยันแค่ครั้งเดียวตอนเชิญเข้ากลุ่มเท่านั้น (ใช้ replyToken ไม่เสียโควต้าข้อความ)
        if (ev.replyToken) {
          replyLineMessage_(ev.replyToken, '✅ เชื่อมต่อกลุ่มนี้กับระบบลาเจ้าหน้าที่เรียบร้อยแล้วค่ะ');
        }
      }
    });
  } catch (err) {
    // เงียบไว้ ไม่ให้ webhook error กระทบระบบหลัก
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}

function replyLineMessage_(replyToken, text) {
  const settings = getSettings_();
  const token = settings['LineChannelAccessToken'];
  if (!token) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true
  });
}

function pushLineMessage_(text) {
  const settings = getSettings_();
  const token = settings['LineChannelAccessToken'];
  const groupId = settings['LineGroupId'];
  if (!token || !groupId) return { success: false, message: 'ยังไม่ได้ตั้งค่า LINE Token หรือ Group ID' };

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: groupId, messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code === 200) return { success: true, message: 'ส่งข้อความสำเร็จ' };
  return { success: false, message: 'ส่งไม่สำเร็จ (HTTP ' + code + '): ' + res.getContentText() };
}

function getLineQuotaStatus(pin) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN แอดมินไม่ถูกต้อง' };
  const token = getSettings_()['LineChannelAccessToken'];
  if (!token) return { success: false, message: 'ยังไม่ได้ตั้งค่า LINE Channel Access Token' };

  const requestOptions = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  };
  try {
    const responses = UrlFetchApp.fetchAll([
      { url: 'https://api.line.me/v2/bot/message/quota', ...requestOptions },
      { url: 'https://api.line.me/v2/bot/message/quota/consumption', ...requestOptions }
    ]);
    const quotaCode = responses[0].getResponseCode();
    const usageCode = responses[1].getResponseCode();
    if (quotaCode !== 200 || usageCode !== 200) {
      return {
        success: false,
        message: 'ตรวจโควตา LINE ไม่สำเร็จ (HTTP ' + quotaCode + '/' + usageCode + ') กรุณาตรวจสอบ Token'
      };
    }

    const quota = JSON.parse(responses[0].getContentText() || '{}');
    const consumption = JSON.parse(responses[1].getContentText() || '{}');
    const used = Math.max(0, Number(consumption.totalUsage) || 0);
    const limited = quota.type === 'limited' && Number.isFinite(Number(quota.value));
    const limit = limited ? Math.max(0, Number(quota.value)) : null;
    const remaining = limited ? Math.max(0, limit - used) : null;
    const percent = limited && limit > 0 ? Math.min(100, (used / limit) * 100) : null;
    let level = 'normal';
    if (percent !== null && percent >= 95) level = 'critical';
    else if (percent !== null && percent >= 80) level = 'warning';

    return {
      success: true,
      quotaType: quota.type || 'unknown',
      used: used,
      limit: limit,
      remaining: remaining,
      percent: percent,
      level: level,
      checkedAt: Utilities.formatDate(new Date(), SYSTEM_TIMEZONE, 'dd/MM/yyyy HH:mm')
    };
  } catch (err) {
    return { success: false, message: 'ตรวจโควตา LINE ไม่สำเร็จ: ' + err.message };
  }
}

// ปุ่ม "ส่งข้อความทดสอบ" ในหน้าตั้งค่า
function sendTestLineMessage(pin) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  return pushLineMessage_('🔔 ทดสอบการเชื่อมต่อระบบลาเจ้าหน้าที่กับ LINE กลุ่มนี้ — ถ้าเห็นข้อความนี้ แสดงว่าเชื่อมต่อสำเร็จค่ะ');
}

function getLeaveFormSignerSettings(pin) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  const settings = getSettings_();
  return {
    success: true,
    supervisorName: settings['LeaveFormSupervisorName'] || '',
    supervisorPosition: settings['LeaveFormSupervisorPosition'] || '',
    supervisorSignatureFileId: settings['LeaveFormSupervisorSignatureFileId'] || '',
    supervisorSignatureImage: signatureImageDataUrl_(settings['LeaveFormSupervisorSignatureFileId'] || ''),
    approverName: settings['LeaveFormApproverName'] || '',
    approverPosition: settings['LeaveFormApproverPosition'] || '',
    approverSignatureFileId: settings['LeaveFormApproverSignatureFileId'] || '',
    approverSignatureImage: signatureImageDataUrl_(settings['LeaveFormApproverSignatureFileId'] || '')
  };
}

function updateLeaveFormSignerSettings(supervisorName, supervisorPosition, supervisorSignatureFileId, approverName, approverPosition, approverSignatureFileId, updatedBy, pin) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  updatedBy = String(updatedBy || 'แอดมิน').trim() || 'แอดมิน';
  const sheet = getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  const updates = {
    LeaveFormSupervisorName: (supervisorName || '').trim(),
    LeaveFormSupervisorPosition: (supervisorPosition || '').trim(),
    LeaveFormSupervisorSignatureFileId: (supervisorSignatureFileId || '').trim(),
    LeaveFormApproverName: (approverName || '').trim(),
    LeaveFormApproverPosition: (approverPosition || '').trim(),
    LeaveFormApproverSignatureFileId: (approverSignatureFileId || '').trim()
  };
  Object.keys(updates).forEach(key => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(updates[key]);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([key, updates[key]]);
  });
  clearSettingsCache_();
  auditSettingChange_('ผู้ลงนามใบลา', 'แก้ไข', 'ผู้บังคับบัญชาและผู้ลงนามคำสั่ง', updatedBy);
  return { success: true, message: 'บันทึกผู้ลงนามคำสั่งสำเร็จ' };
}

// บันทึกการตั้งค่า LINE (Token + จำนวนวันแจ้งเตือนล่วงหน้า)
function updateLineSettings(channelToken, reminderDaysBefore, pin) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  const sheet = getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  const reminderSetting = Number(reminderDaysBefore);
  const updates = {
    LineChannelAccessToken: channelToken,
    ReminderDaysBefore: isNaN(reminderSetting) ? 1 : Math.max(0, reminderSetting)
  };
  Object.keys(updates).forEach(key => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(updates[key]); found = true; }
    }
    if (!found) sheet.appendRow([key, updates[key]]);
  });
  clearSettingsCache_();
  return { success: true, message: 'บันทึกการตั้งค่า LINE สำเร็จ' };
}

// รันทุกวันอัตโนมัติ (ตั้งเวลาผ่าน createDailyTrigger) — เตือนล่วงหน้า + แจ้งวันที่ลาจริง
// ชื่อฟังก์ชันนี้ห้ามเปลี่ยน — ผูกไว้กับ time-driven trigger ที่ตั้งไว้ (createDailyTrigger อ้างอิงชื่อนี้)
function runDailyLineNotification() {
  const result = runDailyLineNotification_(false);
  if (result.hasContent && !result.success) {
    Logger.log('ส่งแจ้งเตือน LINE ไม่สำเร็จ: ' + result.message);
  }
}

// core logic แยกออกมา ให้ทั้ง trigger อัตโนมัติ และปุ่ม "ทดสอบส่งแจ้งเตือนตอนนี้เลย" เรียกใช้ร่วมกันได้
// คืนผลแบบละเอียด เพื่อให้ปุ่มทดสอบแยกได้ว่าไม่มีข้อมูล หรือมีข้อมูลแต่ส่ง LINE ไม่สำเร็จ
function runDailyLineNotification_(isManualTest) {
  const reminderDays = 1;

  const tz = SYSTEM_TIMEZONE;
  const today = new Date();
  const todayStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
  const reminderDate = new Date(today);
  reminderDate.setDate(reminderDate.getDate() + reminderDays);
  const reminderDateStr = Utilities.formatDate(reminderDate, tz, 'yyyy-MM-dd');

  const sheet = getSheetByName('LeaveRecords');
  const data = sheet.getDataRange().getValues();

  const todayLines = [];
  const reminderLines = [];

  for (let i = 1; i < data.length; i++) {
    let start = data[i][3], end = data[i][4];
    if (Object.prototype.toString.call(start) === '[object Date]') start = Utilities.formatDate(start, tz, 'yyyy-MM-dd');
    if (Object.prototype.toString.call(end) === '[object Date]') end = Utilities.formatDate(end, tz, 'yyyy-MM-dd');
    start = normalizeDateString_(start, tz);
    end = normalizeDateString_(end, tz);
    if (!start || !end) continue;

    const staffName = data[i][1];
    const leaveType = data[i][2];

    if (start === reminderDateStr) {
      reminderLines.push('• ' + staffName + ' — ' + leaveType + ' (' + formatSystemDateTH_(start) +
        (end !== start ? ' ถึง ' + formatSystemDateTH_(end) : '') + ')');
    }
    if (todayStr >= start && todayStr <= end) {
      todayLines.push('• ' + staffName + ' — ' + leaveType +
        (end !== start ? ' (ถึง ' + formatSystemDateTH_(end) + ')' : ''));
    }
  }

  let message = isManualTest ? '🧪 (ทดสอบด้วยตนเอง)\n' : '';
  if (reminderLines.length) {
    message += '🔔 แจ้งเตือนล่วงหน้า (อีก ' + reminderDays + ' วันจะถึงวันลา):\n' + reminderLines.join('\n') + '\n\n';
  }
  if (todayLines.length) {
    message += '📋 วันนี้มีเจ้าหน้าที่ลา:\n' + todayLines.join('\n');
  }

  const hasContent = reminderLines.length > 0 || todayLines.length > 0;
  if (!hasContent) return { hasContent: false, success: true, message: 'ไม่มีข้อมูลให้แจ้งเตือน' };

  const sendResult = pushLineMessage_(message.trim());
  return { hasContent: true, success: sendResult.success, message: sendResult.message };
}

// ====================================================================
// สร้างใบลา PDF อัตโนมัติจากแบบฟอร์ม Google Docs (แม่แบบ) — เติมข้อมูลจากรายการลาที่บันทึกไว้
// ====================================================================

const THAI_MONTHS_FULL = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];

// แปลง 'YYYY-MM-DD' เป็น {day, month, year} แบบไทย (วัน, ชื่อเดือนเต็ม, ปี พ.ศ.)
function thaiDateParts_(dateStr) {
  if (!dateStr || dateStr === '-') return { day: '-', month: '-', year: '-' };
  const parts = String(dateStr).split('-');
  const sourceYear = Number(parts[0]);
  const y = sourceYear >= 2400 ? sourceYear : sourceYear + 543;
  const m = Number(parts[1]) - 1;
  const d = Number(parts[2]);
  return { day: String(d), month: THAI_MONTHS_FULL[m] || '-', year: String(y) };
}

function normalizeDateString_(value, tz) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, tz || SYSTEM_TIMEZONE, 'yyyy-MM-dd');
  }
  const text = String(value).trim();
  const match = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (match) {
    const sourceYear = Number(match[1]);
    const year = sourceYear >= 2400 ? sourceYear - 543 : sourceYear;
    return year + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[3]).slice(-2);
  }
  return text;
}

function getLeaveFormOutputFolder_() {
  const settings = getSettings_();
  const folderId = String(settings['LeaveFormOutputFolderId'] || '').trim();
  if (folderId) return DriveApp.getFolderById(folderId);

  const folders = DriveApp.getFoldersByName('LeaveForms');
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder('LeaveForms');
}

function safeFileName_(text) {
  return String(text || '').replace(/[\\/:*?"<>|#%\{\}~&]/g, '_').replace(/\s+/g, ' ').trim();
}

function formatThaiDateText_(dateStr) {
  const p = thaiDateParts_(dateStr);
  if (p.day === '-') return '-';
  return p.day + ' ' + p.month + ' ' + p.year;
}

function signatureImageDataUrl_(fileId) {
  const id = extractDriveFileId_(fileId);
  if (!id) return '';
  try {
    const blob = DriveApp.getFileById(id).getBlob();
    const mimeType = blob.getContentType() || 'image/png';
    if (mimeType.indexOf('image/') !== 0) return '';
    return 'data:' + mimeType + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (err) {
    return '';
  }
}

function extractDriveFileId_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const match = text.match(/\/d\/([a-zA-Z0-9_-]{20,})|[?&]id=([a-zA-Z0-9_-]{20,})|^([a-zA-Z0-9_-]{20,})$/);
  return match ? (match[1] || match[2] || match[3] || '') : '';
}

function uploadSignatureBlob_(dataUrl, fileName) {
  const match = String(dataUrl || '').match(/^data:(image\/(?:png|jpeg|jpg));base64,(.+)$/);
  if (!match) throw new Error('รองรับเฉพาะรูป PNG หรือ JPG');
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 5 * 1024 * 1024) throw new Error('รูปมีขนาดเกิน 5 MB');
  const mime = match[1] === 'image/jpg' ? 'image/jpeg' : match[1];
  const safeName = String(fileName || 'signature.png').replace(/[^a-zA-Z0-9ก-๙._ -]/g, '_');
  return DriveApp.createFile(Utilities.newBlob(bytes, mime, 'ลายเซ็น_' + safeName));
}

function uploadMySignatureImage(staffName, dataUrl, fileName, pin) {
  const isAdmin = checkPin(pin, 'admin');
  if (!isAdmin && !checkStaffPin_(staffName, pin)) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกลายเซ็นจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  const file = uploadSignatureBlob_(dataUrl, fileName);
  const sheet = getSheetByName('Staff');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === staffName) {
      sheet.getRange(i + 1, 10).setValue(file.getId());
      clearStaffCache_();
      return { success: true, message: 'บันทึกรูปลายเซ็นแล้ว', fileId: file.getId(), image: signatureImageDataUrl_(file.getId()) };
    }
  }
  file.setTrashed(true);
  return { success: false, message: 'ไม่พบข้อมูลเจ้าหน้าที่' };
}

function uploadSignerSignatureImage_(settingKey, label, dataUrl, fileName, pin) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN แอดมินไม่ถูกต้อง' };
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกลายเซ็นจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  const file = uploadSignatureBlob_(dataUrl, fileName);
  const sheet = getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === settingKey) {
      sheet.getRange(i + 1, 2).setValue(file.getId());
      found = true;
      break;
    }
  }
  if (!found) sheet.appendRow([settingKey, file.getId()]);
  clearSettingsCache_();
  return { success: true, message: 'บันทึกรูปลายเซ็น' + label + 'แล้ว', fileId: file.getId(), image: signatureImageDataUrl_(file.getId()) };
}

function uploadApproverSignatureImage(dataUrl, fileName, pin) {
  return uploadSignerSignatureImage_('LeaveFormApproverSignatureFileId', 'ผู้ลงนามคำสั่ง', dataUrl, fileName, pin);
}

function uploadSupervisorSignatureImage(dataUrl, fileName, pin) {
  return uploadSignerSignatureImage_('LeaveFormSupervisorSignatureFileId', 'ผู้บังคับบัญชา', dataUrl, fileName, pin);
}

function getLeaveFormData(recordId) {
  const sheet = getSheetByName('LeaveRecords');
  const data = sheet.getDataRange().getValues();
  const tz = SYSTEM_TIMEZONE;

  let record = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === recordId) {
      const start = normalizeDateString_(data[i][3], tz);
      const end = normalizeDateString_(data[i][4], tz);
      record = {
        id: data[i][0],
        staffName: data[i][1],
        leaveType: data[i][2],
        startDate: start,
        endDate: end,
        startDateText: formatThaiDateText_(start),
        endDateText: formatThaiDateText_(end),
        totalDays: Number(data[i][5]) || 0,
        reason: data[i][6] || '',
        recordedBy: data[i][7] || '',
        halfDayPeriod: data[i][9] || '',
        pdfFileId: data[i][10] || ''
      };
      break;
    }
  }
  if (!record) return { success: false, message: 'ไม่พบรายการลานี้' };

  const staff = getStaffRaw_().find(s => s.name === record.staffName) || {};
  const settings = getSettings_();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const fiscalYearBE = fiscalYearBEFromDate_(record.startDate);
  const dayText = record.totalDays === 0.5
    ? 'ครึ่งวัน' + (record.halfDayPeriod ? ' (' + record.halfDayPeriod + ')' : '')
    : record.totalDays + ' วัน';
  const lastLeave = findLastSameTypeLeave_(record.staffName, record.leaveType, record.startDate, record.id) || lastLeaveSeedForType_(staff, record.leaveType);

  let storedPdf = null;
  if (record.pdfFileId) {
    try {
      const pdfFile = DriveApp.getFileById(String(record.pdfFileId));
      storedPdf = {
        fileId: pdfFile.getId(),
        fileName: pdfFile.getName(),
        url: pdfFile.getUrl(),
        previewUrl: 'https://drive.google.com/file/d/' + pdfFile.getId() + '/preview',
        downloadUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId()
      };
    } catch (err) {}
  }

  return {
    success: true,
    generatedDateText: formatThaiDateText_(today),
    documentSettings: getDocumentSettings_(),
    record: record,
    staff: {
      name: staff.leaveFormName || record.staffName,
      position: staff.position || '',
      jobGroup: staff.jobGroup || '',
      personnelType: staff.personnelType || '',
      phone: staff.phone || '',
      address: staff.address || '',
      signatureImage: signatureImageDataUrl_(staff.signatureFileId || ''),
      fiscalYearBE: fiscalYearBE,
      vacationCarryDays: getVacationCarryDaysFor_(staff, fiscalYearBE)
    },
    signatures: {
      supervisor: {
        name: settings['LeaveFormSupervisorName'] || '',
        position: settings['LeaveFormSupervisorPosition'] || '',
        image: signatureImageDataUrl_(settings['LeaveFormSupervisorSignatureFileId'] || '')
      },
      approver: {
        name: settings['LeaveFormApproverName'] || '',
        position: settings['LeaveFormApproverPosition'] || '',
        image: signatureImageDataUrl_(settings['LeaveFormApproverSignatureFileId'] || '')
      }
    },
    lastLeave: lastLeave ? {
      startDate: lastLeave.start,
      endDate: lastLeave.end,
      startDateText: formatThaiDateText_(lastLeave.start),
      endDateText: formatThaiDateText_(lastLeave.end),
      totalDays: lastLeave.days
    } : null,
    dayText: dayText,
    storedPdf: storedPdf
  };
}

function uploadRenderedLeavePdf(recordId, dataUrl, fileName, pin) {
  const sheet = getSheetByName('LeaveRecords');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let staffName = '';
  let startDate = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === recordId) {
      rowIndex = i + 1;
      staffName = String(data[i][1] || '');
      startDate = normalizeDateString_(data[i][3], SYSTEM_TIMEZONE);
      break;
    }
  }
  if (rowIndex < 0) return { success: false, message: 'ไม่พบรายการลาสำหรับบันทึก PDF' };
  const allowed = checkPin(pin, 'admin') || checkPin(pin, 'supervisor') || checkStaffPin_(staffName, pin);
  if (!allowed) return { success: false, message: 'PIN ไม่ถูกต้อง ไม่สามารถบันทึก PDF ได้' };
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกเอกสารจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };

  const match = String(dataUrl || '').match(/^data:application\/pdf;base64,(.+)$/);
  if (!match) return { success: false, message: 'รูปแบบไฟล์ PDF ไม่ถูกต้อง' };
  const bytes = Utilities.base64Decode(match[1]);
  if (bytes.length > 12 * 1024 * 1024) return { success: false, message: 'ไฟล์ PDF มีขนาดเกิน 12 MB' };
  if (bytes.length < 4 || bytes[0] !== 37 || bytes[1] !== 80 || bytes[2] !== 68 || bytes[3] !== 70) {
    return { success: false, message: 'ข้อมูลที่ส่งมาไม่ใช่ไฟล์ PDF' };
  }

  const outputFolder = getLeaveFormOutputFolder_();
  const safeName = safeFileName_(fileName || ('ใบลา_' + staffName + '_' + startDate + '_' + recordId + '.pdf'));
  const finalName = /\.pdf$/i.test(safeName) ? safeName : safeName + '.pdf';
  const pdfFile = outputFolder.createFile(Utilities.newBlob(bytes, 'application/pdf', finalName));
  if (sheet.getLastColumn() < 11 || String(sheet.getRange(1, 11).getValue()).trim() !== 'PdfFileId') {
    sheet.getRange(1, 11).setValue('PdfFileId');
  }
  sheet.getRange(rowIndex, 11).setValue(pdfFile.getId());

  return {
    success: true,
    fileId: pdfFile.getId(),
    pdfUrl: pdfFile.getUrl(),
    pdfPreviewUrl: 'https://drive.google.com/file/d/' + pdfFile.getId() + '/preview',
    pdfDownloadUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId(),
    message: 'สร้าง PDF A4 หนึ่งหน้าและเก็บใน Google Drive เรียบร้อยแล้ว'
  };
}

function previewLeaveFormData(staffName, leaveType, startDate, endDate, isHalfDay, halfDayPeriod, reason) {
  const tz = SYSTEM_TIMEZONE;
  startDate = normalizeDateString_(startDate, tz);
  endDate = normalizeDateString_(endDate, tz);

  if (!staffName || !leaveType || !startDate || !endDate) return { success: false, message: 'กรอกข้อมูลไม่ครบ' };
  if (endDate < startDate) return { success: false, message: 'วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่ม' };

  const finalHalfDayPeriod = (startDate === endDate && isHalfDay) ? (halfDayPeriod || 'เช้า') : '';
  const leaveConfig = getLeaveTypeConfig_(leaveType);
  if (!leaveConfig || !leaveConfig.active) return { success: false, message: 'ประเภทลานี้ถูกปิดใช้งานแล้ว' };
  if (finalHalfDayPeriod && !leaveConfig.allowHalfDay) return { success: false, message: 'ประเภทลานี้ไม่อนุญาตให้ลาครึ่งวัน' };
  const totalDays = finalHalfDayPeriod ? 0.5 : countLeaveDays_(startDate, endDate, leaveType);
  const staff = getStaffRaw_().find(s => s.name === staffName && s.active);
  if (!staff) return { success: false, message: 'ไม่พบข้อมูลเจ้าหน้าที่' };

  const warnings = [];
  const sheet = getSheetByName('LeaveRecords');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] !== staffName || data[i][2] !== leaveType) continue;
    const exStart = normalizeDateString_(data[i][3], tz);
    const exEnd = normalizeDateString_(data[i][4], tz);
    if (exStart === startDate && exEnd === endDate) {
      warnings.push('มีรายการลาแบบเดียวกันช่วงวันที่นี้อยู่แล้ว');
      break;
    }
    if ((startDate <= exEnd) && (endDate >= exStart)) {
      warnings.push('ช่วงวันที่นี้ทับกับรายการ ' + leaveType + ' ที่บันทึกไว้แล้ว (' + exStart + (exEnd !== exStart ? ' ถึง ' + exEnd : '') + ')');
      break;
    }
  }

  const settings = getSettings_();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const fiscalYearBE = fiscalYearBEFromDate_(startDate);
  const record = {
    id: 'PREVIEW',
    staffName: staffName,
    leaveType: leaveType,
    startDate: startDate,
    endDate: endDate,
    startDateText: formatThaiDateText_(startDate),
    endDateText: formatThaiDateText_(endDate),
    totalDays: totalDays,
    reason: reason || '',
    recordedBy: 'พรีวิวก่อนบันทึก',
    halfDayPeriod: finalHalfDayPeriod
  };
  const lastLeave = findLastSameTypeLeave_(staffName, leaveType, startDate, '') || lastLeaveSeedForType_(staff, leaveType);
  return {
    success: true,
    previewOnly: true,
    message: warnings.length ? warnings.join(' · ') : 'พรีวิวใบลา ยังไม่ได้บันทึกจริง',
    warnings: warnings,
    generatedDateText: formatThaiDateText_(today),
    documentSettings: getDocumentSettings_(),
    record: record,
    staff: {
      name: staff.leaveFormName || staffName,
      position: staff.position || '',
      jobGroup: staff.jobGroup || '',
      personnelType: staff.personnelType || '',
      phone: staff.phone || '',
      address: staff.address || '',
      signatureImage: signatureImageDataUrl_(staff.signatureFileId || ''),
      fiscalYearBE: fiscalYearBE,
      vacationCarryDays: getVacationCarryDaysFor_(staff, fiscalYearBE)
    },
    signatures: {
      supervisor: {
        name: settings['LeaveFormSupervisorName'] || '',
        position: settings['LeaveFormSupervisorPosition'] || '',
        image: signatureImageDataUrl_(settings['LeaveFormSupervisorSignatureFileId'] || '')
      },
      approver: {
        name: settings['LeaveFormApproverName'] || '',
        position: settings['LeaveFormApproverPosition'] || '',
        image: signatureImageDataUrl_(settings['LeaveFormApproverSignatureFileId'] || '')
      }
    },
    lastLeave: lastLeave ? {
      startDate: lastLeave.start,
      endDate: lastLeave.end,
      startDateText: formatThaiDateText_(lastLeave.start),
      endDateText: formatThaiDateText_(lastLeave.end),
      totalDays: lastLeave.days
    } : null,
    dayText: totalDays === 0.5 ? 'ครึ่งวัน' + (finalHalfDayPeriod ? ' (' + finalHalfDayPeriod + ')' : '') : totalDays + ' วัน'
  };
}

// บันทึกการตั้งค่า Template ID ของแบบฟอร์มใบลา
function updateDocTemplateSettings(templateIdSickPersonalMaternity, templateIdVacation, pin, templateIdSick, templateIdPersonal, templateIdMaternity) {
  if (!checkPin(pin, 'admin')) return { success: false, message: 'PIN ไม่ถูกต้อง' };
  if (!acquireWriteLock_()) return { success: false, message: 'ระบบกำลังบันทึกข้อมูลจากผู้ใช้อื่น กรุณารอสักครู่แล้วลองใหม่' };
  const sheet = getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  const updates = {
    DocTemplateId_SickPersonalMaternity: (templateIdSickPersonalMaternity || '').trim(),
    DocTemplateId_Sick: (templateIdSick || templateIdSickPersonalMaternity || '').trim(),
    DocTemplateId_Personal: (templateIdPersonal || templateIdSickPersonalMaternity || '').trim(),
    DocTemplateId_Maternity: (templateIdMaternity || templateIdSickPersonalMaternity || '').trim(),
    DocTemplateId_Vacation: (templateIdVacation || '').trim()
  };
  Object.keys(updates).forEach(key => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(updates[key]); found = true; }
    }
    if (!found) sheet.appendRow([key, updates[key]]);
  });
  clearSettingsCache_();
  return { success: true, message: 'บันทึกแม่แบบใบลาสำเร็จ' };
}

function getLeaveFormTemplateId_(leaveType, settings) {
  if (leaveType === 'ลาพักผ่อน') return settings['DocTemplateId_Vacation'] || '';
  if (leaveType === 'ลาป่วย') return settings['DocTemplateId_Sick'] || settings['DocTemplateId_SickPersonalMaternity'] || '';
  if (leaveType === 'ลากิจส่วนตัว') return settings['DocTemplateId_Personal'] || settings['DocTemplateId_SickPersonalMaternity'] || '';
  if (leaveType === 'ลาคลอดบุตร') return settings['DocTemplateId_Maternity'] || settings['DocTemplateId_SickPersonalMaternity'] || '';
  return settings['DocTemplateId_SickPersonalMaternity'] || '';
}

// หาสถิติการลาประเภทหนึ่งๆ ในปีงบประมาณ ก่อนรายการนี้ / ครั้งนี้ / รวม
function computeLeaveStat_(staffName, leaveType, fiscalYearStr, excludeRecordId, thisRecordDays) {
  const sheet = getSheetByName('LeaveRecords');
  const data = sheet.getDataRange().getValues();
  const fiscalYearBE = Number(fiscalYearStr);
  const gregorianEndYear = fiscalYearBE - 543;
  const rangeStart = (gregorianEndYear - 1) + '-10-01';
  const rangeEnd = gregorianEndYear + '-09-30';
  const tz = SYSTEM_TIMEZONE;

  let usedBefore = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] !== staffName || data[i][2] !== leaveType) continue;
    if (data[i][0] === excludeRecordId) continue;
    const d = normalizeDateString_(data[i][3], tz);
    if (d < rangeStart || d > rangeEnd) continue;
    usedBefore += Number(data[i][5]) || 0;
  }
  return { used: usedBefore, thisTime: thisRecordDays, total: usedBefore + thisRecordDays };
}

// หารายการลาประเภทเดียวกันครั้งล่าสุดก่อนหน้ารายการนี้ (ไม่รวมตัวเอง)
function findLastSameTypeLeave_(staffName, leaveType, beforeStartDate, excludeRecordId) {
  const sheet = getSheetByName('LeaveRecords');
  const data = sheet.getDataRange().getValues();
  const tz = SYSTEM_TIMEZONE;
  beforeStartDate = normalizeDateString_(beforeStartDate, tz);
  if (!staffName || !leaveType || !beforeStartDate) return null;
  let best = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] !== staffName || data[i][2] !== leaveType) continue;
    if (data[i][0] === excludeRecordId) continue;
    const s = normalizeDateString_(data[i][3], tz);
    const en = normalizeDateString_(data[i][4], tz);
    if (!s || !en || en < s) continue;
    // วันสิ้นสุดต้องอยู่ก่อนวันเริ่มของใบปัจจุบันจริง จึงถือเป็น "ลาครั้งสุดท้าย"
    if (en >= beforeStartDate) continue;
    if (!best || en > best.end || (en === best.end && s > best.start)) {
      best = { start: s, end: en, days: Number(data[i][5]) || 0 };
    }
  }
  return best;
}

// ====== ฟังก์ชันหลัก: สร้างใบลา PDF + Word แล้วบันทึกไว้ใน Drive พร้อมส่งลิงก์กลับให้หน้าเว็บ ======
function generateLeaveFormPdf(recordId) {
  const settings = getSettings_();
  const sheet = getSheetByName('LeaveRecords');
  const data = sheet.getDataRange().getValues();
  const tz = SYSTEM_TIMEZONE;

  let record = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === recordId) {
      let s = data[i][3], en = data[i][4];
      s = normalizeDateString_(s, tz);
      en = normalizeDateString_(en, tz);
      record = {
        id: data[i][0], staffName: data[i][1], leaveType: data[i][2],
        startDate: String(s), endDate: String(en), totalDays: Number(data[i][5]) || 0,
        reason: data[i][6] || ''
      };
      break;
    }
  }
  if (!record) return { success: false, message: 'ไม่พบรายการลานี้' };

  const staff = getStaffRaw_().find(s => s.name === record.staffName);
  if (!staff) return { success: false, message: 'ไม่พบข้อมูลเจ้าหน้าที่' };

  const isVacation = (record.leaveType === 'ลาพักผ่อน');
  const templateId = getLeaveFormTemplateId_(record.leaveType, settings);
  if (!templateId) {
    return { success: false, message: 'ยังไม่ได้ตั้งค่าแม่แบบใบลา (' + (isVacation ? 'ลาพักผ่อน' : 'ลาป่วย/กิจ/คลอด') +
      ') กรุณาไปตั้งค่าในหน้า "จัดการระบบ" ก่อน' };
  }

  // คำนวณปีงบประมาณจากวันที่เริ่มลา
  const startGYear = Number(record.startDate.substring(0, 4));
  const startGMonth = Number(record.startDate.substring(5, 7));
  const fiscalYearBE = (startGMonth >= 10 ? startGYear + 1 : startGYear) + 543;

  let copyFile, docBody, pdfBlob, docxBlob;
  try {
    const templateFile = DriveApp.getFileById(templateId);
    const outputFolder = getLeaveFormOutputFolder_();
    const leaveFormName = staff.leaveFormName || record.staffName;
    const baseName = safeFileName_('ใบลา_' + leaveFormName + '_' + record.startDate + '_' + record.id);
    copyFile = templateFile.makeCopy(baseName + '_GoogleDoc', outputFolder);
    const doc = DocumentApp.openById(copyFile.getId());
    docBody = doc.getBody();

    const fileDate = thaiDateParts_(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'));
    const startParts = thaiDateParts_(record.startDate);
    const endParts = thaiDateParts_(record.endDate);

    const replace = (marker, value) => docBody.replaceText(marker, String(value != null ? value : '-'));

    replace('<<FILE_DAY>>', fileDate.day);
    replace('<<FILE_MONTH>>', fileDate.month);
    replace('<<FILE_YEAR>>', fileDate.year);
    replace('<<STAFF_NAME>>', leaveFormName);
    replace('<<POSITION>>', staff.position || '-');
    replace('<<START_DAY>>', startParts.day);
    replace('<<START_MONTH>>', startParts.month);
    replace('<<START_YEAR>>', startParts.year);
    replace('<<END_DAY>>', endParts.day);
    replace('<<END_MONTH>>', endParts.month);
    replace('<<END_YEAR>>', endParts.year);
    replace('<<TOTAL_DAYS>>', record.totalDays);
    replace('<<PHONE>>', staff.phone || '-');
    replace('<<ADDRESS>>', staff.address || '-');

    if (isVacation) {
      const quotaInfo = getQuotaFor_(staff.personnelType, 'ลาพักผ่อน');
      const entitled = quotaInfo ? quotaInfo.annualQuotaDays : 0;
      const accum = getVacationCarryDaysFor_(staff, fiscalYearBE);
      replace('<<ACCUM_DAYS>>', accum);
      replace('<<ENTITLED_DAYS>>', entitled);
      replace('<<TOTAL_ENTITLED>>', accum + entitled);
      // ตารางสถิติการลา (ลามาแล้ว/ลาครั้งนี้/รวมเป็น) เว้นว่างไว้ให้ HR ตรวจสอบและกรอกเอง
      replace('<<STAT_VAC_USED>>', '');
      replace('<<STAT_VAC_THIS>>', '');
      replace('<<STAT_VAC_TOTAL>>', '');
    } else {
      replace('<<SUBJECT>>', 'ขอ' + record.leaveType);
      replace('<<SUBJECT_SHORT>>', record.leaveType.replace('ลา', ''));
      replace('<<REASON>>', record.reason || '-');
      replace('<<CHECK_SICK>>', record.leaveType === 'ลาป่วย' ? '☒' : '☐');
      replace('<<CHECK_PERSONAL>>', record.leaveType === 'ลากิจส่วนตัว' ? '☒' : '☐');
      replace('<<CHECK_MATERNITY>>', record.leaveType === 'ลาคลอดบุตร' ? '☒' : '☐');

      const last = findLastSameTypeLeave_(record.staffName, record.leaveType, record.startDate, record.id);
      if (last) {
        const lastStart = thaiDateParts_(last.start), lastEnd = thaiDateParts_(last.end);
        replace('<<LAST_START_DAY>>', lastStart.day); replace('<<LAST_START_MONTH>>', lastStart.month); replace('<<LAST_START_YEAR>>', lastStart.year);
        replace('<<LAST_END_DAY>>', lastEnd.day); replace('<<LAST_END_MONTH>>', lastEnd.month); replace('<<LAST_END_YEAR>>', lastEnd.year);
        replace('<<LAST_TOTAL_DAYS>>', last.days);
      } else {
        ['<<LAST_START_DAY>>','<<LAST_START_MONTH>>','<<LAST_START_YEAR>>','<<LAST_END_DAY>>','<<LAST_END_MONTH>>','<<LAST_END_YEAR>>','<<LAST_TOTAL_DAYS>>']
          .forEach(m => replace(m, '-'));
      }

      // ตารางสถิติการลา เว้นว่างไว้ให้ HR ตรวจสอบและกรอกเอง (ไม่ดึงจากระบบอัตโนมัติตามที่แจ้ง)
      ['SICK', 'PERSONAL', 'MATERNITY'].forEach(key => {
        replace('<<STAT_' + key + '_USED>>', '');
        replace('<<STAT_' + key + '_THIS>>', '');
        replace('<<STAT_' + key + '_TOTAL>>', '');
      });
    }

    doc.saveAndClose();

    const renderedFile = DriveApp.getFileById(copyFile.getId());
    pdfBlob = renderedFile.getAs('application/pdf').setName(baseName + '.pdf');
    docxBlob = renderedFile.getAs('application/vnd.openxmlformats-officedocument.wordprocessingml.document').setName(baseName + '.docx');
    const pdfFile = outputFolder.createFile(pdfBlob);
    const docxFile = outputFolder.createFile(docxBlob);

    return {
      success: true,
      pdfUrl: pdfFile.getUrl(),
      wordUrl: docxFile.getUrl(),
      pdfDownloadUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId(),
      wordDownloadUrl: 'https://drive.google.com/uc?export=download&id=' + docxFile.getId(),
      googleDocUrl: renderedFile.getUrl(),
      message: 'สร้างใบลาเรียบร้อยแล้ว กดดาวน์โหลดไฟล์ลงเครื่องได้เลย'
    };
  } catch (err) {
    return { success: false, message: 'สร้างใบลาไม่สำเร็จ: ' + err.message };
  } finally {
    // เก็บ Google Docs สำเนาไว้ในโฟลเดอร์เดียวกัน เผื่อต้องแก้ไขก่อนส่ง HR
  }
}
