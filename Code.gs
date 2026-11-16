/**
 * Google Apps Script สำหรับเก็บวันที่ว่างลง Google Sheet
 * วิธีติดตั้ง: ดู README.md
 *
 * ชีต "Responses" จะมีคอลัมน์: Timestamp | Name | Date | Note
 * (1 แถว = 1 คน x 1 วันที่ว่าง  → เอาไปทำ Pivot table / นับง่าย)
 */

const SHEET_NAME = 'Responses';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Timestamp', 'Name', 'Date', 'Note']);
    sh.setFrozenRows(1);
    sh.getRange('A1:D1').setFontWeight('bold');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// กันสูตรแปลกๆ ถูกฉีดเข้าชีต (เช่นชื่อที่ขึ้นต้นด้วย = + - @)
function safeText_(s, max) {
  s = String(s || '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function toIso_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v).trim();
}

/** อ่านข้อมูลทั้งหมด → [{name, dates, note}] */
function readAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const values = getSheet_().getDataRange().getValues().slice(1);
  const byName = {};
  values.forEach(function (row) {
    const name = String(row[1]).replace(/^'/, '').trim();
    if (!name) return;
    const key = name.toLowerCase();
    if (!byName[key]) byName[key] = { name: name, dates: [], note: String(row[3] || '') };
    byName[key].dates.push(toIso_(row[2], tz));
  });
  return Object.keys(byName).map(function (k) { return byName[k]; });
}

/** GET → ส่งข้อมูลทั้งหมดกลับไปให้หน้าเว็บสรุป */
function doGet() {
  return json_({ ok: true, responses: readAll_() });
}

/** POST → บันทึก/แก้ไขวันที่ว่างของ 1 คน (ชื่อเดิม = เขียนทับของเก่า) */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const name = safeText_(data.name, 50);
    const note = safeText_(data.note, 200);
    const dates = (data.dates || [])
      .filter(function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(d); })
      .slice(0, 366);
    if (!name || !dates.length) return json_({ ok: false, error: 'ต้องมีชื่อและวันที่อย่างน้อย 1 วัน' });

    const sh = getSheet_();
    const key = name.replace(/^'/, '').toLowerCase();

    // ลบแถวเก่าของคนนี้ (ถ้าเคยส่งแล้ว) แล้วเขียนใหม่
    const all = sh.getDataRange().getValues();
    const keep = all.slice(1).filter(function (r) {
      return String(r[1]).replace(/^'/, '').trim().toLowerCase() !== key;
    });
    const now = new Date();
    const fresh = dates.map(function (d) { return [now, name, d, note]; });
    const rows = keep.concat(fresh);

    if (all.length > 1) sh.getRange(2, 1, all.length - 1, 4).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);

    return json_({ ok: true, saved: dates.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
