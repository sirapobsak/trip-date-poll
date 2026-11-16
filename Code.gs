/**
 * Google Apps Script — เก็บวันที่ว่างลงตาราง "ชื่อ x วันที่" ใน Google Sheet
 *
 * รูปแบบตาราง (ตามชีตต้นแบบ):
 *   แถว 1  : วันที่ว่าง
 *   แถว 2  : เดือน <ชื่อเดือน> <ปี พ.ศ.>
 *   แถว 3  : เลขวันที่ 1..28/29/30/31 (เริ่มคอลัมน์ B)
 *   แถว 4  : "ชื่อ" (คอลัมน์ A) + ตัวย่อวันในสัปดาห์
 *   แถว 5+ : รายชื่อในคอลัมน์ A, ช่องวันที่ที่ว่าง = ✓
 *
 * สคริปต์จะสร้างแท็บ 1 แท็บต่อเดือน ตั้งแต่ START ถึง END โดยก๊อปรูปแบบ+รายชื่อจากชีตต้นแบบ
 * (ชีตต้นแบบ = ชีตแรกที่ช่อง A4 เขียนว่า "ชื่อ")
 */

const START = { y: 2026, m: 12 }; // ธันวาคม 2569
const END = { y: 2027, m: 5 };    // พฤษภาคม 2570
const MARK = '✓';
const DAY_ROW = 3, DOW_ROW = 4, FIRST_NAME_ROW = 5, FIRST_DAY_COL = 2;
const TH_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const TH_DOW = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];

function monthList_() {
  const out = [];
  let y = START.y, m = START.m;
  while (y < END.y || (y === END.y && m <= END.m)) {
    out.push({ y: y, m: m, days: new Date(y, m, 0).getDate(), tab: TH_MONTHS[m - 1] + ' ' + (y + 543) });
    m++; if (m > 12) { m = 1; y++; }
  }
  return out;
}

function findTemplate_(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getRange('A4').getValue()).trim() === 'ชื่อ') return sheets[i];
  }
  throw new Error('ไม่เจอชีตต้นแบบ (ช่อง A4 ต้องเขียนว่า "ชื่อ")');
}

/** รันครั้งเดียวจากหน้า Apps Script ก็ได้ (หรือจะปล่อยให้เว็บเรียกสร้างเองตอนใช้งานครั้งแรกก็ได้) */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tpl = findTemplate_(ss);
  monthList_().forEach(function (mo) {
    if (ss.getSheetByName(mo.tab)) return; // มีแล้ว ไม่ทับข้อมูลเดิม
    buildMonth_(ss, tpl, mo);
  });
}

function buildMonth_(ss, tpl, mo) {
  const sh = tpl.copyTo(ss).setName(mo.tab);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(ss.getNumSheets());

  const lastCol = FIRST_DAY_COL + mo.days - 1;
  const lastRow = Math.max(sh.getLastRow(), FIRST_NAME_ROW);
  const tplLastDayCol = Math.max(FIRST_DAY_COL, lastFilledCol_(tpl, DAY_ROW));

  if (sh.getMaxColumns() < lastCol) sh.insertColumnsAfter(sh.getMaxColumns(), lastCol - sh.getMaxColumns());

  // ถ้าเดือนนี้ยาวกว่าต้นแบบ → ก๊อปรูปแบบคอลัมน์วันสุดท้ายของต้นแบบไปใช้กับคอลัมน์ที่เพิ่ม
  if (lastCol > tplLastDayCol) {
    const src = sh.getRange(1, tplLastDayCol, lastRow, 1);
    src.copyFormatToRange(sh, tplLastDayCol + 1, lastCol, 1, lastRow);
    sh.setColumnWidths(tplLastDayCol + 1, lastCol - tplLastDayCol, sh.getColumnWidth(tplLastDayCol));
  }

  // ล้างข้อมูลเก่า (ช่องวันที่ทั้งหมด) และหัวตารางส่วนเกิน
  const width = sh.getMaxColumns() - FIRST_DAY_COL + 1;
  sh.getRange(FIRST_NAME_ROW, FIRST_DAY_COL, Math.max(1, lastRow - FIRST_NAME_ROW + 1), width).clearContent();
  sh.getRange(1, FIRST_DAY_COL, DOW_ROW, width).breakApart().clearContent();
  if (sh.getMaxColumns() > lastCol) {
    sh.getRange(1, lastCol + 1, sh.getMaxRows(), sh.getMaxColumns() - lastCol).clearFormat();
  }

  // หัวตาราง
  sh.getRange(1, FIRST_DAY_COL, 1, mo.days).merge().setValue('วันที่ว่าง');
  sh.getRange(2, FIRST_DAY_COL, 1, mo.days).merge().setValue('เดือน ' + TH_MONTHS[mo.m - 1] + ' ' + (mo.y + 543));
  const nums = [], dows = [];
  for (let d = 1; d <= mo.days; d++) {
    nums.push(d);
    dows.push(TH_DOW[new Date(mo.y, mo.m - 1, d).getDay()]);
  }
  sh.getRange(DAY_ROW, FIRST_DAY_COL, 1, mo.days).setValues([nums]);
  sh.getRange(DOW_ROW, FIRST_DAY_COL, 1, mo.days).setValues([dows])
    .setFontSize(8).setHorizontalAlignment('center');

  // ช่องที่มี ✓ → สีเขียว
  const grid = sh.getRange(FIRST_NAME_ROW, FIRST_DAY_COL, lastRow - FIRST_NAME_ROW + 1, mo.days);
  grid.setHorizontalAlignment('center');
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(MARK).setBackground('#b7e1cd').setFontColor('#0d652d')
    .setRanges([grid]).build();
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([rule]));
  sh.setFrozenRows(DOW_ROW);
  sh.setFrozenColumns(1);
}

function lastFilledCol_(sh, row) {
  const vals = sh.getRange(row, 1, 1, sh.getMaxColumns()).getValues()[0];
  for (let c = vals.length; c >= 1; c--) if (String(vals[c - 1]).trim() !== '') return c;
  return 1;
}

function names_(sh) {
  const n = Math.max(0, sh.getLastRow() - FIRST_NAME_ROW + 1);
  if (!n) return [];
  return sh.getRange(FIRST_NAME_ROW, 1, n, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (s) { return s; });
}

function pad_(n) { return (n < 10 ? '0' : '') + n; }
function iso_(mo, d) { return mo.y + '-' + pad_(mo.m) + '-' + pad_(d); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureSetup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const months = monthList_();
  if (months.some(function (mo) { return !ss.getSheetByName(mo.tab); })) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { setupSheets(); } finally { lock.releaseLock(); }
  }
  return { ss: ss, months: months };
}

/** GET → รายชื่อ + วันที่ว่างของทุกคน */
function doGet() {
  try {
    const ctx = ensureSetup_();
    const names = names_(ctx.ss.getSheetByName(ctx.months[0].tab));
    const byName = {};
    names.forEach(function (n) { byName[n] = []; });

    ctx.months.forEach(function (mo) {
      const sh = ctx.ss.getSheetByName(mo.tab);
      const rows = names_(sh).length;
      if (!rows) return;
      const vals = sh.getRange(FIRST_NAME_ROW, 1, rows, 1 + mo.days).getValues();
      vals.forEach(function (r) {
        const name = String(r[0]).trim();
        if (!byName[name]) return;
        for (let d = 1; d <= mo.days; d++) {
          if (String(r[d]).trim() !== '') byName[name].push(iso_(mo, d));
        }
      });
    });

    return json_({
      ok: true,
      range: { start: iso_(ctx.months[0], 1), end: iso_(ctx.months[ctx.months.length - 1], ctx.months[ctx.months.length - 1].days) },
      names: names,
      responses: names.filter(function (n) { return byName[n].length; })
        .map(function (n) { return { name: n, dates: byName[n] }; })
    });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/** POST {name, dates:[yyyy-mm-dd]} → เขียน ✓ ลงแถวของคนนั้น (ทับของเดิมทั้งหมด) */
function doPost(e) {
  try {
    const ctx = ensureSetup_();
    const data = JSON.parse(e.postData.contents);
    const name = String(data.name || '').trim();
    const wanted = {};
    (data.dates || []).forEach(function (d) { if (/^\d{4}-\d{2}-\d{2}$/.test(d)) wanted[d] = true; });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      let found = false, saved = 0;
      ctx.months.forEach(function (mo) {
        const sh = ctx.ss.getSheetByName(mo.tab);
        const idx = names_(sh).indexOf(name);
        if (idx < 0) return;
        found = true;
        const row = [];
        for (let d = 1; d <= mo.days; d++) {
          const on = !!wanted[iso_(mo, d)];
          if (on) saved++;
          row.push(on ? MARK : '');
        }
        sh.getRange(FIRST_NAME_ROW + idx, FIRST_DAY_COL, 1, mo.days).setValues([row]);
      });
      if (!found) return json_({ ok: false, error: 'ไม่พบชื่อนี้ในชีต' });
      return json_({ ok: true, saved: saved });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}
