/**
 * Google Apps Script — เก็บวันที่ว่างลงตาราง "ชื่อ x วันที่" ใน Google Sheet
 *
 * รูปแบบตาราง (ตามชีตต้นแบบ):
 *   แถว 1  : หัวข้อ / วันที่ว่าง
 *   แถว 2  : คำอธิบายสี / เดือน <ชื่อเดือน> <ปี พ.ศ.>
 *   แถว 3  : เลขวันที่ 1..28/29/30/31 (เริ่มคอลัมน์ B)
 *   แถว 4  : "ชื่อ" (คอลัมน์ A) + ตัวย่อวันในสัปดาห์
 *   แถว 5+ : รายชื่อในคอลัมน์ A
 *            ✓ = ว่าง (เขียว)   ✗ = ไม่ว่าง (แดง)   ช่องว่าง = ยังไม่กรอก
 *   ท้ายตาราง: สรุปจำนวนคนที่ว่าง / ไม่ว่าง ของแต่ละวัน
 *
 * หลังวางโค้ดใหม่: เลือกฟังก์ชัน setupSheets แล้วกด ▶ Run หนึ่งครั้ง เพื่อจัดรูปแบบทุกแท็บ
 */

const START = { y: 2026, m: 12 }; // ธันวาคม 2569
const END = { y: 2027, m: 5 };    // พฤษภาคม 2570
const MARK = '✓';      // ว่าง
const NO_MARK = '✗';   // ไม่ว่าง
const DAY_ROW = 3, DOW_ROW = 4, FIRST_NAME_ROW = 5, FIRST_DAY_COL = 2;
const TH_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const TH_DOW = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];

// สี
const C = {
  title: '#0f766e', titleText: '#ffffff',
  month: '#ccfbf1', monthText: '#134e4a',
  head: '#f1f5f9', headText: '#0f172a', dowText: '#64748b',
  weekendHead: '#fde68a', weekendCell: '#fffbeb',
  rowA: '#ffffff', rowB: '#f8fafc',
  border: '#cbd5e1',
  okBg: '#b7e1cd', okText: '#0d652d',
  noBg: '#f4c7c3', noText: '#a50e0e',
  sumOk: '#ecfdf5', sumNo: '#fef2f2',
};

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

/** สร้างแท็บเดือนที่ยังไม่มี + จัดรูปแบบทุกแท็บ (รันซ้ำได้ ข้อมูลไม่หาย) */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const months = monthList_();
  const tabNames = months.map(function (mo) { return mo.tab; });
  const tpl = findTemplate_(ss);
  months.forEach(function (mo) {
    let sh = ss.getSheetByName(mo.tab);
    if (!sh) sh = buildMonth_(ss, tpl, mo);
    styleMonth_(sh, mo);
  });
  // ซ่อนชีตต้นแบบ (ยังอยู่ ไม่ได้ลบ — คลิกขวาที่แถบชีตด้านล่าง > แสดง เพื่อเปิดกลับ)
  if (tabNames.indexOf(tpl.getName()) < 0 && !tpl.isSheetHidden()) {
    ss.setActiveSheet(ss.getSheetByName(months[0].tab));
    tpl.hideSheet();
  }
}

/** ก๊อปชีตต้นแบบ แล้วเปลี่ยนหัวตารางเป็นของเดือนนั้น */
function buildMonth_(ss, tpl, mo) {
  const sh = tpl.copyTo(ss).setName(mo.tab);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(ss.getNumSheets());

  const lastCol = FIRST_DAY_COL + mo.days - 1;
  if (sh.getMaxColumns() < lastCol) sh.insertColumnsAfter(sh.getMaxColumns(), lastCol - sh.getMaxColumns());

  const width = sh.getMaxColumns() - FIRST_DAY_COL + 1;
  const n = Math.max(1, nameCount_(sh));
  sh.getRange(FIRST_NAME_ROW, FIRST_DAY_COL, n, width).clearContent();
  sh.getRange(1, FIRST_DAY_COL, DOW_ROW, width).breakApart().clearContent();

  const nums = [], dows = [];
  for (let d = 1; d <= mo.days; d++) {
    nums.push(d);
    dows.push(TH_DOW[new Date(mo.y, mo.m - 1, d).getDay()]);
  }
  sh.getRange(1, FIRST_DAY_COL).setValue('วันที่ว่าง');
  sh.getRange(2, FIRST_DAY_COL).setValue('เดือน ' + TH_MONTHS[mo.m - 1] + ' ' + (mo.y + 543));
  sh.getRange(DAY_ROW, FIRST_DAY_COL, 1, mo.days).setValues([nums]);
  sh.getRange(DOW_ROW, FIRST_DAY_COL, 1, mo.days).setValues([dows]);
  return sh;
}

/** จัดรูปแบบแท็บเดือนให้อ่านง่าย (ไม่แตะข้อมูล ✓ / ✗) */
function styleMonth_(sh, mo) {
  const n = nameCount_(sh);
  const lastCol = FIRST_DAY_COL + mo.days - 1;
  const lastNameRow = FIRST_NAME_ROW + Math.max(n, 1) - 1;
  const sumRow = lastNameRow + 2;

  // ตัดคอลัมน์ส่วนเกินทางขวาทิ้ง (ว่างอยู่แล้ว) และเพิ่มแถวให้พอสำหรับสรุป
  if (sh.getMaxColumns() > lastCol) sh.deleteColumns(lastCol + 1, sh.getMaxColumns() - lastCol);
  if (sh.getMaxRows() < sumRow + 1) sh.insertRowsAfter(sh.getMaxRows(), sumRow + 1 - sh.getMaxRows());

  const table = sh.getRange(1, 1, sumRow + 1, lastCol);
  table.breakApart();
  table.clearFormat();
  sh.getRange(lastNameRow + 1, 1, 1, lastCol).clearContent(); // แถวเว้นว่างก่อนสรุป

  // วันเสาร์-อาทิตย์
  const weekend = [];
  for (let d = 1; d <= mo.days; d++) {
    const w = new Date(mo.y, mo.m - 1, d).getDay();
    weekend.push(w === 0 || w === 6);
  }

  // ขนาด / ฟอนต์ทั้งตาราง
  sh.setHiddenGridlines(true);
  sh.setTabColor(C.title);
  sh.setColumnWidth(1, 280);
  sh.setColumnWidths(FIRST_DAY_COL, mo.days, 34);
  table.setFontSize(10).setVerticalAlignment('middle').setFontColor(C.headText);

  // แถว 1: หัวข้อ
  sh.getRange('A1').setValue('🧳 นัดวันไปเที่ยว');
  sh.getRange(1, FIRST_DAY_COL, 1, mo.days).merge();
  sh.getRange(1, 1, 1, lastCol).setBackground(C.title).setFontColor(C.titleText)
    .setFontWeight('bold').setFontSize(13);
  sh.getRange(1, FIRST_DAY_COL).setHorizontalAlignment('center');
  sh.setRowHeight(1, 34);

  // แถว 2: คำอธิบายสี + ชื่อเดือน
  sh.getRange('A2').setValue('🟩 ว่าง    🟥 ไม่ว่าง    ⬜ ยังไม่กรอก');
  sh.getRange(2, FIRST_DAY_COL, 1, mo.days).merge();
  sh.getRange(2, 1, 1, lastCol).setBackground(C.month).setFontColor(C.monthText).setFontWeight('bold');
  sh.getRange(2, FIRST_DAY_COL).setHorizontalAlignment('center').setFontSize(12);
  sh.setRowHeight(2, 28);

  // แถว 3-4: เลขวันที่ + วันในสัปดาห์
  sh.getRange(DAY_ROW, 1).setValue('วันที่');
  sh.getRange(DAY_ROW, 1, 2, lastCol).setBackground(C.head);
  sh.getRange(DAY_ROW, 1, 1, lastCol).setFontWeight('bold');
  sh.getRange(DAY_ROW, FIRST_DAY_COL, 2, mo.days).setHorizontalAlignment('center');
  sh.getRange(DOW_ROW, FIRST_DAY_COL, 1, mo.days).setFontSize(8).setFontColor(C.dowText);
  sh.getRange(DOW_ROW, 1).setFontWeight('bold');
  sh.setRowHeights(DAY_ROW, 2, 22);
  weekend.forEach(function (isWe, i) {
    if (isWe) sh.getRange(DAY_ROW, FIRST_DAY_COL + i, 2, 1).setBackground(C.weekendHead);
  });

  // แถวรายชื่อ: สลับสีแถว + ไฮไลต์เสาร์-อาทิตย์
  if (n > 0) {
    const bg = [];
    for (let r = 0; r < n; r++) {
      const base = r % 2 ? C.rowB : C.rowA;
      const row = [base];
      weekend.forEach(function (isWe) { row.push(isWe ? C.weekendCell : base); });
      bg.push(row);
    }
    sh.getRange(FIRST_NAME_ROW, 1, n, lastCol).setBackgrounds(bg);
    sh.getRange(FIRST_NAME_ROW, FIRST_DAY_COL, n, mo.days)
      .setHorizontalAlignment('center').setFontWeight('bold');
    sh.setRowHeights(FIRST_NAME_ROW, n, 24);
  }

  // เส้นตาราง
  sh.getRange(DAY_ROW, 1, lastNameRow - DAY_ROW + 1, lastCol)
    .setBorder(true, true, true, true, true, true, C.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(1, 1, lastNameRow, lastCol)
    .setBorder(true, true, true, true, null, null, C.title, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // สรุปท้ายตาราง: กี่คนว่าง / ไม่ว่าง ในแต่ละวัน
  const grid = 'R' + FIRST_NAME_ROW + 'C:R' + lastNameRow + 'C';
  const ok = [], no = [];
  for (let d = 0; d < mo.days; d++) {
    ok.push('=COUNTIF(' + grid + ',"' + MARK + '")');
    no.push('=COUNTIF(' + grid + ',"' + NO_MARK + '")');
  }
  sh.getRange(sumRow, 1, 2, 1).setValues([['✅ ว่าง (คน)'], ['❌ ไม่ว่าง (คน)']]);
  sh.getRange(sumRow, FIRST_DAY_COL, 2, mo.days).setFormulasR1C1([ok, no]);
  const sum = sh.getRange(sumRow, 1, 2, lastCol);
  sum.setFontWeight('bold')
    .setBorder(true, true, true, true, true, true, C.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(sumRow, 1, 1, lastCol).setBackground(C.sumOk).setFontColor(C.okText);
  sh.getRange(sumRow + 1, 1, 1, lastCol).setBackground(C.sumNo).setFontColor(C.noText);
  sh.getRange(sumRow, FIRST_DAY_COL, 2, mo.days).setHorizontalAlignment('center');
  sh.setRowHeights(sumRow, 2, 24);

  // สีอัตโนมัติ: ✓ เขียว, ✗ แดง, แถวสรุป "ว่าง" ยิ่งเยอะยิ่งเขียวเข้ม
  const cells = sh.getRange(FIRST_NAME_ROW, FIRST_DAY_COL, Math.max(n, 1), mo.days);
  const okSum = sh.getRange(sumRow, FIRST_DAY_COL, 1, mo.days);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(MARK)
      .setBackground(C.okBg).setFontColor(C.okText).setRanges([cells]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(NO_MARK)
      .setBackground(C.noBg).setFontColor(C.noText).setRanges([cells]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMaxpoint('#34a853').setRanges([okSum]).build(),
  ]);

  sh.setFrozenRows(DOW_ROW);
  sh.setFrozenColumns(1);
}

/** จำนวนรายชื่อ (นับจากแถว 5 ลงไปจนเจอช่องว่าง) */
function nameCount_(sh) {
  const max = sh.getLastRow() - FIRST_NAME_ROW + 1;
  if (max <= 0) return 0;
  const v = sh.getRange(FIRST_NAME_ROW, 1, max, 1).getValues();
  let n = 0;
  while (n < v.length && String(v[n][0]).trim() !== '') n++;
  return n;
}

function names_(sh) {
  const n = nameCount_(sh);
  if (!n) return [];
  return sh.getRange(FIRST_NAME_ROW, 1, n, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); });
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
    lock.waitLock(30000);
    try { setupSheets(); } finally { lock.releaseLock(); }
  }
  return { ss: ss, months: months };
}

/** GET → รายชื่อ + วันที่ว่าง (✓) ของทุกคนที่กรอกแล้ว */
function doGet() {
  try {
    const ctx = ensureSetup_();
    const names = names_(ctx.ss.getSheetByName(ctx.months[0].tab));
    const byName = {}, filled = {};
    names.forEach(function (n) { byName[n] = []; });

    ctx.months.forEach(function (mo) {
      const sh = ctx.ss.getSheetByName(mo.tab);
      const rows = nameCount_(sh);
      if (!rows) return;
      const vals = sh.getRange(FIRST_NAME_ROW, 1, rows, 1 + mo.days).getValues();
      vals.forEach(function (r) {
        const name = String(r[0]).trim();
        if (!byName[name]) return;
        for (let d = 1; d <= mo.days; d++) {
          const v = String(r[d]).trim();
          if (v) filled[name] = true;
          if (v === MARK) byName[name].push(iso_(mo, d));
        }
      });
    });

    return json_({
      ok: true,
      range: { start: iso_(ctx.months[0], 1), end: iso_(ctx.months[ctx.months.length - 1], ctx.months[ctx.months.length - 1].days) },
      names: names,
      responses: names.filter(function (n) { return filled[n]; })
        .map(function (n) { return { name: n, dates: byName[n] }; })
    });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/** POST {name, dates:[yyyy-mm-dd]} → วันที่เลือก = ✓ (เขียว), วันอื่นทั้งหมด = ✗ (แดง) */
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
          row.push(on ? MARK : NO_MARK);
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
