/* =========================================================================
   Code.gs — ตัวกลาง (API) เชื่อมเว็บแอป "ป้ายระบุและควบคุมการใช้จัมโบ้"
   เข้ากับ Google Sheet นี้

   วิธีติดตั้ง:
   1) เปิด Google Sheet ที่จะใช้เก็บข้อมูล → เมนู ส่วนขยาย (Extensions) → Apps Script
   2) ลบโค้ดเดิมทั้งหมดในไฟล์ Code.gs แล้ววางไฟล์นี้แทน
   3) แก้ค่า SECRET_TOKEN ด้านล่างเป็นรหัสลับที่ตั้งเอง (ตัวอักษร/ตัวเลขผสมกันยาวๆ)
   4) เลือกฟังก์ชัน "setupSheets" จากช่องเลือกฟังก์ชันด้านบน แล้วกด Run (▶) หนึ่งครั้ง
      เพื่อสร้างชีต Jumbos / Usage และข้อมูลตัวอย่าง
   5) กด Deploy → New deployment → เลือกประเภท "Web app"
        - Execute as: Me
        - Who has access: Anyone
      แล้วกด Deploy คัดลอก "Web app URL" ที่ได้
   6) นำ URL และ SECRET_TOKEN ไปใส่ในไฟล์ index.html ของเว็บแอป
   ========================================================================= */

const SECRET_TOKEN = 'CHANGE_THIS_TOKEN_1234';  // <-- แก้เป็นรหัสลับของคุณเอง แล้วใส่ค่าเดียวกันใน index.html
const SHEET_JUMBOS = 'Jumbos';
const SHEET_USAGE = 'Usage';

// จำกัดจำนวนครั้งการใช้งานแยกตามชนิดข้าว (ต้องตรงกับ MAX_USES_BY_TYPE ใน index.html)
const MAX_USES_BY_TYPE = { 'ข้าวไทย': 20, 'ข้าวญี่ปุ่น': 5 };
const MAX_USES_DEFAULT = 20;
function getMaxUses_(type){ return MAX_USES_BY_TYPE[type] || MAX_USES_DEFAULT; }

function doGet(e) {
  try {
    if ((e.parameter.token || '') !== SECRET_TOKEN) return jsonOut({ error: 'unauthorized' });
    const action = e.parameter.action;
    if (action === 'list') return jsonOut(getAllData());
    return jsonOut({ error: 'unknown action' });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if ((body.token || '') !== SECRET_TOKEN) return jsonOut({ error: 'unauthorized' });
    const action = body.action;
    if (action === 'addJumbo') return jsonOut(addJumbo(body.type));
    if (action === 'deleteJumbo') return jsonOut(deleteJumboRow(body.id));
    if (action === 'addEntry') return jsonOut(addEntry(body.jumboId, body.entry));
    if (action === 'markDamaged') return jsonOut(markDamaged(body.jumboId, body.recorder));
    if (action === 'importJumbo') return jsonOut(importJumbo(body.id, body.jno, body.type));
    if (action === 'importEntries') return jsonOut(importEntries(body.jumboId, body.entries));
    return jsonOut({ error: 'unknown action' });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/** รันฟังก์ชันนี้ 1 ครั้งจาก Apps Script editor ก่อนใช้งานจริง */
function setupSheets() {
  const jSheet = getSheet_(SHEET_JUMBOS);
  if (jSheet.getLastRow() === 0) {
    jSheet.appendRow(['id', 'jno', 'type', 'createdAt', 'status']);
    jSheet.getRange('A:A').setNumberFormat('@'); // บังคับคอลัมน์ id เป็นข้อความเสมอ
    const now = new Date().toISOString();
    jSheet.appendRow(['JB-SEED-TH01', 1, 'ข้าวไทย', now, 'active']);
    jSheet.appendRow(['JB-SEED-JP01', 1, 'ข้าวญี่ปุ่น', now, 'active']);
  }
  const uSheet = getSheet_(SHEET_USAGE);
  if (uSheet.getLastRow() === 0) {
    uSheet.appendRow(['jumboId', 'date', 'clean', 'complete', 'recorder', 'inspector', 'note', 'ts']);
    uSheet.getRange('A:A').setNumberFormat('@');
    uSheet.getRange('B:B').setNumberFormat('@'); // บังคับคอลัมน์วันที่เป็นข้อความ ไม่ให้ Sheets แปลงเป็น Date อัตโนมัติ
  }
}

function toDateText_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v;
}

function getAllData() {
  const jSheet = getSheet_(SHEET_JUMBOS);
  const jVals = jSheet.getDataRange().getValues();
  const jumbos = [];
  for (let i = 1; i < jVals.length; i++) {
    const r = jVals[i];
    if (!r[0]) continue;
    jumbos.push({ id: String(r[0]), jno: Number(r[1]), type: String(r[2]), status: String(r[4] || 'active') });
  }

  const uSheet = getSheet_(SHEET_USAGE);
  const uVals = uSheet.getDataRange().getValues();
  const usage = {};
  for (let i = 1; i < uVals.length; i++) {
    const r = uVals[i];
    if (!r[0]) continue;
    const id = String(r[0]);
    if (!usage[id]) usage[id] = [];
    usage[id].push({
      date: toDateText_(r[1]),
      clean: r[2] === true || r[2] === 'TRUE' || r[2] === 'true',
      complete: r[3] === true || r[3] === 'TRUE' || r[3] === 'true',
      recorder: String(r[4] || ''),
      inspector: String(r[5] || ''),
      note: String(r[6] || ''),
      ts: r[7]
    });
  }
  return { jumbos, usage };
}

function findJumboRow_(id) {
  const jSheet = getSheet_(SHEET_JUMBOS);
  const vals = jSheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === id) return { rowIndex: i + 1, type: String(vals[i][2]), status: String(vals[i][4] || 'active') };
  }
  return null;
}

function addJumbo(type) {
  const jSheet = getSheet_(SHEET_JUMBOS);
  const vals = jSheet.getDataRange().getValues();
  let maxJno = 0;
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][2] === type && Number(vals[i][1]) > maxJno) maxJno = Number(vals[i][1]);
  }
  const id = 'JB-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  const jno = maxJno + 1;
  jSheet.appendRow([id, jno, type, new Date().toISOString(), 'active']);
  return { id, jno, type, status: 'active' };
}

function deleteJumboRow(id) {
  const jSheet = getSheet_(SHEET_JUMBOS);
  const jVals = jSheet.getDataRange().getValues();
  for (let i = jVals.length - 1; i >= 1; i--) {
    if (String(jVals[i][0]) === id) { jSheet.deleteRow(i + 1); break; }
  }
  const uSheet = getSheet_(SHEET_USAGE);
  const uVals = uSheet.getDataRange().getValues();
  for (let i = uVals.length - 1; i >= 1; i--) {
    if (String(uVals[i][0]) === id) uSheet.deleteRow(i + 1);
  }
  return { ok: true };
}

function addEntry(jumboId, entry) {
  const jumbo = findJumboRow_(jumboId);
  if (!jumbo) return { ok: false, reason: 'not_found' };
  if (jumbo.status === 'damaged') return { ok: false, reason: 'damaged' };

  const uSheet = getSheet_(SHEET_USAGE);
  const vals = uSheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === jumboId) count++;
  }
  const maxUses = getMaxUses_(jumbo.type); // คำนวณจากชนิดข้าวจริงในชีต ไม่เชื่อค่าที่ฝั่งเว็บส่งมาเพื่อความถูกต้อง
  if (count >= maxUses) return { ok: false, reason: 'limit' };
  uSheet.appendRow([
    jumboId,
    entry.date,
    entry.clean ? true : false,
    entry.complete ? true : false,
    entry.recorder || '',
    entry.inspector || '',
    entry.note || '',
    entry.ts || Date.now()
  ]);
  return { ok: true };
}

/** ตีเป็นชำรุด: ล็อกจัมโบ้ใบนี้ไม่ให้บันทึกการใช้งานเพิ่มได้อีก พร้อมบันทึกประวัติว่าใครตี/เมื่อไหร่ */
function markDamaged(jumboId, recorder) {
  const jumbo = findJumboRow_(jumboId);
  if (!jumbo) return { ok: false, reason: 'not_found' };
  if (jumbo.status === 'damaged') return { ok: true, already: true };

  const jSheet = getSheet_(SHEET_JUMBOS);
  jSheet.getRange(jumbo.rowIndex, 5).setValue('damaged'); // คอลัมน์ E = status

  const uSheet = getSheet_(SHEET_USAGE);
  uSheet.appendRow([
    jumboId,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    false,
    false,
    recorder || '',
    '',
    'ตีเป็นชำรุด',
    Date.now()
  ]);
  return { ok: true };
}

/** ใช้สำหรับย้ายข้อมูลเดิมเข้ามาเท่านั้น (เช่น จาก Firebase) — คง id/jno เดิมไว้ตรงๆ ไม่สร้างใหม่ */
function importJumbo(id, jno, type) {
  const jSheet = getSheet_(SHEET_JUMBOS);
  const vals = jSheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === id) return { ok: true, skipped: 'already exists' };
  }
  jSheet.appendRow([id, jno, type, new Date().toISOString(), 'active']);
  return { ok: true };
}

/** ใช้สำหรับย้ายประวัติการใช้งานเดิมเข้ามาเท่านั้น — ไม่เช็คโควตาสูงสุด เพื่อให้คงประวัติเดิมไว้ครบ */
function importEntries(jumboId, entries) {
  if (!entries || !entries.length) return { ok: true, count: 0 };
  const uSheet = getSheet_(SHEET_USAGE);
  entries.forEach(entry => {
    uSheet.appendRow([
      jumboId,
      entry.date,
      entry.clean ? true : false,
      entry.complete ? true : false,
      entry.recorder || '',
      entry.inspector || '',
      entry.note || '',
      entry.ts || Date.now()
    ]);
  });
  return { ok: true, count: entries.length };
}
