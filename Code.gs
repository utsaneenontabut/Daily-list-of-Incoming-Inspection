/**
 * Daily list of Incoming Inspection — ตัวกลางข้อมูลบน Google Sheet (Apps Script)
 * วางโค้ดนี้ใน Extensions > Apps Script ของ Google Sheet แล้วทำตาม README
 */
const SH = { people: 'ผู้ตรวจ', appr: 'ผู้ปล่อย', state: 'สถานะ', daily: 'รายการประจำวัน' };
const PEOPLE_HEAD = ['ชื่อ', 'มา', 'Appearance', 'โควตา', 'Dimension', 'Function กลุ่ม 1', 'Function กลุ่ม 2'];
const APPR_HEAD = ['MRP', 'ผู้ปล่อย'];
const DAILY_COLS = ['id', 'batch', 'date', 'half', 'pk', 'name', 'type', 'po', 'qty', 'pack', 'vname', 'vno', 'country', 'lt', 'mrp', 'approver', 'pa', 'pd', 'pf', 'isNew', 'created', 'build', 'old', 'halal', 'tag', 'urgent', 'remark', 'app', 'dim', 'fun', 'status'];
const CHUNK = 40000;   // ช่องเซลล์เก็บได้ไม่เกิน 50,000 ตัวอักษร

function sh_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); if (head) s.appendRow(head); }
  return s;
}
function token_() { return PropertiesService.getScriptProperties().getProperty('TOKEN'); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** รันครั้งเดียว: สร้างชีตและรหัสผ่านของระบบ แล้วดูรหัสที่เมนู ดู > บันทึกการดำเนินการ */
function setup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('TOKEN')) p.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, ''));
  sh_(SH.people, PEOPLE_HEAD); sh_(SH.appr, APPR_HEAD); sh_(SH.state, ['key', 'value']);
  const d = sh_(SH.daily, DAILY_COLS);
  d.getRange(1, 1, d.getMaxRows(), DAILY_COLS.length).setNumberFormat('@');   // เก็บเป็นข้อความ กันวันที่/รหัสถูกแปลง
  sh_(SH.people).getRange(1, 1, 500, 7).setNumberFormat('@');
  sh_(SH.appr).getRange(1, 1, 500, 2).setNumberFormat('@');
  Logger.log('TOKEN = ' + p.getProperty('TOKEN'));
}

function rows_(name, ncol) {
  const s = sh_(name); const n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, ncol).getDisplayValues().filter(r => String(r[0]).trim() !== '');
}
function kvGet_() {
  const s = sh_(SH.state, ['key', 'value']), n = s.getLastRow(), out = {};
  if (n < 2) return out;
  s.getRange(2, 1, n - 1, 2).getValues().forEach(r => {
    const m = String(r[0]).match(/^(.*)#(\d+)$/);
    if (m) { (out[m[1]] = out[m[1]] || [])[+m[2]] = r[1]; } else out[r[0]] = [r[1]];
  });
  Object.keys(out).forEach(k => { out[k] = out[k].join(''); });
  return out;
}
function kvSet_(obj) {
  const s = sh_(SH.state, ['key', 'value']), rows = [];
  Object.keys(obj).forEach(k => {
    const v = String(obj[k] == null ? '' : obj[k]);
    for (let i = 0, c = 0; i === 0 || i < v.length; i += CHUNK, c++) rows.push([k + '#' + c, v.slice(i, i + CHUNK)]);
  });
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 2).clearContent();
  if (rows.length) { s.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows); }
}
function dailyAll_() {
  const s = sh_(SH.daily, DAILY_COLS), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, DAILY_COLS.length).getDisplayValues().filter(r => r[0] !== '').map(r => { const o = {}; DAILY_COLS.forEach((c, i) => { o[c] = r[i]; }); return o; });
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.token !== token_()) return json_({ error: 'bad token' });
  return json_({ ok: true, people: rows_(SH.people, 7), appr: rows_(SH.appr, 2), kv: kvGet_(), daily: dailyAll_() });
}

function doPost(e) {
  let b; try { b = JSON.parse(e.postData.contents); } catch (x) { return json_({ error: 'bad json' }); }
  if (b.token !== token_()) return json_({ error: 'bad token' });
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (b.action === 'state') {
      const p = sh_(SH.people, PEOPLE_HEAD), a = sh_(SH.appr, APPR_HEAD);
      if (p.getLastRow() > 1) p.getRange(2, 1, p.getLastRow() - 1, 7).clearContent();
      if ((b.people || []).length) p.getRange(2, 1, b.people.length, 7).setNumberFormat('@').setValues(b.people);
      if (a.getLastRow() > 1) a.getRange(2, 1, a.getLastRow() - 1, 2).clearContent();
      if ((b.appr || []).length) a.getRange(2, 1, b.appr.length, 2).setNumberFormat('@').setValues(b.appr);
      kvSet_(Object.assign({}, b.kv || {}, { ts: String(Date.now()) }));
    } else if (b.action === 'daily') {
      const s = sh_(SH.daily, DAILY_COLS), n = s.getLastRow();
      const ids = {}; if (n > 1) s.getRange(2, 1, n - 1, 1).getValues().forEach((r, i) => { ids[r[0]] = i + 2; });
      const add = [];
      (b.rows || []).forEach(o => {
        const row = DAILY_COLS.map(c => (o[c] === undefined || o[c] === null) ? '' : String(o[c]));
        if (ids[o.id]) s.getRange(ids[o.id], 1, 1, DAILY_COLS.length).setNumberFormat('@').setValues([row]); else add.push(row);
      });
      if (add.length) s.getRange(s.getLastRow() + 1, 1, add.length, DAILY_COLS.length).setNumberFormat('@').setValues(add);
    } else if (b.action === 'dailyDel') {
      const s = sh_(SH.daily, DAILY_COLS), set = {}; (b.ids || []).forEach(i => { set[i] = 1; });
      for (let r = s.getLastRow(); r >= 2; r--) if (set[s.getRange(r, 1).getValue()]) s.deleteRow(r);
    } else return json_({ error: 'unknown action' });
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}
