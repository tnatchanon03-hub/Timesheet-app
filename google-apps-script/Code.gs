// Timesheet cloud sync — วางโค้ดนี้ใน Google Apps Script ของ Google Sheet (ส่วนขยาย > Apps Script)
// แล้ว Deploy เป็น Web app (ดูขั้นตอนใน SYNC_SETUP.md)

// ตั้งรหัสลับของคุณเอง (ต้องตรงกับที่กรอกในแอป แท็บ ⚙️ ตั้งค่า > ซิงก์ข้อมูลขึ้น Google Sheets)
const SECRET = 'CHANGE_ME';

const DATA_SHEET = 'data';          // เก็บข้อมูลดิบของแอป (ห้ามแก้ไขด้วยมือ)
const VIEW_SHEET = 'ตารางเวลา';     // ตารางอ่านง่าย สร้างใหม่ทุกครั้งที่มีการซิงก์

function doGet(e) {
  if (!e || !e.parameter || e.parameter.secret !== SECRET) return json_({ ok: false, error: 'unauthorized' });
  const rows = readRows_(dataSheet_());
  const data = {};
  Object.keys(rows).forEach(k => { data[k] = rows[k].value; });
  return json_({ ok: true, data: data });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad request' }); }
  if (body.secret !== SECRET) return json_({ ok: false, error: 'unauthorized' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = dataSheet_();
    const rows = readRows_(sh);
    const items = body.items || {};
    const now = new Date();
    const appended = [];
    Object.keys(items).forEach(k => {
      if (String(k).indexOf('timesheet_') !== 0) return;
      const v = String(items[k]);
      if (rows[k]) {
        sh.getRange(rows[k].row, 2, 1, 2).setValues([[v, now]]);
      } else {
        appended.push([k, v, now]);
      }
    });
    if (appended.length) sh.getRange(sh.getLastRow() + 1, 1, appended.length, 3).setValues(appended);
    rebuildView_();
    return json_({ ok: true, saved: Object.keys(items).length });
  } finally {
    lock.releaseLock();
  }
}

function dataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['key', 'value', 'updatedAt']]);
    sh.setFrozenRows(1);
  }
  // เก็บเป็นข้อความล้วน กัน Sheets แปลง JSON/ตัวเลขเอง
  sh.getRange('A:B').setNumberFormat('@');
  return sh;
}

function readRows_(sh) {
  const result = {};
  const last = sh.getLastRow();
  if (last < 2) return result;
  const values = sh.getRange(2, 1, last - 1, 2).getValues();
  values.forEach((r, i) => { if (r[0]) result[r[0]] = { row: i + 2, value: String(r[1]) }; });
  return result;
}

function rebuildView_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = readRows_(dataSheet_());
  const out = [];
  Object.keys(rows).forEach(k => {
    const m = /^timesheet_(\d{4})_(\d{1,2})$/.exec(k);
    if (!m) return;
    let data;
    try { data = JSON.parse(rows[k].value); } catch (err) { return; }
    if (!data || !data.days) return;
    const y = Number(m[1]), mo = Number(m[2]);
    Object.keys(data.days).forEach(d => {
      const day = data.days[d];
      if (!/^\d+$/.test(d) || !day || (!day.in && !day.out && !day.note)) return;
      if (day.note) out.push([new Date(y, mo, Number(d)), '', '', day.note]);
      else out.push([new Date(y, mo, Number(d)), day.in || '', day.out || '', '']);
    });
  });
  out.sort((a, b) => a[0] - b[0]);

  let sh = ss.getSheetByName(VIEW_SHEET);
  if (!sh) sh = ss.insertSheet(VIEW_SHEET, 0);
  sh.clearContents();
  sh.getRange(1, 1, 1, 4).setValues([['วันที่', 'เวลาเข้า', 'เวลาออก', 'หมายเหตุ']]);
  sh.setFrozenRows(1);
  if (out.length) {
    sh.getRange(2, 2, out.length, 2).setNumberFormat('@');
    sh.getRange(2, 1, out.length, 4).setValues(out);
    sh.getRange(2, 1, out.length, 1).setNumberFormat('dd/mm/yyyy');
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
