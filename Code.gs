// วางโค้ดนี้ใน Google Sheet > ส่วนขยาย > Apps Script แล้วทำให้ใช้งานได้ (Web app, ผู้เข้าถึง: ทุกคน)
const PIN = '';            // ตั้งรหัสถ้าต้องการ (ต้องตรงกับ CFG.PIN ใน index.html)
const NAME = 'Attendance';

function sh_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(NAME);
  if (!s) {
    s = ss.insertSheet(NAME);
    s.appendRow(['date', 'period', 'no', 'status', 'updated']);
    s.getRange('A:A').setNumberFormat('@');
  }
  return s;
}
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet(e) {
  const q = e.parameter;
  if (PIN && q.pin !== PIN) return out_([]);
  const PR = ['a', 's', 'l', 't', 'p'], m = {}, o = [];
  sh_().getDataRange().getValues().slice(1).forEach(r => {
    const d = String(r[0]);
    if (!r[3] || d < q.from || d > q.to) return;
    if (q.agg) {                                  // รวมเป็นสถานะรายวัน (หนักสุด) ลดขนาดข้อมูล
      const k = d + '|' + r[2];
      if (!(k in m) || PR.indexOf(r[3]) < PR.indexOf(m[k][3])) m[k] = [d, '*', +r[2], r[3]];
    } else o.push([d, +r[1], +r[2], r[3]]);
  });
  return out_(q.agg ? Object.values(m) : o);
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (PIN && b.pin !== PIN) return out_({ ok: false });
  const L = LockService.getScriptLock();
  L.waitLock(20000);
  const s = sh_(), idx = {};
  s.getDataRange().getValues().forEach((r, i) => idx[r[0] + '|' + r[1] + '|' + r[2]] = i + 1);
  b.items.forEach(x => {
    const k = x.d + '|' + x.p + '|' + x.n;
    if (idx[k]) s.getRange(idx[k], 4, 1, 2).setValues([[x.c, new Date()]]);
    else if (x.c) { s.appendRow([x.d, x.p, x.n, x.c, new Date()]); idx[k] = s.getLastRow(); }
  });
  L.releaseLock();
  return out_({ ok: true });
}
