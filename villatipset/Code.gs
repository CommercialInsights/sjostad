// Villatipset – ringlista. Apps Script kopplat till kalkylarket (Tillägg > Apps Script).
// 1) Projektinställningar > Skriptegenskaper: lägg till KEY med den nyckel ni vill dela med Fredrik.
// 2) Distribuera > Ny distribution > Webbapp. Kör som: Jag. Åtkomst: Alla.
// 3) Klistra in webbappens URL som API i index.html.
const EDIT = ['kontaktad', 'av', 'status', 'anteckning'];

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function keyOk_(k) { const K = PropertiesService.getScriptProperties().getProperty('KEY'); return !!K && k === K; }
function sheet_() { return SpreadsheetApp.getActive().getSheets()[0]; }

function doGet(e) {
  if (!keyOk_(e.parameter.key)) return json_({ error: 'fel_nyckel' });
  const v = sheet_().getDataRange().getValues();
  const h = v.shift();
  const rows = v.filter(r => r[0] !== '').map(r => {
    const o = {};
    h.forEach((k, i) => { o[k] = (r[i] instanceof Date) ? Utilities.formatDate(r[i], 'Europe/Stockholm', 'yyyy-MM-dd HH:mm') : r[i]; });
    o.kontaktad = o.kontaktad === true || String(o.kontaktad).toUpperCase() === 'TRUE';
    return o;
  });
  return json_({ rows: rows });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (!keyOk_(b.key)) return json_({ error: 'fel_nyckel' });
    const sh = sheet_();
    const v = sh.getDataRange().getValues();
    const h = v[0];
    const ri = v.findIndex((r, i) => i > 0 && r[0] === b.id);
    if (ri < 0) return json_({ error: 'okänt_id' });
    const t = Utilities.formatDate(new Date(), 'Europe/Stockholm', 'yyyy-MM-dd HH:mm');
    EDIT.forEach(k => { if (k in b) sh.getRange(ri + 1, h.indexOf(k) + 1).setValue(b[k]); });
    sh.getRange(ri + 1, h.indexOf('t') + 1).setValue(t);
    return json_({ ok: true, t: t });
  } finally { lock.releaseLock(); }
}
