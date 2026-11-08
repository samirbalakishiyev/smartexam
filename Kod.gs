/* SmartExam · Google Cədvəl skripti (Apps Script)
   Created by Samir Balakişiyev · WhatsApp 055 795 86 25

   QURAŞDIRMA (bir dəfə):
   1) Yeni Google Cədvəl açın (məs. "SmartExam"). Uzantılar → Apps Script.
   2) Bu faylın hamısını Kod.gs-ə yapışdırın. Aşağıda PASS sətrində öz şifrənizi yazın.
   3) ⚙️ Layihə parametrləri → Skript xassələri → "CLAUDE_KEY" adı ilə Claude API açarınızı əlavə edin
      (My New Vocabulary skriptindəki eyni açar ola bilər).
   4) Deploy → New deployment → Web app: Execute as = Me, Who has access = Anyone → Deploy, icazələri verin.
   5) Verilən …/exec linkini SmartExam saytında Parametrlər → «Google Cədvəl bağlantısı»na yazın, şifrəni də yazın, «Qoşul».
   Yeniləmək: Deploy → Manage deployments → ✏️ → Version: New version → Deploy (link dəyişmir).

   Vərəqlər:
   - "Data"      — tətbiqin bütün məlumatı (qruplar, şagirdlər, sınaqlar, nəticələr, parametrlər). Əl ilə dəyişməyin.
   - "Nəticələr" — oxumaq üçün cədvəl: sınaq, şagird, test, paraqraf, faiz. Avtomatik yenilənir. */

const PASS = 'BURAYA-SIFRE-YAZIN';      /* 🔐 sayta girmək üçün şifrə — heç kimə deməyin */
const AI_MODEL = 'claude-opus-5-5';
const COLLS = ['groups', 'students', 'exams', 'results', 'settings'];
const CHUNK = 45000;                    /* bir xanaya ən çox ~50 000 simvol sığır */
const RES_HEAD = ['ID', 'Sınaq', 'Tarix', 'Şagird', 'Qrup', 'Test (25)', 'Paraqraf (5)', 'Faiz', 'Yoxlanmalı suallar', 'Yenilənib'];

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('SmartExam skripti işləyir.');
}

function doPost(e) {
  let o = {};
  try { o = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { return out({ ok: false, error: 'bad' }); }
  const auth = checkPass_(o.p);
  if (auth) return out(auth);
  try {
    if (o.a === 'ai') return out(ai_(o));          /* süni intellekt gözlənilərkən cədvəl bağlanmır */
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (o.a === 'load') return out({ ok: true, data: load_() });
      if (o.a === 'put') { if (!COLLS.includes(o.c) || !o.obj || !o.obj.id) return out({ ok: false, error: 'bad' }); putDoc_(o.c, String(o.obj.id), o.obj); if (o.c === 'results' && Array.isArray(o.view)) putView_(String(o.obj.id), o.view); return out({ ok: true }); }
      if (o.a === 'del') { if (!COLLS.includes(o.c)) return out({ ok: false, error: 'bad' }); delDoc_(o.c, String(o.id)); if (o.c === 'results') delView_(String(o.id)); return out({ ok: true }); }
      if (o.a === 'set') { putDoc_('settings', 'main', o.settings || {}); return out({ ok: true }); }
      return out({ ok: false, error: 'a' });
    } finally { lock.releaseLock(); }
  } catch (err) {
    console.error(err && err.stack || err);
    return out({ ok: false, error: String(err && err.message || 'api').slice(0, 40) });
  }
}

/* 🔐 şifrə: 5 səhv cəhddən sonra 15 dəqiqə bağlanır */
function checkPass_(p) {
  if (!PASS || PASS === 'BURAYA-SIFRE-YAZIN') return { ok: false, error: 'nopass' };
  const c = CacheService.getScriptCache(), n = +(c.get('se_fail') || 0);
  if (n >= 5) return { ok: false, error: 'wait' };
  if (String(p || '') !== PASS) { c.put('se_fail', String(n + 1), 900); return { ok: false, error: 'pass', left: 4 - n }; }
  if (n) c.remove('se_fail');
  return null;
}

/* ---------- Data vərəqi: A kolleksiya · B id · C yenilənib · D.. JSON hissələri ---------- */
function dataSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Data');
  if (!sh) { sh = ss.insertSheet('Data'); sh.getRange(1, 1, 1, 4).setValues([['Kolleksiya', 'ID', 'Yenilənib', 'JSON']]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function load_() {
  const sh = dataSheet_(), last = sh.getLastRow(), data = { groups: [], students: [], exams: [], results: [], settings: {} };
  if (last < 2) return data;
  const w = Math.max(4, sh.getLastColumn());
  sh.getRange(2, 1, last - 1, w).getValues().forEach(r => {
    const c = String(r[0]); if (!COLLS.includes(c)) return;
    const json = r.slice(3).map(x => String(x || '')).join('');
    let v; try { v = JSON.parse(json); } catch (e) { return; }
    if (c === 'settings') data.settings = v; else data[c].push(v);
  });
  return data;
}
function findRow_(sh, c, id) {
  const hits = sh.getRange('B:B').createTextFinder(id).matchEntireCell(true).findAll();
  for (const h of hits) { const r = h.getRow(); if (r > 1 && sh.getRange(r, 1).getValue() === c) return r; }
  return 0;
}
function putDoc_(c, id, obj) {
  const sh = dataSheet_(), json = JSON.stringify(obj), parts = [];
  for (let i = 0; i < json.length; i += CHUNK) parts.push(json.slice(i, i + CHUNK));
  if (!parts.length) parts.push('{}');
  let r = findRow_(sh, c, id);
  if (!r) r = sh.getLastRow() + 1;
  const w = Math.max(4, sh.getLastColumn());
  if (w > 3) sh.getRange(r, 4, 1, w - 3).clearContent();
  sh.getRange(r, 1, 1, 3 + parts.length).setNumberFormat('@').setValues([[c, id, new Date().toISOString()].concat(parts)]);
}
function delDoc_(c, id) {
  const sh = dataSheet_(), r = findRow_(sh, c, id);
  if (r) sh.deleteRow(r);
}

/* ---------- Nəticələr vərəqi (oxumaq üçün) ---------- */
function resSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Nəticələr');
  if (!sh) { sh = ss.insertSheet('Nəticələr', 0); sh.getRange(1, 1, 1, RES_HEAD.length).setValues([RES_HEAD]).setFontWeight('bold'); sh.setFrozenRows(1); sh.hideColumns(1); }
  return sh;
}
function viewRow_(sh, id) {
  const h = sh.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
  return h && h.getRow() > 1 ? h.getRow() : 0;
}
function putView_(id, v) {
  const sh = resSheet_();
  const row = [id].concat(v.slice(0, 8).map(x => x === null || x === undefined ? '' : x));
  while (row.length < RES_HEAD.length - 1) row.push('');
  row.push(new Date());
  const r = viewRow_(sh, id) || sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, RES_HEAD.length).setValues([row]);
}
function delView_(id) {
  const sh = resSheet_(), r = viewRow_(sh, id);
  if (r) sh.deleteRow(r);
}

/* ---------- Claude: vərəq şəkillərinin oxunması və paraqrafın yoxlanması ---------- */
function ai_(o) {
  const key = PropertiesService.getScriptProperties().getProperty('CLAUDE_KEY');
  if (!key) return { ok: false, error: 'nokey' };
  const imgs = (Array.isArray(o.images) ? o.images : []).slice(0, 4).filter(i => i && i.data);
  if (Array.isArray(o.images) && o.images.length && !imgs.length) return { ok: false, error: 'noimg' };
  const content = imgs.map(i => ({ type: 'image', source: { type: 'base64', media_type: /^image\/(jpeg|png|webp|gif)$/.test(i.type) ? i.type : 'image/jpeg', data: i.data } }));
  content.push({ type: 'text', text: String(o.prompt || '').slice(0, 60000) });
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: AI_MODEL, max_tokens: o.big ? 6000 : 2000, messages: [{ role: 'user', content: content }] })
  });
  const code = res.getResponseCode(), body = res.getContentText();
  if (code !== 200) {
    console.error('Claude API ' + code + ': ' + body.slice(0, 500));
    return { ok: false, error: code === 401 ? 'badkey' : (code === 400 && /credit/i.test(body)) ? 'credit' : 'api' + code };
  }
  const j = JSON.parse(body), t = (j.content || []).filter(c => c.type === 'text').map(c => c.text).join('');
  const a = t.indexOf('{'), b = t.lastIndexOf('}');
  if (a < 0 || b < a) { console.error('Cavab: ' + t.slice(0, 800)); return { ok: false, error: 'json' }; }
  try { return { ok: true, json: JSON.parse(t.slice(a, b + 1)) }; }
  catch (e) { console.error('JSON: ' + t.slice(0, 800)); return { ok: false, error: 'json' }; }
}
