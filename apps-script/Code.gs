/**
 * Jason 工作追蹤｜後端（Google Apps Script）
 *
 * 網頁放在 GitHub Pages（APP_URL），用公司 Google 帳號登入，這裡只負責驗證身分和讀寫試算表。
 *
 * 工作表：
 *   人員      ─ 信箱、姓名、身分（管理者／攝影師／檢視者）
 *   設定      ─ 每月一列：品牌配額、工時預算等
 *   明細      ─ 每支影片一列
 *   異動紀錄  ─ 誰在什麼時間改了什麼
 *
 * 權限：
 *   管理者  ─ 全部可改：配額、工時、明細、由配額重建明細
 *   攝影師  ─ 明細全部可改：新增／刪除影片、名稱、品牌、分級、類型、交期、狀態
 *   檢視者  ─ 只能看。允許網域（公司信箱）的人就算不在「人員」名單上，登入後也是檢視者
 *
 * 部署：網頁應用程式，執行身分「我」，誰可以存取「所有人」。
 */

/* ---------- 要確認的設定 ---------- */
const CLIENT_ID = '565026992827-outdn7hb027dounk7cdt07go54r7g116.apps.googleusercontent.com'; // 與打卡系統共用
const ALLOWED_DOMAINS = ['neverdie.com.tw', 'riseme.com.tw'];
const APP_URL = 'https://jinrayyork.github.io/jason/';
const SESSION_DAYS = 30;

const SHEET = { people: '人員', cfg: '設定', rows: '明細', log: '異動紀錄' };
const BRANDS = ['Gymray', 'RISEME', 'NEVERDIE'];
const GRADES = ['A', 'B', 'C'];
const ALL_GRADES = ['A', 'B', 'C', 'D'];
const TYPES = ['Reels', '季度主片', '其他'];
const STATUSES = ['待製作', '拍攝中', '剪輯中', '待審核', '修改中', '已完成'];
const ROLE_BY_LABEL = { '管理者': 'admin', '攝影師': 'editor', '檢視者': 'viewer' };
const ROLE_NAME = { admin: '管理者', editor: '攝影師', viewer: '檢視者' };
// 後來加的設定放在最後幾欄，才不會打亂舊資料：[鍵, 欄名, 預設值, 上限]
const EXTRA = [['basepts', '基準點數', 28, 500], ['ptsA', 'A級點數', 1, 50], ['ptsB', 'B級點數', 2, 50], ['ptsC', 'C級點數', 4, 50]];
const BASE = { matrix: [[3, 3, 0], [1, 3, 1], [0, 2, 1]], hours: 40, available: 160, other: 32, edit: 3, dcount: 0, dpoints: 12 };
EXTRA.forEach(x => { BASE[x[0]] = x[2]; });
const NUM_KEYS = ['hours', 'available', 'other', 'edit', 'dcount', 'dpoints'];

const PEOPLE_HEAD = ['信箱', '姓名', '身分'];
const CFG_HEAD = ['月份']
  .concat(...BRANDS.map(b => GRADES.map(g => b + ' ' + g + '級')))
  .concat(['外拍時數', '正常工時預算', '其他預排工時', '每點後製工時', 'D級專案數', 'D級每案點數', '更新時間', '更新者'], EXTRA.map(x => x[1]));
const EXTRA_COL = 18; // EXTRA 從第 19 欄開始（0 起算為 18）
const ROW_HEAD = ['ID', '月份', '影片／專案名稱', '品牌', '分級', '類型', '交期', '狀態', '完成', '更新時間', '更新者'];
const LOG_HEAD = ['時間', '人員', '動作', '月份', '內容'];

/* ---------- 網頁入口 ---------- */

// 直接打開 Apps Script 網址時，引導到 GitHub 上的正式網址
function doGet() {
  return HtmlService.createHtmlOutput(
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font-family:system-ui,sans-serif;padding:40px 20px;text-align:center">' +
    '<h2>Jason 工作追蹤</h2><p>請改用這個網址：</p>' +
    '<p><a href="' + APP_URL + '" target="_top">' + APP_URL + '</a></p></div>'
  ).setTitle('Jason 工作追蹤');
}

// GitHub 上的網頁都透過這裡呼叫：{ fn, args, session, idToken }
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = String(body.fn || '');
    const args = Array.isArray(body.args) ? body.args : [];
    if (fn === 'login') {
      const email = verifyGoogleIdToken_(body.idToken);
      const me = person_(email);
      out = { ok: true, session: createSession_(email), data: getState_(me, args[0], null) };
    } else if (fn === 'logout') {
      deleteSession_(body.session);
      out = { ok: true };
    } else {
      const me = person_(sessionEmail_(body.session));
      if (!API[fn]) throw new Error('不支援的操作：' + fn);
      out = { ok: true, data: API[fn].apply(null, [me].concat(args)) };
    }
  } catch (err) {
    out = { ok: false, auth: !!(err && err.auth), error: (err && err.message) || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

const API = {
  getState: (me, month, knownVersion) => getState_(me, month, knownVersion),
  saveConfig: saveConfig_,
  updateRow: updateRow_,
  addRow: addRow_,
  deleteRow: deleteRow_,
  rebuildRows: rebuildRows_
};

/* ---------- 試算表選單與初始化 ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Jason 工作追蹤')
    .addItem('① 初始化工作表', 'setup')
    .addItem('讓所有人重新登入', 'clearSessions')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const people = ensureSheet_(ss, SHEET.people, PEOPLE_HEAD);
  const cfg = ensureSheet_(ss, SHEET.cfg, CFG_HEAD);
  const rows = ensureSheet_(ss, SHEET.rows, ROW_HEAD);
  const log = ensureSheet_(ss, SHEET.log, LOG_HEAD);
  // 文字格式：避免 Google 試算表把「2026-11」「2026-11-20」自動轉成日期
  cfg.getRange('A:A').setNumberFormat('@');
  cfg.getRange(1, CFG_HEAD.indexOf('更新時間') + 1, cfg.getMaxRows(), 1).setNumberFormat('@');
  ['A:A', 'B:B', 'G:G', 'J:J'].forEach(a => rows.getRange(a).setNumberFormat('@'));
  ['A:A', 'D:D'].forEach(a => log.getRange(a).setNumberFormat('@'));
  // 「身分」欄做成下拉選單
  people.getRange(2, 3, people.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(ROLE_BY_LABEL), true).build());
  // 把自己加成管理者
  const owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (owner && !readPeople_()[owner]) people.appendRow([owner, '', '管理者']);
  ss.setActiveSheet(people);
}

/** 第一次使用或改版後執行一次：讓 Google 跳出「連線到外部服務」的授權（登入驗證需要） */
function authorize() {
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=test', { muteHttpExceptions: true });
  Logger.log('授權完成，可以關掉這個視窗了');
}

function clearSessions() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(props.getProperties()).forEach(k => { if (k.indexOf('sess_') === 0) props.deleteProperty(k); });
}

function ensureSheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function sheet_(name, head) {
  return ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), name, head);
}

/* ---------- 登入與身分 ---------- */

function authError_(msg) { const e = new Error(msg); e.auth = true; return e; }

/** 向 Google 驗證登入憑證，回傳已確認的公司信箱 */
function verifyGoogleIdToken_(idToken) {
  if (!idToken) throw authError_('登入失敗，請再試一次。');
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw authError_('登入憑證無效或已過期，請重新登入。');
  const p = JSON.parse(res.getContentText());
  const email = String(p.email || '').toLowerCase();
  const domain = email.split('@')[1] || '';
  if (p.aud !== CLIENT_ID) throw authError_('登入憑證不屬於這個系統。');
  if (['accounts.google.com', 'https://accounts.google.com'].indexOf(p.iss) < 0) throw authError_('登入憑證來源不正確。');
  if (String(p.email_verified) !== 'true') throw authError_('這個 Google 帳號的信箱尚未驗證。');
  if (Number(p.exp) * 1000 < Date.now()) throw authError_('登入已過期，請重新登入。');
  if (ALLOWED_DOMAINS.indexOf(domain) < 0) throw authError_('請用公司信箱登入（' + ALLOWED_DOMAINS.join('、') + '）。');
  // hd 由 Google 填入，代表是公司 Workspace 管理的帳號，不是用公司信箱另外註冊的個人帳號
  if (String(p.hd || '').toLowerCase() !== domain) throw authError_('請用公司的 Google Workspace 帳號登入。');
  return email;
}

/* 登入後發一組 30 天有效的通行碼，之後每次操作只帶通行碼 */
function createSession_(email) {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  Object.entries(props.getProperties()).forEach(([k, v]) => {
    if (k.indexOf('sess_') === 0) { try { if (JSON.parse(v).exp < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); } }
  });
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  props.setProperty('sess_' + token, JSON.stringify({ email: email, exp: now + SESSION_DAYS * 86400000 }));
  return token;
}

function sessionEmail_(token) {
  token = String(token || '');
  if (!/^[0-9a-f]{64}$/.test(token)) throw authError_('請先登入。');
  const raw = PropertiesService.getScriptProperties().getProperty('sess_' + token);
  if (!raw) throw authError_('登入已失效，請重新登入。');
  const s = JSON.parse(raw);
  if (s.exp < Date.now()) { deleteSession_(token); throw authError_('登入已過期，請重新登入。'); }
  return s.email;
}

function deleteSession_(token) {
  token = String(token || '');
  if (/^[0-9a-f]{64}$/.test(token)) PropertiesService.getScriptProperties().deleteProperty('sess_' + token);
}

function readPeople_() {
  const map = {};
  sheet_(SHEET.people, PEOPLE_HEAD).getDataRange().getDisplayValues().slice(1).forEach(r => {
    const email = String(r[0] || '').trim().toLowerCase();
    if (email) map[email] = { name: String(r[1] || '').trim(), role: ROLE_BY_LABEL[String(r[2] || '').trim()] || 'viewer' };
  });
  return map;
}

// 每次操作都重新查「人員」分頁，改身分立即生效
function person_(email) {
  const p = readPeople_()[email];
  const domain = email.split('@')[1] || '';
  if (!p && ALLOWED_DOMAINS.indexOf(domain) < 0) throw authError_('這個帳號沒有使用權限。');
  return { email: email, name: (p && p.name) || email.split('@')[0], role: p ? p.role : 'viewer' };
}

function require_(me, roles) {
  if (roles.indexOf(me.role) < 0) throw new Error('你的身分是「' + ROLE_NAME[me.role] + '」，沒有權限進行這個操作。');
}

/* ---------- 版本號（讓網頁只在資料有變時才重新載入） ---------- */

function version_() {
  return Number(PropertiesService.getScriptProperties().getProperty('VERSION') || 0);
}
function bump_() {
  PropertiesService.getScriptProperties().setProperty('VERSION', String(version_() + 1));
}

/* ---------- 讀取 ---------- */

function num_(v) { const n = Number(String(v).replace(/,/g, '')); return isFinite(n) && n > 0 ? n : 0; }

function readCfgMap_() {
  const vals = sheet_(SHEET.cfg, CFG_HEAD).getDataRange().getDisplayValues().slice(1);
  const map = {};
  vals.forEach((r, i) => {
    const m = String(r[0]).trim();
    if (!/^\d{4}-\d{2}$/.test(m)) return;
    const c = { matrix: BRANDS.map((b, bi) => GRADES.map((g, gi) => num_(r[1 + bi * 3 + gi]))) };
    NUM_KEYS.forEach((k, ki) => { c[k] = num_(r[10 + ki]); });
    EXTRA.forEach((x, xi) => { const v = r[EXTRA_COL + xi]; c[x[0]] = v === '' || v == null ? x[2] : num_(v); });
    map[m] = { row: i + 2, cfg: c };
  });
  return map;
}

function readRows_() {
  return sheet_(SHEET.rows, ROW_HEAD).getDataRange().getDisplayValues().slice(1).map((r, i) => ({
    _row: i + 2,
    id: r[0], month: String(r[1]).trim(), name: r[2], brand: r[3], grade: r[4], type: r[5],
    due: r[6], status: r[7], done: r[8] === '是', updatedAt: r[9], updatedBy: r[10]
  })).filter(r => r.id);
}

function inheritCfg_(map, month) {
  const keys = Object.keys(map).sort();
  const earlier = keys.filter(k => k < month);
  const src = earlier.length ? earlier[earlier.length - 1] : keys[keys.length - 1];
  return JSON.parse(JSON.stringify(src ? map[src].cfg : BASE));
}

function thisMonth_() {
  return Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM');
}

function getState_(me, month, knownVersion) {
  month = month ? cleanMonth_(month) : thisMonth_();
  const version = version_();
  const who = { email: me.email, name: me.name, role: me.role };
  if (knownVersion !== null && knownVersion !== undefined && Number(knownVersion) === version) {
    return { same: true, version: version, me: who };
  }
  const map = readCfgMap_();
  const all = readRows_();
  const months = {};
  Object.keys(map).forEach(m => { months[m] = 1; });
  all.forEach(r => { if (/^\d{4}-\d{2}$/.test(r.month)) months[r.month] = 1; });
  return {
    same: false, version: version, me: who, month: month,
    months: Object.keys(months).sort().reverse(),
    cfg: map[month] ? map[month].cfg : inheritCfg_(map, month),
    cfgSaved: !!map[month],
    rows: all.filter(r => r.month === month).map(r => { const o = Object.assign({}, r); delete o._row; return o; })
  };
}

/* ---------- 寫入 ---------- */

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { const out = fn(); bump_(); return out; }
  finally { lock.releaseLock(); }
}

function stamp_() {
  return Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
}

function log_(me, action, month, detail) {
  sheet_(SHEET.log, LOG_HEAD).appendRow([stamp_(), me.name + '（' + me.email + '）', action, month, String(detail).slice(0, 500)]);
}

function cleanMonth_(m) {
  m = String(m || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(m)) throw new Error('月份格式錯誤：' + m);
  return m;
}

function cleanCfg_(c) {
  c = c || {};
  const out = { matrix: BRANDS.map((b, i) => GRADES.map((g, j) => Math.min(100, Math.round(num_(((c.matrix || [])[i] || [])[j]))))) };
  NUM_KEYS.forEach(k => { out[k] = Math.min(300, num_(c[k])); });
  EXTRA.forEach(x => { out[x[0]] = c[x[0]] === undefined ? x[2] : Math.min(x[3], Math.round(num_(c[x[0]]) * 10) / 10); });
  return out;
}

function safeText_(s) {
  s = String(s == null ? '' : s).slice(0, 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s; // 避免被當成公式
}

function cleanRow_(o) {
  const r = {};
  if ('name' in o) r.name = String(o.name || '').trim().slice(0, 200) || '未命名影片';
  if ('brand' in o) { if (BRANDS.indexOf(o.brand) < 0) throw new Error('品牌不正確'); r.brand = o.brand; }
  if ('grade' in o) { if (ALL_GRADES.indexOf(o.grade) < 0) throw new Error('分級不正確'); r.grade = o.grade; }
  if ('type' in o) { if (TYPES.indexOf(o.type) < 0) throw new Error('類型不正確'); r.type = o.type; }
  if ('due' in o) { const d = String(o.due || ''); if (d && !/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('交期格式不正確'); r.due = d; }
  if ('status' in o) { if (STATUSES.indexOf(o.status) < 0) throw new Error('狀態不正確'); r.status = o.status; }
  if ('done' in o) r.done = !!o.done;
  return r;
}

function rowArr_(o, me) {
  return [o.id, o.month, safeText_(o.name), o.brand, o.grade, o.type, o.due, o.status, o.done ? '是' : '否', stamp_(), me.name];
}

function newId_() { return 'v' + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

function writeCfg_(month, cfg, me) {
  const sh = sheet_(SHEET.cfg, CFG_HEAD);
  const map = readCfgMap_();
  sh.getRange(1, 1, 1, CFG_HEAD.length).setValues([CFG_HEAD]); // 舊表補上「基準點數」欄名
  const line = [month].concat([].concat(...cfg.matrix), NUM_KEYS.map(k => cfg[k]), [stamp_(), me.name], EXTRA.map(x => cfg[x[0]]));
  if (map[month]) sh.getRange(map[month].row, 1, 1, line.length).setValues([line]);
  else sh.appendRow(line);
}

function saveConfig_(me, month, cfg) {
  require_(me, ['admin']);
  month = cleanMonth_(month);
  cfg = cleanCfg_(cfg);
  withLock_(() => {
    writeCfg_(month, cfg, me);
    log_(me, '更新設定', month, JSON.stringify(cfg));
  });
  return getState_(me, month, null);
}

function updateRow_(me, id, patch) {
  require_(me, ['admin', 'editor']);
  const clean = cleanRow_(patch || {});
  let month;
  withLock_(() => {
    const row = readRows_().filter(r => r.id === id)[0];
    if (!row) throw new Error('找不到這支影片，可能已被刪除，請重新整理。');
    month = row.month;
    Object.assign(row, clean);
    if ('status' in clean && !('done' in clean)) row.done = row.status === '已完成';
    if ('done' in clean && !('status' in clean)) row.status = row.done ? '已完成' : '剪輯中';
    sheet_(SHEET.rows, ROW_HEAD).getRange(row._row, 1, 1, ROW_HEAD.length).setValues([rowArr_(row, me)]);
    log_(me, '修改影片', month, row.name + '｜' + JSON.stringify(clean));
  });
  return getState_(me, month, null);
}

function addRow_(me, month, data) {
  require_(me, ['admin', 'editor']);
  month = cleanMonth_(month);
  const base = { name: '新增影片', brand: 'Gymray', grade: 'B', type: 'Reels', due: '', status: '待製作', done: false };
  const row = Object.assign(base, cleanRow_(data || {}), { id: newId_(), month: month });
  withLock_(() => {
    sheet_(SHEET.rows, ROW_HEAD).appendRow(rowArr_(row, me));
    log_(me, '新增影片', month, row.name);
  });
  return getState_(me, month, null);
}

function deleteRow_(me, id) {
  require_(me, ['admin', 'editor']);
  let month;
  withLock_(() => {
    const row = readRows_().filter(r => r.id === id)[0];
    if (!row) throw new Error('找不到這支影片，可能已被刪除。');
    month = row.month;
    sheet_(SHEET.rows, ROW_HEAD).deleteRow(row._row);
    log_(me, '刪除影片', month, row.name);
  });
  return getState_(me, month, null);
}

/** 依品牌配額重建「該月」明細（會取代該月現有明細） */
function rebuildRows_(me, month, cfg) {
  require_(me, ['admin']);
  month = cleanMonth_(month);
  cfg = cleanCfg_(cfg);
  withLock_(() => {
    writeCfg_(month, cfg, me);
    const sh = sheet_(SHEET.rows, ROW_HEAD);
    readRows_().filter(r => r.month === month).map(r => r._row)
      .sort((a, b) => b - a).forEach(n => sh.deleteRow(n));
    const fresh = [];
    BRANDS.forEach((b, i) => cfg.matrix[i].forEach((count, j) => {
      for (let k = 0; k < count; k++) {
        fresh.push(rowArr_({ id: newId_(), month: month, name: b + ' ' + GRADES[j] + '級影片 ' + (k + 1), brand: b,
          grade: GRADES[j], type: 'Reels', due: '', status: '待製作', done: false }, me));
      }
    }));
    if (fresh.length) sh.getRange(sh.getLastRow() + 1, 1, fresh.length, ROW_HEAD.length).setValues(fresh);
    log_(me, '由配額重建明細', month, fresh.length + ' 支');
  });
  return getState_(me, month, null);
}
