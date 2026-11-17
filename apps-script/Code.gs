/**
 * 三品牌影音製作 KPI｜線上同步版（Google Apps Script）
 *
 * 資料存在這份 Google 試算表的三個工作表：
 *   設定      ─ 每月一列：品牌配額、工時預算等
 *   明細      ─ 每支影片一列
 *   異動紀錄  ─ 誰（管理者／攝影師）在什麼時間改了什麼
 *
 * 權限（用網頁上的「編輯密碼」區分）：
 *   管理者   ─ 全部可改：配額、工時、新增／刪除影片、分級、品牌
 *   攝影師   ─ 明細全部可改：新增／刪除影片、名稱、品牌、分級、類型、交期、狀態
 *              （配額、工時預算與「由配額重建明細」仍只限管理者）
 *   沒登入   ─ 只能檢視（公司同事）
 * 密碼存在「指令碼屬性」，從試算表選單「KPI 儀表板 → 顯示編輯密碼」查看。
 */

const SHEET = { cfg: '設定', rows: '明細', log: '異動紀錄' };
const BRANDS = ['Gymray', 'RISEME', 'NEVERDIE'];
const GRADES = ['A', 'B', 'C'];
const ALL_GRADES = ['A', 'B', 'C', 'D'];
const TYPES = ['Reels', '季度主片', '其他'];
const STATUSES = ['待製作', '拍攝中', '剪輯中', '待審核', '修改中', '已完成'];
const ROLE_NAME = { admin: '管理者', editor: '攝影師', viewer: '檢視者' };
const BASE = { matrix: [[3, 3, 0], [1, 3, 1], [0, 2, 1]], hours: 40, available: 160, other: 32, edit: 3, dcount: 0, dpoints: 12 };
const NUM_KEYS = ['hours', 'available', 'other', 'edit', 'dcount', 'dpoints'];

const CFG_HEAD = ['月份']
  .concat(...BRANDS.map(b => GRADES.map(g => b + ' ' + g + '級')))
  .concat(['外拍時數', '正常工時預算', '其他預排工時', '每點後製工時', 'D級專案數', 'D級每案點數', '更新時間', '更新者']);
const ROW_HEAD = ['ID', '月份', '影片／專案名稱', '品牌', '分級', '類型', '交期', '狀態', '完成', '更新時間', '更新者'];
const LOG_HEAD = ['時間', '身分', '動作', '月份', '內容'];

/* ---------- 網頁入口 ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Jason 工作追蹤')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 試算表選單與初始化 ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('KPI 儀表板')
    .addItem('① 初始化工作表', 'setup')
    .addItem('顯示編輯密碼', 'showPins')
    .addItem('重新產生編輯密碼', 'resetPins')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ensureSheet_(ss, SHEET.cfg, CFG_HEAD);
  const rows = ensureSheet_(ss, SHEET.rows, ROW_HEAD);
  const log = ensureSheet_(ss, SHEET.log, LOG_HEAD);
  // 文字格式：避免 Google 試算表把「2026-11」「2026-11-20」自動轉成日期
  cfg.getRange('A:A').setNumberFormat('@');
  cfg.getRange(1, CFG_HEAD.length - 1, cfg.getMaxRows(), 1).setNumberFormat('@');
  ['A:A', 'B:B', 'G:G', 'J:J'].forEach(a => rows.getRange(a).setNumberFormat('@'));
  ['A:A', 'D:D'].forEach(a => log.getRange(a).setNumberFormat('@'));
  pins_();
  showPins();
}

function showPins() {
  const p = pins_();
  const msg = '管理者密碼（你自己用）：' + p.admin + '\n攝影師密碼：' + p.editor +
    '\n\n公司同事不需要密碼，打開網址即可檢視。';
  try { SpreadsheetApp.getUi().alert('KPI 儀表板編輯密碼', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(msg); }
}

function resetPins() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('ADMIN_PIN');
  props.deleteProperty('EDITOR_PIN');
  showPins();
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

/* ---------- 密碼與身分 ---------- */

function pins_() {
  const props = PropertiesService.getScriptProperties();
  let admin = props.getProperty('ADMIN_PIN');
  let editor = props.getProperty('EDITOR_PIN');
  if (!admin) { admin = randomPin_(); props.setProperty('ADMIN_PIN', admin); }
  if (!editor || editor === admin) { editor = randomPin_(); props.setProperty('EDITOR_PIN', editor); }
  return { admin: admin, editor: editor };
}

function randomPin_() {
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 8; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

function roleOf_(pin) {
  pin = String(pin || '').trim();
  if (!pin) return 'viewer';
  const p = pins_();
  if (pin === p.admin) return 'admin';
  if (pin === p.editor) return 'editor';
  return null; // 密碼錯誤
}

function require_(pin, roles) {
  const r = roleOf_(pin);
  if (roles.indexOf(r) < 0) throw new Error('沒有權限進行這個操作，請確認編輯密碼。');
  return r;
}

function login(pin) {
  const r = roleOf_(pin);
  if (!r || r === 'viewer') { Utilities.sleep(800); throw new Error('密碼錯誤'); }
  return r;
}

/* ---------- 版本號（讓網頁只在資料有變時才重新載入） ---------- */

function version_() {
  return Number(PropertiesService.getScriptProperties().getProperty('VERSION') || 0);
}
function bump_() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('VERSION', String(version_() + 1));
}

/* ---------- 讀取 ---------- */

function num_(v) { const n = Number(String(v).replace(/,/g, '')); return isFinite(n) && n > 0 ? n : 0; }

function readCfgMap_() {
  const sh = sheet_(SHEET.cfg, CFG_HEAD);
  const vals = sh.getDataRange().getDisplayValues().slice(1);
  const map = {};
  vals.forEach((r, i) => {
    const m = String(r[0]).trim();
    if (!/^\d{4}-\d{2}$/.test(m)) return;
    const matrix = BRANDS.map((b, bi) => GRADES.map((g, gi) => num_(r[1 + bi * 3 + gi])));
    const c = { matrix: matrix };
    NUM_KEYS.forEach((k, ki) => { c[k] = num_(r[10 + ki]); });
    map[m] = { row: i + 2, cfg: c };
  });
  return map;
}

function readRows_() {
  const sh = sheet_(SHEET.rows, ROW_HEAD);
  return sh.getDataRange().getDisplayValues().slice(1).map((r, i) => ({
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

function getState(month, pin, knownVersion) {
  month = cleanMonth_(month);
  let role = roleOf_(pin);
  const pinInvalid = role === null;
  if (pinInvalid) role = 'viewer';
  const version = version_();
  if (knownVersion !== null && knownVersion !== undefined && Number(knownVersion) === version) {
    return { same: true, version: version, role: role, pinInvalid: pinInvalid };
  }
  const map = readCfgMap_();
  const all = readRows_();
  const months = {};
  Object.keys(map).forEach(m => { months[m] = 1; });
  all.forEach(r => { if (/^\d{4}-\d{2}$/.test(r.month)) months[r.month] = 1; });
  const rows = all.filter(r => r.month === month).map(r => { const o = Object.assign({}, r); delete o._row; return o; });
  return {
    same: false, version: version, role: role, pinInvalid: pinInvalid, month: month,
    months: Object.keys(months).sort().reverse(),
    cfg: map[month] ? map[month].cfg : inheritCfg_(map, month),
    cfgSaved: !!map[month],
    rows: rows
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

function log_(role, action, month, detail) {
  sheet_(SHEET.log, LOG_HEAD).appendRow([stamp_(), ROLE_NAME[role] || role, action, month, String(detail).slice(0, 500)]);
}

function cleanMonth_(m) {
  m = String(m || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(m)) throw new Error('月份格式錯誤：' + m);
  return m;
}

function cleanCfg_(c) {
  c = c || {};
  const out = { matrix: BRANDS.map((b, i) => GRADES.map((g, j) => Math.min(100, Math.round(num_(((c.matrix || [])[i] || [])[j])) ))) };
  NUM_KEYS.forEach(k => { out[k] = Math.min(300, num_(c[k])); });
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

function rowArr_(o, role) {
  return [o.id, o.month, safeText_(o.name), o.brand, o.grade, o.type, o.due, o.status, o.done ? '是' : '否', stamp_(), ROLE_NAME[role]];
}

function newId_() { return 'v' + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

function writeCfg_(month, cfg, role) {
  const sh = sheet_(SHEET.cfg, CFG_HEAD);
  const map = readCfgMap_();
  const flat = [].concat(...cfg.matrix);
  const line = [month].concat(flat, NUM_KEYS.map(k => cfg[k]), [stamp_(), ROLE_NAME[role]]);
  if (map[month]) sh.getRange(map[month].row, 1, 1, line.length).setValues([line]);
  else sh.appendRow(line);
}

function saveConfig(month, cfg, pin) {
  const role = require_(pin, ['admin']);
  month = cleanMonth_(month);
  cfg = cleanCfg_(cfg);
  withLock_(() => {
    writeCfg_(month, cfg, role);
    log_(role, '更新設定', month, JSON.stringify(cfg));
  });
  return getState(month, pin, null);
}

function updateRow(id, patch, pin) {
  const role = require_(pin, ['admin', 'editor']);
  patch = patch || {};
  const clean = cleanRow_(patch);
  let month;
  withLock_(() => {
    const row = readRows_().filter(r => r.id === id)[0];
    if (!row) throw new Error('找不到這支影片，可能已被刪除，請重新整理。');
    month = row.month;
    Object.assign(row, clean);
    if ('status' in clean && !('done' in clean)) row.done = row.status === '已完成';
    if ('done' in clean && !('status' in clean)) row.status = row.done ? '已完成' : '剪輯中';
    sheet_(SHEET.rows, ROW_HEAD).getRange(row._row, 1, 1, ROW_HEAD.length).setValues([rowArr_(row, role)]);
    log_(role, '修改影片', month, row.name + '｜' + JSON.stringify(clean));
  });
  return getState(month, pin, null);
}

function addRow(month, data, pin) {
  const role = require_(pin, ['admin', 'editor']);
  month = cleanMonth_(month);
  const base = { name: '新增影片', brand: 'Gymray', grade: 'B', type: 'Reels', due: '', status: '待製作', done: false };
  const row = Object.assign(base, cleanRow_(data || {}), { id: newId_(), month: month });
  withLock_(() => {
    sheet_(SHEET.rows, ROW_HEAD).appendRow(rowArr_(row, role));
    log_(role, '新增影片', month, row.name);
  });
  return getState(month, pin, null);
}

function deleteRow(id, pin) {
  const role = require_(pin, ['admin', 'editor']);
  let month;
  withLock_(() => {
    const row = readRows_().filter(r => r.id === id)[0];
    if (!row) throw new Error('找不到這支影片，可能已被刪除。');
    month = row.month;
    sheet_(SHEET.rows, ROW_HEAD).deleteRow(row._row);
    log_(role, '刪除影片', month, row.name);
  });
  return getState(month, pin, null);
}

/** 依品牌配額重建「該月」明細（會取代該月現有明細） */
function rebuildRows(month, cfg, pin) {
  const role = require_(pin, ['admin']);
  month = cleanMonth_(month);
  cfg = cleanCfg_(cfg);
  withLock_(() => {
    writeCfg_(month, cfg, role);
    const sh = sheet_(SHEET.rows, ROW_HEAD);
    readRows_().filter(r => r.month === month).map(r => r._row)
      .sort((a, b) => b - a).forEach(n => sh.deleteRow(n));
    const fresh = [];
    BRANDS.forEach((b, i) => cfg.matrix[i].forEach((count, j) => {
      for (let k = 0; k < count; k++) {
        fresh.push(rowArr_({ id: newId_(), month: month, name: b + ' ' + GRADES[j] + '級影片 ' + (k + 1), brand: b,
          grade: GRADES[j], type: 'Reels', due: '', status: '待製作', done: false }, role));
      }
    }));
    if (fresh.length) sh.getRange(sh.getLastRow() + 1, 1, fresh.length, ROW_HEAD.length).setValues(fresh);
    log_(role, '由配額重建明細', month, fresh.length + ' 支');
  });
  return getState(month, pin, null);
}
