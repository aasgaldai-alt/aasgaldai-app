// @ts-nocheck
// Аасгалдай API — Google Apps Script логикийг (Code.gs) Supabase дээр ажиллуулах орчин.
// Энэ файлыг build_api.py үүсгэдэг: Code.gs-ийг доорх __CODE__ хэсэгт хуулна. Гараар бүү засаарай.
import { createHash, randomUUID } from "node:crypto";
import { Buffer } from "node:buffer";
import { vapidGenerate, sendPush } from "./webpush.js";

export const KEY_HASH = "f5e809095950b7176a2a97a76a6dcb3130c015274c6a1e585d736e8d3bc2b01d";
const TZ = "Asia/Ulaanbaatar";
const APP_URL = "https://app.aasgaldai.mn/";
const FILE_MARK = "\u0001FILE:";
const WORKER = Math.random().toString(36).slice(2, 7);

/* ---------- огноо, кодлол ---------- */
const fmts = {};
function partsOf(d, tz) {
  const f = fmts[tz] || (fmts[tz] = new Intl.DateTimeFormat("en-GB", { timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23" }));
  const o = {};
  for (const p of f.formatToParts(d)) o[p.type] = p.value;
  return o;
}
function formatDate(d, tz, pat) {
  if (!(d instanceof Date) || isNaN(d)) throw new Error("Invalid argument: date");
  const o = partsOf(d, tz || TZ);
  return String(pat).replace(/yyyy|yy|MM|dd|HH|mm|ss/g, (t) => ({ yyyy: o.year, yy: o.year.slice(2), MM: o.month, dd: o.day, HH: o.hour, mm: o.minute, ss: o.second })[t]);
}
export function encCell(v) { return v instanceof Date ? (isNaN(v) ? "" : { $d: v.getTime() }) : (v === undefined || v === null ? "" : v); }
export function decCell(v) { return v && typeof v === "object" && "$d" in v ? new Date(v.$d) : v; }
export function encRows(rows) {
  const out = rows.map((r) => {
    const a = (r || []).map(encCell);
    let n = a.length;
    while (n > 0 && a[n - 1] === "") n--;
    a.length = n;
    return a;
  });
  let m = out.length;
  while (m > 0 && out[m - 1].length === 0) m--;
  out.length = m;
  return out;
}
export function decRows(rows) { return (rows || []).map((r) => (r || []).map(decCell)); }
// Apps Script-ийн Utilities.computeDigest(alg, string) нь латин бус үсгийг ямар кодлолоор хувиргадгийг шилжүүлэх үед тодорхойлно
export const ENCODERS = {
  utf8: (s) => Buffer.from(s, "utf8"),
  latin1q: (s) => Buffer.from(Array.from(s, (ch) => { const c = ch.charCodeAt(0); return c > 255 ? 63 : c; })),
  low8: (s) => Buffer.from(s, "latin1"),
  utf16le: (s) => Buffer.from(s, "utf16le"),
  utf16be: (s) => { const b = Buffer.from(s, "utf16le"); for (let i = 0; i + 1 < b.length; i += 2) { const t = b[i]; b[i] = b[i + 1]; b[i + 1] = t; } return b; },
  utf16bom: (s) => Buffer.concat([Buffer.from([0xfe, 0xff]), ENCODERS.utf16be(s)]),
};
export function digestBytes(alg, s, enc) { return Array.from(createHash(alg).update((ENCODERS[enc] || ENCODERS.utf8)(String(s))).digest()).map((b) => (b > 127 ? b - 256 : b)); }
export function sha256hex(s) { return createHash("sha256").update(String(s), "utf8").digest("hex"); }

// Google Sheets утгыг хадгалахдаа хийдэг хувиргалт: '-ээр эхэлсэн бол текст, цэвэр тоо бол тоо
function norm(v) {
  if (v === undefined || v === null) return "";
  if (typeof v === "string") {
    if (v.charAt(0) === "'") return v.slice(1);
    const t = v.trim();
    if (t.length > 0 && t.length < 16 && /^[-+]?\d+(\.\d+)?$/.test(t)) return Number(t);
    return v;
  }
  if (typeof v === "number") return isFinite(v) ? v : "";
  if (v instanceof Date) return isNaN(v) ? "" : new Date(v.getTime());
  return v;
}
const isEmpty = (x) => x === "" || x === null || x === undefined;

/* ---------- SpreadsheetApp ---------- */
class Range {
  constructor(sh, r, c, nr, nc) {
    if (!(r >= 1) || !(c >= 1)) throw new Error("Range-ийн мөр, багана 1-ээс эхэлнэ");
    if (!(nr >= 1)) throw new Error("The number of rows in the range must be at least 1.");
    if (!(nc >= 1)) throw new Error("The number of columns in the range must be at least 1.");
    this.sh = sh; this.r = r; this.c = c; this.nr = nr; this.nc = nc;
  }
  getRow() { return this.r; }
  getColumn() { return this.c; }
  getNumRows() { return this.nr; }
  getNumColumns() { return this.nc; }
  getValues() {
    const rows = this.sh.rows, out = [];
    for (let i = 0; i < this.nr; i++) {
      const src = rows[this.r - 1 + i] || [], row = [];
      for (let j = 0; j < this.nc; j++) { const v = src[this.c - 1 + j]; row.push(isEmpty(v) ? "" : v); }
      out.push(row);
    }
    return out;
  }
  getDisplayValues() { return this.getValues().map((r) => r.map((v) => (v instanceof Date ? formatDate(v, TZ, "yyyy-MM-dd") : String(v)))); }
  getValue() { return this.getValues()[0][0]; }
  _put(i, j, v) {
    const rows = this.sh.rows, ri = this.r - 1 + i, ci = this.c - 1 + j;
    while (rows.length <= ri) rows.push([]);
    const row = rows[ri] || (rows[ri] = []);
    while (row.length < ci) row.push("");
    row[ci] = v;
  }
  setValues(vals) {
    if (!Array.isArray(vals) || vals.length !== this.nr) throw new Error("The number of rows in the data does not match the number of rows in the range. The data has " + (vals && vals.length) + " but the range has " + this.nr + ".");
    this.sh._w();
    for (let i = 0; i < this.nr; i++) {
      if (!Array.isArray(vals[i]) || vals[i].length !== this.nc) throw new Error("The number of columns in the data does not match the number of columns in the range. The data has " + (vals[i] && vals[i].length) + " but the range has " + this.nc + ".");
      for (let j = 0; j < this.nc; j++) this._put(i, j, norm(vals[i][j]));
    }
    return this;
  }
  setValue(v) {
    this.sh._w();
    const x = norm(v);
    for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this._put(i, j, x instanceof Date ? new Date(x.getTime()) : x);
    return this;
  }
  clearContent() { this.sh._w(); for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this._put(i, j, ""); return this; }
  deleteCells(dim) {
    this.sh._w();
    const rows = this.sh.rows;
    if (dim === "ROWS") {
      for (let j = this.c - 1; j < this.c - 1 + this.nc; j++) {
        for (let i = this.r - 1; i < rows.length; i++) {
          const below = rows[i + this.nr];
          const v = below && !isEmpty(below[j]) ? below[j] : "";
          const row = rows[i] || (rows[i] = []);
          while (row.length <= j) row.push("");
          row[j] = v;
        }
      }
    } else {
      for (let i = this.r - 1; i < this.r - 1 + this.nr; i++) { const row = rows[i]; if (row) row.splice(this.c - 1, this.nc); }
    }
    return this;
  }
  setFontWeight() { return this; }
  setNumberFormat() { return this; }
  setBackground() { return this; }
  setWrap() { return this; }
}
class Sheet {
  constructor(book, name, entry) { this.book = book; this.name = name; this.entry = entry; this.rows = entry ? entry.rows : []; this.own = !entry; }
  _w() { if (!this.own) { this.rows = this.rows.map((r) => (r ? r.slice() : [])); this.own = true; } this.book.dirty.add(this); }
  getName() { return this.name; }
  getLastRow() {
    for (let i = this.rows.length - 1; i >= 0; i--) { const r = this.rows[i]; if (r && r.some((x) => !isEmpty(x))) return i + 1; }
    return 0;
  }
  getLastColumn() {
    let m = 0;
    for (const r of this.rows) { if (!r) continue; for (let j = r.length - 1; j >= m; j--) if (!isEmpty(r[j])) { m = j + 1; break; } }
    return m;
  }
  getMaxColumns() { return Math.max(26, this.getLastColumn()); }
  getMaxRows() { return Math.max(1000, this.rows.length); }
  insertColumnsAfter() { return this; }
  insertColumns() { return this; }
  setFrozenRows() { return this; }
  setColumnWidth() { return this; }
  autoResizeColumns() { return this; }
  getRange(r, c, nr, nc) {
    if (typeof r === "string") throw new Error("A1 хаягаар Range авах дэмжигдээгүй");
    return new Range(this, r, c, nr === undefined ? 1 : nr, nc === undefined ? 1 : nc);
  }
  getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  appendRow(row) { this._w(); this.rows.length = this.getLastRow(); this.rows.push(row.map(norm)); return this; }
  deleteRow(i) { this._w(); if (i >= 1 && i <= this.rows.length) this.rows.splice(i - 1, 1); return this; }
  deleteRows(i, n) { this._w(); this.rows.splice(i - 1, n); return this; }
  clear() { this._w(); this.rows = []; return this; }
}
class Book {
  constructor(snap) { this.snap = snap; this.sheets = new Map(); this.dirty = new Set(); }
  getSheetByName(n) {
    if (!this.sheets.has(n)) { const e = this.snap[n]; if (!e) return null; this.sheets.set(n, new Sheet(this, n, e)); }
    return this.sheets.get(n);
  }
  insertSheet(n) {
    if (this.getSheetByName(n)) throw new Error('A sheet with the name "' + n + '" already exists.');
    const s = new Sheet(this, n, null); this.sheets.set(n, s); this.dirty.add(s); return s;
  }
  sys(n) { return this.getSheetByName(n) || this.insertSheet(n); }
  getSheets() { return Object.keys(this.snap).filter((n) => n.indexOf("__") !== 0).map((n) => this.getSheetByName(n)); }
  getId() { return "supabase"; }
  getUrl() { return APP_URL; }
}

/* ---------- бусад Apps Script үйлчилгээ ---------- */
const isoCache = new Map(); // нэвтрэх оролдлогын тоолуур (instance дотор)
function makeCtx(snap) {
  const book = new Book(snap), reqCache = new Map(), ctx = { book, uploads: [] };
  const props = {
    _sh() { return book.getSheetByName("__props"); },
    getProperty(k) { const s = this._sh(); if (!s) return null; const r = s.rows.find((x) => x && x[0] === k); return r ? String(r[1]) : null; },
    getProperties() { const s = this._sh(), o = {}; if (s) s.rows.forEach((x) => { if (x && x[0]) o[x[0]] = String(x[1]); }); return o; },
    setProperty(k, v) {
      const s = book.sys("__props"); s._w();
      const r = s.rows.find((x) => x && x[0] === k);
      if (r) r[1] = String(v); else s.rows.push([k, String(v)]);
      return this;
    },
    deleteProperty(k) { const s = this._sh(); if (!s) return this; s._w(); s.rows = s.rows.filter((x) => !(x && x[0] === k)); return this; },
  };
  ctx.SpreadsheetApp = { getActiveSpreadsheet: () => book, getActive: () => book, Dimension: { ROWS: "ROWS", COLUMNS: "COLUMNS" }, flush() {}, openById() { throw new Error("Энэ үйлдэл шинэ сервер дээр ажиллахгүй"); } };
  ctx.PropertiesService = { getScriptProperties: () => props };
  ctx.CacheService = {
    getScriptCache: () => ({
      get(k) {
        if (k.indexOf("fail_") === 0) { const h = isoCache.get(k); if (!h || h.exp < Date.now()) { isoCache.delete(k); return null; } return h.v; }
        return reqCache.has(k) ? reqCache.get(k) : null;
      },
      put(k, v, ttl) { if (k.indexOf("fail_") === 0) isoCache.set(k, { v: String(v), exp: Date.now() + (ttl || 600) * 1000 }); else reqCache.set(k, String(v)); },
      remove(k) { isoCache.delete(k); reqCache.delete(k); },
    }),
  };
  ctx.LockService = { getScriptLock: () => ({ waitLock() {}, tryLock() { return true; }, releaseLock() {}, hasLock() { return true; } }) };
  ctx.Utilities = {
    DigestAlgorithm: { SHA_256: "sha256", SHA_1: "sha1", MD5: "md5" },
    computeDigest(alg, s) { return digestBytes(alg, s, props.getProperty("__DIGEST_ENC") || "utf8"); },
    formatDate,
    getUuid: () => randomUUID(),
    base64Encode: (b) => Buffer.from(typeof b === "string" ? Buffer.from(b, "utf8") : Uint8Array.from(b, (x) => x & 255)).toString("base64"),
    base64Decode: (s) => Array.from(Buffer.from(String(s), "base64")).map((b) => (b > 127 ? b - 256 : b)),
    sleep() {},
  };
  ctx.MailApp = {
    sendEmail(to, subject, body) {
      if (typeof to === "object" && to) { subject = to.subject; body = to.body || to.htmlBody; to = to.to; }
      const s = book.sys("__mail"); s._w(); s.rows.push([String(to), String(subject || ""), String(body || ""), Date.now()]);
    },
  };
  ctx.ScriptApp = { getService: () => ({ getUrl: () => APP_URL }) };
  ctx.DriveApp = { getFileById() { throw new Error("DriveApp дэмжигдээгүй"); } };
  ctx.Logger = { log: (...a) => console.log(...a) };
  return ctx;
}

/* ---------- Code.gs логик ---------- */
function buildApp(ctx) {
  const SpreadsheetApp = ctx.SpreadsheetApp, PropertiesService = ctx.PropertiesService, CacheService = ctx.CacheService, LockService = ctx.LockService,
    Utilities = ctx.Utilities, MailApp = ctx.MailApp, ScriptApp = ctx.ScriptApp, DriveApp = ctx.DriveApp, Logger = ctx.Logger,
    HtmlService = undefined, ContentService = undefined;
  /**
 * Аасгалдай гүйлгээний хүсэлт — Google Apps Script backend.
 * "Аасгалдай гүйлгээний хүсэлт" Google Sheet-ийн Extensions > Apps Script дотор байрлана.
 * Хуудсууд:
 *   Хүсэлт    — хүсэлт бүр нэг мөр
 *   Төсөл     — Код | Нэр | Идэвхтэй | Гэрээний дүн | 5 суутгал | Чанарын баталгаа % | Гэрээний зураг | (томьёо L:N)
 *   Санхүүжилт — ID | Огноо | Төсөл | Тайлбар | Дүн | НӨАТ | Баримт | Бүртгэсэн | Тэмдэглэл
 *   Ангилал, Банк
 *   Хэрэглэгч — Нэр | Үүрэг (Ажилчин/Нягтлан/Админ/Захирал) | PIN (нууцалсан) | Имэйл | Утас | PIN солих
 */
var ST = { WAIT: 'Хүлээгдэж буй', OK: 'Зөвшөөрсөн', NO: 'Буцаасан', PAID: 'Төлсөн' };
var ROLES = ['Ажилчин', 'Оператор', 'Механик', 'Нярав', 'Мастер', 'Чанарын инженер', 'Инженер', 'Тогооч', 'ХАБ инженер', 'Төслийн удирдагч', 'Хүний нөөц', 'Нягтлан', 'Админ', 'Захирал'];
var PENDING = 'Хүлээгдэж буй';
var VATS = ['НӨАТ-тай', 'НӨАТ-гүй'];
var GEN = 'ЕРӨНХИЙ', GEN_NAME = 'Компанийн ерөнхий (төсөлд хамаарахгүй)';
var RT = { LOAN: 'Зээл олгох', ADV: 'Урьдчилгаа цалин', PRIN: 'Зээлийн үндсэн төлбөр', INT: 'Зээлийн хүү, шимтгэл' };
var RT_ALL = [RT.LOAN, RT.ADV, RT.PRIN, RT.INT];
var RT_NONEXP = [RT.LOAN, RT.ADV, RT.PRIN]; // зардалд тооцогдохгүй (авлага/өр)
function loanLimit_() { return num_(PropertiesService.getScriptProperties().getProperty('LOAN_LIMIT')) || 5000000; }
function isDir_(u) { return u.role === 'Захирал'; }
function isAdmin_(u) { return u.role === 'Админ'; }
function isMaster_(u) { return u.role === 'Мастер'; }
function isMech_(u) { return u.role === 'Механик'; }
function isStore_(u) { return u.role === 'Нярав'; }
function isOper_(u) { return u.role === 'Оператор'; }
function isEng_(u) { return u.role === 'Инженер'; }
function canUsers_(u) { return isBoss_(u) || isHR_(u) || isAdmin_(u); }
function fuelView_(u) { return isBoss_(u) || isMech_(u) || isStore_(u) || isOper_(u) || isPM_(u); }
function fuelAdd_(u) { return isBoss_(u) || isMech_(u) || isStore_(u) || isOper_(u); }
function approveLimit_() { return num_(PropertiesService.getScriptProperties().getProperty('APPROVE_LIMIT')) || 10000000; }
// Батлах эрх: 10 сая хүртэл — тухайн төслийн удирдагч эсвэл захирал; түүнээс дээш — зөвхөн захирал. Нягтлан батлахгүй, зөвхөн төлнө/буцаана.
function canApprove_(u, r) {
  if (r.by === u.name && !isDir_(u)) return false;
  if (isDir_(u)) return true;
  if (r.proj === GEN) return false; // компанийн ерөнхий зардлыг зөвхөн захирал
  var lim = RT_ALL.indexOf(r.rtype) >= 0 ? loanLimit_() : approveLimit_();
  return isPM_(u) && canProj_(u, r.proj) && num_(r.amount) <= lim;
}
function pmSees_(u, r) { return isPM_(u) && r.proj !== GEN && canProj_(u, r.proj); }
function isCook_(u) { return u.role === 'Тогооч'; }
function isPM_(u) { return u.role === 'Төслийн удирдагч'; }
function isHR_(u) { return u.role === 'Хүний нөөц'; }
function isHab_(u) { return u.role === 'ХАБ инженер'; }
function isBoss_(u) { return u.role === 'Нягтлан' || isDir_(u); }
var COLS = ['id','date','by','proj','cat','what','qty','payee','bank','acct','amount','photo','status','decided','paid','desc','note','vat','ebarimt','phone','loc','rtype','link'];
var PCOLS = ['code','name','active','contract','c1','c2','c3','c4','c5','pct','photo']; // Төсөл A:K
var ICOLS = ['id','date','proj','desc','amount','vat','photo','by','note']; // Санхүүжилт A:I
var PHOTO_FOLDER = 'Аасгалдай баримт';
var INCOME_FOLDER = 'Аасгалдай санхүүжилт';
var CONTRACT_FOLDER = 'Аасгалдай гэрээ';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Аасгалдай хүсэлт')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doPost(e) {
  var out = handle_(JSON.parse(e.postData.contents));
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function api(body) { return JSON.stringify(handle_(JSON.parse(body))); }

function handle_(b) {
  try {
    if (b.action === 'login') return login_(b.login, b.pin);
    if (b.action === 'register') return register_(b.r);
    var me = auth_(b);
    if (!me) return { error: 'auth', message: 'Дахин нэвтэрнэ үү' };
    if (b.action === 'changePin') return changePin_(me, b.oldPin, b.newPin);
    if (me.mustChange) return { error: 'mustchange', message: 'Эхлээд шинэ PIN тохируулна уу' };
    var boss = isBoss_(me);
    var director = isDir_(me);
    switch (b.action) {
      case 'meta': return meta_();
      case 'contacts': return { contacts: contacts_() };
      case 'people':
        if (!boss && !isHR_(me) && !isPM_(me) && !isMaster_(me)) return deny_();
        return peopleData_(me);
      case 'markP':
        if (!director && !isHR_(me) && !isPM_(me) && !isMaster_(me)) return deny_();
        markP_(me, b.a); return peopleData_(me);
      case 'my': return myData_(me);
      case 'meals':
        if (!isCook_(me) && !boss && !isHR_(me) && !isPM_(me) && !isMaster_(me)) return deny_();
        return mealData_(me);
      case 'mealAdd':
        if (!isCook_(me)) return deny_();
        mealAdd_(me, b.meal, b.photo); return mealData_(me);
      case 'mealDelete':
        if (!isCook_(me) && !director) return deny_();
        mealDelete_(me, b.id); return mealData_(me);
      case 'mealPhoto':
        var ml = meals_().filter(function (x) { return x.id === b.id; })[0];
        return { data: ml && ml.photo ? fileData_(ml.photo) : '' };
      case 'leaveReq': leaveReq_(me, b.lv); return myData_(me);
      case 'leaveDecide':
        if (!boss && !isPM_(me)) return deny_();
        leaveDecide_(me, b.id, b.ok, b.note); return peopleData_(me);
      case 'taskSend':
        if (!director && !isPM_(me) && !isMaster_(me)) return deny_();
        taskSend_(me, b.t); return peopleData_(me);
      case 'taskDelete':
        if (!director && !isPM_(me) && !isMaster_(me)) return deny_();
        taskDelete_(me, b.id); return peopleData_(me);
      case 'taskAck': taskUpd_(me, b.id, 'ack'); return myData_(me);
      case 'taskReport': taskUpd_(me, b.id, b.done ? 'done' : 'notdone', b.note); return myData_(me);
      case 'saveWage':
        if (!boss) return deny_();
        saveWage_(b.w); return peopleData_(me);
      case 'payCfg':
        if (!boss) return deny_();
        PropertiesService.getScriptProperties().setProperty('PAYCFG', JSON.stringify(b.cfg || {})); return peopleData_(me);
      case 'hab': return habData_(me);
      case 'habSave':
        if (!isHab_(me)) return deny_();
        habSave_(me, b.note); return habData_(me);
      case 'habDelete':
        if (!isHab_(me)) return deny_();
        habDelete_(b.id); return habData_(me);
      case 'habRead': habRead_(me, b.id); return habData_(me);
      case 'habReport': habReport_(me, b.rep, b.photo); return habData_(me);
      case 'habStatus':
        if (!isHab_(me) && !director) return deny_();
        habStatus_(me, b.id, b.status, b.act); return habData_(me);
      case 'habPhoto':
        var hr = habReports_().filter(function (x) { return x.id === b.id; })[0];
        if (!hr || !hr.photo || (!isHab_(me) && !director && hr.by !== me.name)) return { data: '' };
        return { data: fileData_(hr.photo) };
      case 'setLocReq':
        if (!director) return deny_();
        PropertiesService.getScriptProperties().setProperty('LOC_REQUIRED', b.on ? 'true' : 'false');
        CacheService.getScriptCache().remove('meta'); return meta_();
      case 'boot':
        var allRows = rows_();
        var mine = boss ? allRows : allRows.filter(function (r) { return r.by === me.name || pmSees_(me, r); });
        var out = { reqs: mine.slice(-500), meta: meta_() };
        if (boss) out.projects = dashboard_(allRows);
        try { out.hab = habBadge_(me); } catch (e) {}
        out.myProjects = me.projects;
        out.canReq = me.canReq !== false;
        if (boss || isPM_(me)) try { out.loansOpen = loanList_().filter(function (l) { return l.type === 'Авсан' && l.bal > 0; }).map(function (l) { return { id: l.id, party: l.party, proj: l.proj, bal: l.bal }; }); out.emps = activeUsers_().map(function (u) { return u.name; }); } catch (e) {}
        try { out.hr = hrBadge_(me); } catch (e) {}
        try { out.meal = lastMeal_(me); } catch (e) {}
        return out;
      case 'list':
        var all = rows_();
        if (!boss) all = all.filter(function (r) { return r.by === me.name || pmSees_(me, r); }); // ажилчин зөвхөн өөрийнхийг, ТУ өөрийн төслийнхийг
        return { reqs: all.slice(-500) };
      case 'create': return { req: create_(me, b.req, b.photo) };
      case 'status':
        if (!boss && !isPM_(me)) return deny_();
        return { req: setStatus_(me, b.id, b.status, b.note) };
      case 'photo':
        var r = rows_().filter(function (x) { return x.id === b.id; })[0];
        if (!r || !r.photo || (!boss && r.by !== me.name && !pmSees_(me, r))) return { data: '' };
        return { data: fileData_(r.photo) };
      case 'dashboard':
        if (!boss) return deny_();
        return { projects: dashboard_() };
      case 'contract':
        if (!boss) return deny_();
        var p = projects_().filter(function (x) { return x.code === b.code; })[0];
        var cids = p && p.photo ? String(p.photo).split(',').filter(String) : [];
        var pg = Math.max(0, Math.min(cids.length - 1, +b.page || 0));
        return { data: cids.length ? fileData_(cids[pg]) : '', page: pg, pages: cids.length };
      case 'saveProject':
        if (!director) return deny_();
        return { project: saveProject_(b.project, b.photo, b.photos) };
      case 'addContractPages':
        if (!director) return deny_();
        return { pages: addContractPages_(b.code, b.photos) };
      case 'deleteProject':
        if (!director) return deny_();
        deleteProject_(b.code); return { ok: true };
      case 'addIncome':
        if (!boss) return deny_();
        return { income: addIncome_(me, b.inc, b.photo) };
      case 'deleteIncome':
        if (!director) return deny_();
        deleteIncome_(b.id); return { ok: true };
      case 'incomePhoto':
        if (!boss) return deny_();
        var inc = incomesRaw_().filter(function (x) { return x.id === b.id; })[0];
        return { data: inc && inc.photo ? fileData_(inc.photo) : '' };
      case 'userList':
        if (!canUsers_(me)) return deny_();
        return { users: pubUsers_(boss, true) };
      case 'saveUser':
        if (!canUsers_(me)) return deny_();
        saveUser_(me, b.user, b.isNew); return { users: pubUsers_(boss, true) };
      case 'deleteUser':
        if (!canUsers_(me)) return deny_();
        deleteUser_(me, b.target); return { users: pubUsers_(boss, true) };
      case 'loans':
        if (!boss) return deny_();
        return loanData_();
      case 'loanSave':
        if (!boss) return deny_();
        loanSave_(me, b.l); return loanData_();
      case 'loanTx':
        if (!boss) return deny_();
        loanTx_(me, b.tx); return loanData_();
      case 'loanDel':
        if (!director) return deny_();
        loanDel_(b.id, b.tx); return loanData_();
      case 'payDeduct':
        if (!boss) return deny_();
        payDeduct_(me, b.ym, b.items); return loanData_();
      case 'setLoanLimit':
        if (!director) return deny_();
        PropertiesService.getScriptProperties().setProperty('LOAN_LIMIT', String(Math.round(num_(b.limit)) || 5000000));
        CacheService.getScriptCache().remove('meta'); return meta_();
      case 'sched': return schedData_(me);
      case 'schedItem': schedItem_(me, b.it); return schedData_(me);
      case 'schedItemDel': schedItemDel_(me, b.id); return schedData_(me);
      case 'schedDay': schedDay_(me, b.proj, b.date, b.entries); return schedData_(me);
      case 'setLimit':
        if (!director) return deny_();
        PropertiesService.getScriptProperties().setProperty('APPROVE_LIMIT', String(Math.round(num_(b.limit)) || 10000000));
        CacheService.getScriptCache().remove('meta'); return meta_();
      case 'letters':
        if (!isBoss_(me)) return deny_();
        return letterData_(me);
      case 'letterSave':
        if (!isBoss_(me)) return deny_();
        return { letter: letterSave_(me, b.l), data: letterData_(me) };
      case 'letterVoid':
        if (!isBoss_(me)) return deny_();
        letterVoid_(me, b.no); return letterData_(me);
      case 'letterStamp':
        if (!isBoss_(me)) return deny_();
        var stId = PropertiesService.getScriptProperties().getProperty('LETTER_STAMP');
        return { data: stId ? fileData_(stId) : '' };
      case 'letterCfg':
        if (!isDir_(me) && me.role !== 'Нягтлан') return deny_();
        letterCfgSave_(b.cfg || {}, b.photo); return letterData_(me);
      case 'fuel': case 'equipPhoto': case 'fuelPhoto':
        if (!fuelView_(me)) return deny_();
        break;
      case 'addFuel': case 'mark':
        if (!fuelAdd_(me)) return deny_();
        break;
      case 'saveEquip': case 'deleteFuel':
        if (!boss && !isMech_(me)) return deny_();
        break;
      case 'deleteEquip':
        if (!director) return deny_();
        break;
    }
    switch (b.action) {
      case 'fuel': return fuelData_(me, boss);
      case 'saveEquip':
        saveEquip_(b.eq, b.photo); return fuelData_(me, boss);
      case 'deleteEquip':
        if (!director) return deny_();
        deleteEquip_(b.id); return fuelData_(me, boss);
      case 'equipPhoto':
        var eq = equips_().filter(function (x) { return x.id === b.id; })[0];
        return { data: eq && eq.photo ? fileData_(eq.photo) : '' };
      case 'addFuel': addFuel_(me, b.f, b.photo); return fuelData_(me, boss);
      case 'mark': mark_(me, boss, b.a); return fuelData_(me, boss);
      case 'deleteFuel':
        deleteFuel_(b.id); return fuelData_(me, boss);
      case 'fuelPhoto':
        var fr = fuelsRaw_().filter(function (x) { return x.id === b.id; })[0];
        if (!fr || !fr.photo || (!boss && !isMech_(me) && !isStore_(me) && !isPM_(me) && fr.by !== me.name)) return { data: '' };
        return { data: fileData_(fr.photo) };
    }
    return { error: 'bad', message: 'Unknown action' };
  } catch (err) {
    return { error: 'server', message: String(err && err.message || err) };
  }
}
function deny_() { return { error: 'bad', message: 'Эрх хүрэхгүй' }; }

/* ---------- хэрэглэгч, нэвтрэлт ----------
 * Хэрэглэгч хуудсыг гараар засвал 5 минутын дотор апп-д тусна (кэш).
 * Хэрэглэгч хуудас: A Нэр | B Үүрэг | C PIN (нууцалсан) | D Имэйл | E Утас | F PIN солих (TRUE = анх нэвтрэхэд солих) | G Төслүүд (* эсвэл хоосон = бүх төсөл, эсвэл ТТ2,ГАЗ)
 *                   H Албан тушаал | I Цалингийн төрөл (Өдөр/Сар) | J Цалин ₮ | K Сарын ажлын өдөр
 * PIN нь "h1:" + SHA-256 хэлбэрээр хадгалагдана. Хуучин ил PIN-тэй мөр анх нэвтрэхэд автоматаар нууцлагдаж, солих шаардлагатай болно.
 */
var MAX_FAIL = 5, LOCK_SEC = 15 * 60;
function users_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('users');
  if (hit) return JSON.parse(hit);
  var sh = sh_('Хэрэглэгч'); if (sh.getLastRow() < 2) return [];
  if (sh.getMaxColumns() < 19) sh.insertColumnsAfter(sh.getMaxColumns(), 19 - sh.getMaxColumns());
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 19).getValues();
  var out = [];
  v.forEach(function (r, i) {
    if (!r[0]) return;
    out.push({ row: i + 2, name: String(r[0]).trim(), role: String(r[1]).trim(), pin: String(r[2]).trim(), email: String(r[3] || '').trim(),
               phone: digits_(r[4]), mustChange: r[5] === true || String(r[5]).toUpperCase() === 'TRUE', projects: projList_(r[6]),
               position: String(r[7] || '').trim(), wageType: String(r[8]) === 'Сар' ? 'Сар' : 'Өдөр', wage: num_(r[9]), wdays: num_(r[10]) || 22,
               surname: String(r[11] || '').trim(), given: String(r[12] || '').trim(), age: num_(r[13]) || '', sex: String(r[14] || ''),
               canReq: !(r[15] === false || String(r[15]).toUpperCase() === 'FALSE'), status: String(r[16] || '').trim(), regAt: fmtDate_(r[17]), sched: r[18] === true || String(r[18]).toUpperCase() === 'TRUE' });
  });
  cache.put('users', JSON.stringify(out), 300);
  return out;
}
function projList_(x) { var p = String(x || '').trim(); if (!p || p === '*') return null; return p.split(/[,\s]+/).filter(function (c) { return c; }); }
function canProj_(u, code) { return !u.projects || u.projects.indexOf(code) >= 0; }
function dropUsers_() { CacheService.getScriptCache().remove('users'); }
function pubUsers_(full, personal) { return users_().map(function (u) { var o = { name: u.name, role: u.role, email: u.email, phone: u.phone, mustChange: u.mustChange, projects: u.projects, position: u.position, canReq: u.canReq, status: u.status, sched: u.sched };
  if (personal) { o.surname = u.surname; o.given = u.given; o.age = u.age; o.sex = u.sex; o.regAt = u.regAt; }
  if (full) { o.wageType = u.wageType; o.wage = u.wage; o.wdays = u.wdays; } return o; }); }
function activeUsers_() { return users_().filter(function (u) { return u.status !== PENDING; }); }
function contacts_() { return activeUsers_().filter(function (u) { return u.phone; }).map(function (u) { return { name: u.name, role: u.role, position: u.position, phone: u.phone }; }); }
var UHEAD2 = ['Овог', 'Өөрийн нэр', 'Нас', 'Хүйс', 'Хүсэлтийн эрх', 'Төлөв', 'Бүртгүүлсэн'];
function userHead_(sh) { if (!String(sh.getRange(1, 12).getValue())) sh.getRange(1, 12, 1, 7).setValues([UHEAD2]); if (!String(sh.getRange(1, 19).getValue())) sh.getRange(1, 19).setValue('График хөтлөгч'); }
function register_(r) {
  r = r || {};
  var sur = String(r.surname || '').trim(), giv = String(r.given || '').trim(), age = Math.round(num_(r.age)), sex = String(r.sex || '').trim(), phone = digits_(r.phone), pin = String(r.pin || '').trim();
  if (!sur || !giv) throw new Error('Овог, нэрээ бичнэ үү');
  if (!(age >= 16 && age <= 85)) throw new Error('Насаа зөв бичнэ үү');
  if (['Эрэгтэй', 'Эмэгтэй'].indexOf(sex) < 0) throw new Error('Хүйсээ сонгоно уу');
  if (phone.length !== 8) throw new Error('Утасны дугаар 8 оронтой байна');
  checkPin_('Ажилчин', pin);
  var all = users_();
  if (all.some(function (x) { return x.phone === phone; })) throw new Error('Энэ утасны дугаар бүртгэлтэй байна. Нэвтрэх хэсгээс орно уу.');
  if (all.filter(function (x) { return x.status === PENDING; }).length >= 50) throw new Error('Бүртгэлийн хүсэлт олон хүлээгдэж байна. Хүний нөөцөд хандана уу.');
  var base = sur.charAt(0).toUpperCase() + '.' + giv.charAt(0).toUpperCase() + giv.slice(1), name = base, k = 2;
  while (all.some(function (x) { return x.name === name; })) name = base + ' ' + (k++);
  var sh = sh_('Хэрэглэгч'); userHead_(sh);
  sh.appendRow([name, 'Ажилчин', hashPin_(name, pin), '', "'" + phone, false, '', '', 'Өдөр', '', '', sur, giv, age, sex, true, PENDING, new Date()]);
  dropUsers_();
  var to = users_().filter(function (u) { return (isHR_(u) || isDir_(u) || isAdmin_(u)) && u.email; }).map(function (u) { return u.email; });
  if (to.length) try { MailApp.sendEmail(to.join(','), 'Шинэ бүртгэл: ' + name, sur + ' ' + giv + ' (' + age + ', ' + sex + ', ' + phone + ') апп-д бүртгүүллээ. Тохиргоо → Хэрэглэгчид хэсгээс баталгаажуулна уу.\n\n' + (ScriptApp.getService().getUrl() || '')); } catch (e) {}
  return { ok: true, name: name };
}
function digits_(x) { return String(x === undefined || x === null ? '' : x).replace(/\D/g, ''); }
function hex_(bytes) { return bytes.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join(''); }
function secret_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function hashPin_(name, pin) { return 'h1:' + hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret_() + '|' + name + '|' + pin)); }
function pinOk_(u, pin) { return u.pin.indexOf('h1:') === 0 ? u.pin === hashPin_(u.name, pin) : (u.pin !== '' && u.pin === pin); }
function minLen_(role) { return ['Ажилчин', 'Тогооч', 'Оператор', 'Механик', 'Нярав', 'Мастер', 'Чанарын инженер', 'Инженер'].indexOf(role) >= 0 ? 6 : 8; }
function checkPin_(role, pin) {
  var n = minLen_(role);
  if (!new RegExp('^\\d{' + n + ',12}$').test(pin)) throw new Error('PIN ' + n + '-12 оронтой тоо байна');
  if (/^(\d)\1+$/.test(pin) || '0123456789012'.indexOf(pin) >= 0 || '9876543210987'.indexOf(pin) >= 0) throw new Error('Хэт амархан PIN байна (жишээ нь 111111, 123456). Өөр PIN сонгоно уу');
}
function setPin_(u, pin, mustChange) {
  var sh = sh_('Хэрэглэгч');
  sh.getRange(u.row, 3).setValue(hashPin_(u.name, pin));
  sh.getRange(u.row, 6).setValue(!!mustChange);
  dropUsers_();
}
function token_(u) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, u.name + '|' + u.pin + '|' + u.role + '|' + secret_())); }
function login_(login, pin) {
  login = String(login || '').trim(); pin = String(pin || '').trim();
  if (!login || !pin) return { error: 'bad', message: 'Утас, PIN кодоо оруулна уу' };
  var key = 'fail_' + (digits_(login) || login.toLowerCase());
  var cache = CacheService.getScriptCache(), fails = Number(cache.get(key) || 0);
  if (fails >= MAX_FAIL) return { error: 'locked', message: 'Олон удаа буруу оруулсан тул 15 минут түгжигдлээ. Дараа дахин оролдоно уу.' };
  var d = digits_(login);
  var u = users_().filter(function (x) { return (d.length >= 8 && x.phone === d) || x.name.toLowerCase() === login.toLowerCase(); })[0];
  if (!u || !pinOk_(u, pin)) {
    fails++; cache.put(key, String(fails), LOCK_SEC);
    var left = MAX_FAIL - fails;
    return { error: 'bad', message: left > 0 ? 'Утас эсвэл PIN буруу байна. ' + left + ' оролдлого үлдлээ.' : 'Олон удаа буруу оруулсан тул 15 минут түгжигдлээ.' };
  }
  cache.remove(key);
  if (u.status === PENDING) return { error: 'bad', message: 'Бүртгэл тань хараахан баталгаажаагүй байна. Хүний нөөц эсвэл захирал баталгаажуулсны дараа нэвтэрнэ.' };
  if (u.pin.indexOf('h1:') !== 0) { setPin_(u, pin, true); u = users_().filter(function (x) { return x.name === u.name; })[0]; } // хуучин ил PIN-ийг нууцлах
  return { user: { name: u.name, role: u.role, token: token_(u), mustChange: u.mustChange } };
}
function auth_(b) {
  if (!b.name || !b.token) return null;
  var u = users_().filter(function (x) { return x.name === b.name; })[0];
  return u && u.status !== PENDING && token_(u) === b.token ? u : null;
}
function changePin_(me, oldPin, newPin) {
  oldPin = String(oldPin || '').trim(); newPin = String(newPin || '').trim();
  if (!pinOk_(me, oldPin)) throw new Error('Одоогийн PIN буруу байна');
  if (oldPin === newPin) throw new Error('Шинэ PIN хуучнаасаа өөр байх ёстой');
  checkPin_(me.role, newPin);
  setPin_(me, newPin, false);
  var u = users_().filter(function (x) { return x.name === me.name; })[0];
  return { user: { name: u.name, role: u.role, token: token_(u), mustChange: false } };
}
function saveUser_(me, u, isNew) {
  var name = String(u.name || '').trim(), role = String(u.role || '').trim(), email = String(u.email || '').trim();
  var phone = digits_(u.phone), tempPin = String(u.tempPin || '').trim();
  if (!name || ROLES.indexOf(role) < 0) throw new Error('Нэр, үүргээ зөв оруулна уу');
  if (phone.length < 8) throw new Error('Утасны дугаар 8 оронтой байна');
  if (role === 'Захирал' && !isDir_(me)) throw new Error('Захирлын эрх олгох боломжгүй');
  if (isHR_(me) && ['Нягтлан', 'Админ', 'Захирал'].indexOf(role) >= 0) throw new Error('Хүний нөөц энэ эрхийг олгох боломжгүй');
  if (isAdmin_(me) && ['Нягтлан', 'Захирал'].indexOf(role) >= 0) throw new Error('Админ энэ эрхийг олгох боломжгүй');
  if (name === me.name && role !== me.role) throw new Error('Өөрийн эрхийг өөрчлөх боломжгүй');
  if (isNew && !tempPin) throw new Error('Түр PIN үүсээгүй байна');
  if (tempPin) checkPin_(role, tempPin);
  var all = users_(), ex = all.filter(function (x) { return x.name === name; })[0];
  if (isNew && ex) throw new Error('Ийм нэртэй хэрэглэгч байна');
  if (!isNew && !ex) throw new Error('Хэрэглэгч олдсонгүй');
  if (all.some(function (x) { return x.phone === phone && x.name !== name; })) throw new Error('Энэ утасны дугаар өөр хэрэглэгчид бүртгэлтэй байна');
  if (ex && ex.role === 'Захирал' && !isDir_(me)) throw new Error('Захирлын мэдээллийг засах эрхгүй');
  if (ex && isHR_(me) && (ex.role === 'Нягтлан' || isDir_(ex))) throw new Error('Энэ хэрэглэгчийг засах эрхгүй');
  if (ex && isAdmin_(me) && (ex.role === 'Нягтлан' || isDir_(ex))) throw new Error('Энэ хэрэглэгчийг засах эрхгүй');
  if (ex && ex.name === me.name && tempPin) throw new Error('Өөрийн PIN-ийг "Миний PIN солих" хэсгээс солино уу');
  var sh = sh_('Хэрэглэгч');
  var pinCell = tempPin ? hashPin_(name, tempPin) : ex.pin, must = tempPin ? true : ex.mustChange;
  var pr = u.projects === '*' || u.projects === null ? '*' : (u.projects || []).map(function (c) { return String(c).trim(); }).filter(function (c) { return c; }).join(',');
  if (!pr) throw new Error('Дор хаяж нэг төсөл сонгоно уу');
  var row = [name, role, pinCell, email, "'" + phone, must, pr, String(u.position || '').trim()];
  var ext = [String(u.surname != null ? u.surname : (ex ? ex.surname : '')).trim(), String(u.given != null ? u.given : (ex ? ex.given : '')).trim(), u.age != null && u.age !== '' ? Math.round(num_(u.age)) : (ex ? ex.age : ''),
             String(u.sex != null ? u.sex : (ex ? ex.sex : '')), u.canReq != null ? !!u.canReq : (ex ? ex.canReq : true), '', ex && ex.regAt ? ex.regAt : ''];
  if (!String(sh.getRange(1, 7).getValue())) sh.getRange(1, 7, 1, 5).setValues([['Төслүүд', 'Албан тушаал', 'Цалингийн төрөл', 'Цалин', 'Сарын ажлын өдөр']]);
  userHead_(sh);
  var sched = role === 'Инженер' && (u.sched != null ? !!u.sched : !!(ex && ex.sched));
  if (u.sched && role !== 'Инженер') throw new Error('Ажлын график хөтлөх эрхийг зөвхөн Инженерт олгоно');
  if (isNew) sh.appendRow(row.concat(['Өдөр', '', ''], ext.slice(0, 6), [new Date(), sched]));
  else { sh.getRange(ex.row, 1, 1, 8).setValues([row]); sh.getRange(ex.row, 12, 1, 6).setValues([ext.slice(0, 6)]); sh.getRange(ex.row, 19).setValue(sched); }
  if (sched) { // нэг төсөлд нэг л график хөтлөгч
    var mine = projList_(pr);
    all.forEach(function (x) {
      if (x.name === name || !x.sched) return;
      var overlap = !mine || !x.projects || x.projects.some(function (c) { return mine.indexOf(c) >= 0; });
      if (overlap) sh.getRange(x.row, 19).setValue(false);
    });
  }
  dropUsers_();
  return pubUsers_();
}
function deleteUser_(me, target) {
  if (target === me.name) throw new Error('Өөрийгөө устгах боломжгүй');
  var sh = sh_('Хэрэглэгч'), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]).trim() === target) {
      if (String(v[i][1]).trim() === 'Захирал' && !isDir_(me)) throw new Error('Захирлыг устгах эрхгүй');
      if (isHR_(me) && ['Нягтлан', 'Админ', 'Захирал'].indexOf(String(v[i][1]).trim()) >= 0) throw new Error('Энэ хэрэглэгчийг хасах эрхгүй');
      if (isAdmin_(me) && ['Нягтлан', 'Захирал'].indexOf(String(v[i][1]).trim()) >= 0) throw new Error('Энэ хэрэглэгчийг хасах эрхгүй');
      sh.deleteRow(i + 1);
    }
  }
  dropUsers_();
  return pubUsers_();
}

/* ---------- өгөгдөл ---------- */
function sh_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function fmtDate_(d) { return d instanceof Date ? Utilities.formatDate(d, 'Asia/Ulaanbaatar', 'yyyy-MM-dd HH:mm') : String(d || ''); }
function num_(x) { var n = Number(String(x === undefined ? '' : x).replace(/[,\s₮]/g, '')); return isFinite(n) ? n : 0; }
function fileData__gas_(id) {
  var f = DriveApp.getFileById(id);
  return 'data:' + f.getMimeType() + ';base64,' + Utilities.base64Encode(f.getBlob().getBytes());
}
function savePhoto__gas_(folderName, name, dataUrl) {
  var it = DriveApp.getFoldersByName(folderName);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
  var mime = dataUrl.substring(5, dataUrl.indexOf(';')) || 'image/jpeg';
  return folder.createFile(Utilities.newBlob(Utilities.base64Decode(dataUrl.split(',')[1]), mime, name)).getId();
}
function rows_() {
  var v = sh_('Хүсэлт').getDataRange().getValues().slice(1);
  return v.filter(function (r) { return r[0] && String(r[0]).indexOf('EX') !== 0; }).map(rowObj_);
}
function rowObj_(r) {
  var o = {};
  COLS.forEach(function (c, i) { o[c] = (c === 'date' || c === 'paid') ? fmtDate_(r[i]) : (c === 'amount' ? num_(r[i]) : String(r[i] === undefined ? '' : r[i])); });
  return o;
}
function projects_() {
  var psh = sh_('Төсөл'); if (psh.getMaxColumns() < 16) psh.insertColumnsAfter(psh.getMaxColumns(), 16 - psh.getMaxColumns());
  var v = psh.getRange(2, 1, Math.max(1, psh.getLastRow() - 1), 16).getValues(); // O:P = талбайн байршил
  return v.filter(function (r) { return r[0]; }).map(function (r) {
    var o = {};
    PCOLS.forEach(function (c, i) { o[c] = r[i]; });
    o.code = String(o.code); o.name = String(o.name);
    o.active = !(o.active === false || String(o.active).toUpperCase() === 'FALSE');
    ['contract','c1','c2','c3','c4','c5'].forEach(function (k) { o[k] = num_(o[k]); });
    o.pct = o.pct === '' ? 0.05 : num_(o.pct);
    o.photo = String(o.photo || '');
    o.lat = num_(r[14]); o.lng = num_(r[15]);
    return o;
  });
}
function meta_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('meta');
  if (hit) return JSON.parse(hit);
  var c = sh_('Ангилал').getDataRange().getValues().slice(1).filter(function (r) { return r[0] && String(r[3]).toUpperCase() !== 'FALSE'; });
  var b = sh_('Банк').getDataRange().getValues().slice(1).filter(function (r) { return r[0]; });
  ensureGen_();
  var res = {
    loanLimit: loanLimit_(),
    projects: projects_().filter(function (p) { return p.active; }).map(function (p) { return [p.code, p.name]; }),
    categories: c.map(function (r) { return [r[0], r[1]]; }),
    banks: b.map(function (r) { return r[0]; }),
    locRequired: locRequired_(),
    approveLimit: approveLimit_()
  };
  cache.put('meta', JSON.stringify(res), 600);
  return res;
}
function ensureGen_() {
  var sh = sh_('Төсөл'), v = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() === GEN) return;
  var next = v.length + 1;
  for (var j = 1; j < v.length; j++) if (!String(v[j][0]).trim()) { next = j + 1; break; }
  sh.getRange(next, 1, 1, 11).setValues([[GEN, GEN_NAME, true, 0, 0, 0, 0, 0, 0, 0, '']]);
}
function locRequired_() { return PropertiesService.getScriptProperties().getProperty('LOC_REQUIRED') !== 'false'; }
function locStr_(l) {
  if (!l) return '';
  var lat = Number(l.lat), lng = Number(l.lng), acc = Math.round(Number(l.acc) || 0);
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180 || (!lat && !lng)) return '';
  return lat.toFixed(6) + ',' + lng.toFixed(6) + ',' + acc;
}
function dashboard_(rowsOpt) {
  ensureGen_();
  var reqs = rowsOpt || rows_(), out = projects_(), incomes = incomesRaw_();
  out.forEach(function (p) {
    var s = { ok: { n: 0, sum: 0 }, paid: { n: 0, sum: 0 }, wait: { n: 0, sum: 0 }, no: { n: 0, sum: 0 }, cats: {}, vatSum: 0, vatN: 0, eboSum: 0, eboN: 0, nonExp: { n: 0, sum: 0, by: {} }, years: {} };
    reqs.forEach(function (r) {
      if (r.proj !== p.code) return;
      if (RT_NONEXP.indexOf(r.rtype) >= 0) { // зээл, урьдчилгаа, үндсэн төлбөр — зардал биш
        if (r.status === ST.OK || r.status === ST.PAID) { s.nonExp.n++; s.nonExp.sum += r.amount; s.nonExp.by[r.rtype] = (s.nonExp.by[r.rtype] || 0) + r.amount; }
        return;
      }
      if (r.status === ST.OK || r.status === ST.PAID) { var y = String(r.date).slice(0, 4); s.years[y] = s.years[y] || { exp: 0, inc: 0 }; s.years[y].exp += r.amount; }
      var k = r.status === ST.OK ? 'ok' : r.status === ST.PAID ? 'paid' : r.status === ST.WAIT ? 'wait' : r.status === ST.NO ? 'no' : null;
      if (!k) return;
      s[k].n++; s[k].sum += r.amount;
      if (k === 'ok' || k === 'paid') { s.cats[r.cat] = (s.cats[r.cat] || 0) + r.amount; if (r.vat === 'НӨАТ-тай') { s.vatSum += r.amount; s.vatN++; } else if (r.ebarimt === 'Тийм') { s.eboSum += r.amount; s.eboN++; } }
    });
    p.stats = s; p.hasPhoto = !!p.photo; p.pages = p.photo ? String(p.photo).split(',').filter(String).length : 0; delete p.photo;
    p.incomes = incomes.filter(function (x) { return x.proj === p.code; }).map(function (x) {
      return { id: x.id, date: x.date, desc: x.desc, amount: x.amount, vat: x.vat, by: x.by, note: x.note, hasPhoto: !!x.photo };
    });
    p.incomeSum = p.incomes.reduce(function (a, x) { return a + x.amount; }, 0);
    p.incomes.forEach(function (x) { var y = String(x.date).slice(0, 4); s.years[y] = s.years[y] || { exp: 0, inc: 0 }; s.years[y].inc += x.amount; });
  });
  return out;
}
var CONTRACT_MAX_PAGES = 40;
function contractIds_(code, photos, from) {
  return (photos || []).filter(function (d) { return typeof d === 'string' && d.indexOf('data:image') === 0; })
    .map(function (d, i) { return savePhoto_(CONTRACT_FOLDER, code + '_geree_' + (from + i + 1) + '.jpg', d); });
}
function addContractPages_(code, photos) {
  code = String(code || '').trim().toUpperCase();
  var sh = sh_('Төсөл'), v = sh.getDataRange().getValues(), idx = -1;
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() === code) idx = i;
  if (idx < 1) throw new Error('Төсөл олдсонгүй');
  var cur = String(v[idx][10] || '').split(',').filter(String);
  if (cur.length + (photos || []).length > CONTRACT_MAX_PAGES) throw new Error('Гэрээний хуулбар ' + CONTRACT_MAX_PAGES + ' хуудаснаас хэтэрч болохгүй');
  var ids = cur.concat(contractIds_(code, photos, cur.length));
  sh.getRange(idx + 1, 11).setValue(ids.join(','));
  CacheService.getScriptCache().remove('meta');
  return ids.length;
}
function saveProject_(p, photo, photos) {
  CacheService.getScriptCache().remove('meta');
  var code = String(p.code || '').trim().toUpperCase(), name = String(p.name || '').trim();
  if (!code || !name) throw new Error('Төслийн код, нэрийг оруулна уу');
  var sh = sh_('Төсөл'), v = sh.getDataRange().getValues(), idx = -1;
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() === code) idx = i;
  if (p.isNew && idx > 0) throw new Error('Ийм кодтой төсөл байна');
  var photoId = idx > 0 ? String(v[idx][10] || '') : '';
  if (photos && photos.length) {
    if (photos.length > CONTRACT_MAX_PAGES) throw new Error('Гэрээний хуулбар ' + CONTRACT_MAX_PAGES + ' хуудаснаас хэтэрч болохгүй');
    var nids = contractIds_(code, photos, 0); if (nids.length) photoId = nids.join(',');
  } else if (photo && photo.indexOf('data:image') === 0) photoId = savePhoto_(CONTRACT_FOLDER, code + '_geree.jpg', photo);
  var row = [code, name, p.active !== false, num_(p.contract), num_(p.c1), num_(p.c2), num_(p.c3), num_(p.c4), num_(p.c5), num_(p.pct), photoId];
  var site = [p.lat ? num_(p.lat) : '', p.lng ? num_(p.lng) : ''];
  if (idx > 0) { sh.getRange(idx + 1, 1, 1, row.length).setValues([row]); sh.getRange(idx + 1, 15, 1, 2).setValues([site]); }
  else {
    // L:N баганын ARRAYFORMULA доош хоосон мөр үүсгэдэг тул A баганын эхний хоосон мөрийг хайна
    var next = v.length + 1;
    for (var j = 1; j < v.length; j++) if (!String(v[j][0]).trim()) { next = j + 1; break; }
    sh.getRange(next, 1, 1, row.length).setValues([row]); sh.getRange(next, 15, 1, 2).setValues([site]);
    if (!String(sh.getRange(1, 15).getValue())) sh.getRange(1, 15, 1, 2).setValues([['Талбай өргөрөг', 'Талбай уртраг']]);
  }
  return dashboard_().filter(function (x) { return x.code === code; })[0];
}
function deleteProject_(code) {
  code = String(code || '').trim();
  var nReq = rows_().filter(function (r) { return r.proj === code; }).length;
  var nInc = incomesRaw_().filter(function (x) { return x.proj === code; }).length;
  if (nReq || nInc) throw new Error('Энэ төсөлд ' + nReq + ' хүсэлт, ' + nInc + ' санхүүжилт бүртгэлтэй тул устгах боломжгүй. Оронд нь "Засах" дотроос "Идэвхтэй" тэмдгийг арилгаж хаана уу.');
  var sh = sh_('Төсөл'), v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]).trim() === code) {
      sh.getRange(i + 1, 15, 1, 2).deleteCells(SpreadsheetApp.Dimension.ROWS);
      sh.getRange(i + 1, 1, 1, PCOLS.length).deleteCells(SpreadsheetApp.Dimension.ROWS);
      CacheService.getScriptCache().remove('meta');
      return;
    }
  }
  throw new Error('Төсөл олдсонгүй');
}

/* ---------- гүйцэтгэлийн санхүүжилт ---------- */
function incomesRaw_() {
  var sh = sh_('Санхүүжилт'); if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ICOLS.length).getValues().filter(function (r) { return r[0]; }).map(function (r) {
    var o = {}; ICOLS.forEach(function (c, i) { o[c] = r[i]; });
    o.id = String(o.id); o.date = o.date instanceof Date ? Utilities.formatDate(o.date, 'Asia/Ulaanbaatar', 'yyyy-MM-dd') : String(o.date || '');
    o.amount = num_(o.amount); o.proj = String(o.proj); o.desc = String(o.desc || ''); o.vat = String(o.vat || ''); o.by = String(o.by || ''); o.note = String(o.note || ''); o.photo = String(o.photo || '');
    return o;
  });
}
function addIncome_(me, f, photo) {
  var amount = num_(f.amount);
  if (!f.proj || !(amount > 0) || !f.date || VATS.indexOf(f.vat) < 0) throw new Error('Мэдээлэл дутуу байна');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sh_('Санхүүжилт');
    var id = 'I' + String(f.date).replace(/-/g, '').slice(2) + '-' + sh.getLastRow();
    var photoId = (photo && photo.indexOf('data:image') === 0) ? savePhoto_(INCOME_FOLDER, id + '.jpg', photo) : '';
    var d = String(f.date).split('-');
    sh.appendRow([id, new Date(+d[0], +d[1] - 1, +d[2]), f.proj, String(f.desc || '').trim(), amount, f.vat, photoId, me.name, f.note || '']);
  } finally { lock.releaseLock(); }
  return { id: id };
}
function deleteIncome_(id) {
  var sh = sh_('Санхүүжилт'), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === id) { sh.deleteRow(i + 1); return; }
  throw new Error('Олдсонгүй');
}
function mkDesc_(f) {
  return [f.proj, f.cat, f.what, f.qty].filter(function (x) { return x; }).join(' ').replace(/\s+/g, ' ').trim().slice(0, 80);
}
function create_(me, f, photo) {
  var amount = num_(f.amount);
  var rtype = String(f.rtype || ''), link = '';
  if (rtype) {
    if (RT_ALL.indexOf(rtype) < 0) throw new Error('Гүйлгээний төрөл буруу');
    if (!isBoss_(me) && !isPM_(me)) throw new Error('Зээл, урьдчилгааны хүсэлтийг нягтлан эсвэл төслийн удирдагч илгээнэ');
    f.cat = rtype; f.vat = 'НӨАТ-гүй'; f.ebarimt = rtype === RT.INT && f.ebarimt;
    if (rtype === RT.ADV) {
      var emp = activeUsers_().filter(function (u) { return u.name === f.emp; })[0];
      if (!emp) throw new Error('Урьдчилгаа авах ажилтныг сонгоно уу');
      var mo = Math.round(num_(f.months)) || 1; if (mo < 1 || mo > 12) throw new Error('Суутгах сар 1-12 байна');
      link = JSON.stringify({ emp: emp.name, months: mo });
      if (!f.what) f.what = emp.name + ' урьдчилгаа цалин';
    } else if (rtype === RT.LOAN) {
      link = JSON.stringify({ due: String(f.due || ''), rate: num_(f.rate) });
      if (!f.what) f.what = String(f.payee || '') + ' зээл';
    } else {
      var ln = loans_().filter(function (x) { return x.id === f.loan && x.type === 'Авсан'; })[0];
      if (!ln) throw new Error('Аль зээлийн төлбөр болохыг сонгоно уу');
      link = JSON.stringify({ loan: ln.id });
      if (!f.what) f.what = ln.party + ' ' + (rtype === RT.PRIN ? 'зээлийн үндсэн төлбөр' : 'зээлийн хүү');
    }
  }
  if (!f.proj || !f.cat || !f.what || !(amount > 0) || !f.payee || !f.bank || !f.acct || VATS.indexOf(f.vat) < 0 || String(f.phone || '').replace(/\D/g, '').length < 8) throw new Error('Мэдээлэл дутуу байна');
  if (f.proj !== GEN && !canProj_(me, f.proj)) throw new Error('Танд ' + f.proj + ' төсөлд хүсэлт илгээх эрх алга. Нягтланд хандана уу.');
  if (me.canReq === false) throw new Error('Танд төлбөрийн хүсэлт илгээх эрх олгогдоогүй байна.');
  var loc = locStr_(f.loc);
  if (!loc && !isBoss_(me) && locRequired_()) throw new Error('Байршил тогтоогдсонгүй. Утасны байршлаа (Location) асаагаад дахин илгээнэ үү.');
  var acct = String(f.acct).replace(/\D/g, ''), row;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sh_('Хүсэлт'), now = new Date();
    var id = 'R' + Utilities.formatDate(now, 'Asia/Ulaanbaatar', 'yyMMdd') + '-' + sh.getLastRow();
    var photoId = (photo && photo.indexOf('data:image') === 0) ? savePhoto_(PHOTO_FOLDER, id + '.jpg', photo) : '';
    row = [id, now, me.name, f.proj, f.cat, String(f.what).trim(), f.qty || '', String(f.payee).trim(), f.bank, "'" + acct, amount,
           photoId, ST.WAIT, '', '', mkDesc_(f), f.note || '', f.vat, (f.vat === 'НӨАТ-тай' || f.ebarimt) ? 'Тийм' : 'Үгүй', "'" + String(f.phone).replace(/[^\d+]/g, ''), loc, rtype, link];
    if (sh.getMaxColumns() < 23) sh.insertColumnsAfter(sh.getMaxColumns(), 23 - sh.getMaxColumns());
    if (!String(sh.getRange(1, 21).getValue())) sh.getRange(1, 21).setValue('Байршил');
    if (!String(sh.getRange(1, 22).getValue())) sh.getRange(1, 22, 1, 2).setValues([['Гүйлгээний төрөл', 'Зээлийн холбоос']]);
    sh.appendRow(row);
  } finally { lock.releaseLock(); }
  notify_(me, row[0], f, amount);
  var o = rowObj_(row); o.acct = acct; o.phone = String(f.phone).replace(/[^\d+]/g, ''); return o;
}
function setStatus_(me, id, status, note) {
  if ([ST.OK, ST.NO, ST.PAID].indexOf(status) < 0) throw new Error('Буруу төлөв');
  var sh = sh_('Хүсэлт'), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] === id) {
      var cur = v[i][12], rq = { by: String(v[i][2]), proj: String(v[i][3]), amount: num_(v[i][10]), rtype: String(v[i][21] || '') };
      if (status === ST.OK && !canApprove_(me, rq)) throw new Error(rq.by === me.name ? 'Өөрийн хүсэлтийг өөрөө батлах боломжгүй' : rq.proj === GEN ? 'Компанийн ерөнхий зардлыг зөвхөн захирал батална' : isPM_(me) && rq.amount > (RT_ALL.indexOf(rq.rtype) >= 0 ? loanLimit_() : approveLimit_()) ? (RT_ALL.indexOf(rq.rtype) >= 0 ? loanLimit_() : approveLimit_()).toLocaleString() + '₮-өөс дээш хүсэлтийг зөвхөн захирал батална' : 'Энэ хүсэлтийг батлах эрхгүй');
      if (status === ST.NO && !isBoss_(me) && !pmSees_(me, rq)) throw new Error('Буцаах эрхгүй');
      if (status === ST.PAID && !isBoss_(me)) throw new Error('Төлбөрийг зөвхөн нягтлан, захирал тэмдэглэнэ');
      if (status === ST.PAID && cur !== ST.OK) throw new Error('Эхлээд зөвшөөрөх шаардлагатай');
      if ((status === ST.OK || status === ST.NO) && cur !== ST.WAIT) throw new Error('Энэ хүсэлт аль хэдийн шийдвэрлэгдсэн');
      sh.getRange(i + 1, 13).setValue(status);
      sh.getRange(i + 1, 14).setValue(me.name);
      if (status === ST.PAID) sh.getRange(i + 1, 15).setValue(new Date());
      if (note) sh.getRange(i + 1, 17).setValue(note);
      if (status === ST.PAID && rq.rtype) loanOnPaid_(me, rowObj_(sh.getRange(i + 1, 1, 1, COLS.length).getValues()[0]));
      return rowObj_(sh.getRange(i + 1, 1, 1, COLS.length).getValues()[0]);
    }
  }
  throw new Error('Хүсэлт олдсонгүй');
}
function notify_(me, id, f, amount) {
  var to = users_().filter(function (u) { return isBoss_(u) && u.email; }).map(function (u) { return u.email; });
  if (!to.length) return;
  var url = ScriptApp.getService().getUrl() || '';
  try {
    MailApp.sendEmail(to.join(','), 'Шинэ хүсэлт: ' + f.proj + ' ' + f.what + ' — ' + amount.toLocaleString() + '₮',
      me.name + ' гүйлгээний хүсэлт илгээлээ.\n\n' +
      'Төсөл: ' + f.proj + '  Ангилал: ' + f.cat + '\n' +
      'Юу: ' + f.what + (f.qty ? ' (' + f.qty + ')' : '') + '\n' +
      'Хүлээн авагч: ' + f.payee + ', ' + f.bank + ', утас ' + f.phone + '\n' +
      'Дүн: ' + amount.toLocaleString() + '₮ (' + (f.vat === 'НӨАТ-тай' ? 'НӨАТ-тай' : (f.ebarimt ? 'НӨАТ-гүй, и-баримттай' : 'Баримтгүй')) + ')\n\n' +
      'Батлах: ' + url + '\n(ID ' + id + ')');
  } catch (e) { /* имэйлийн хязгаар хэтэрсэн ч хүсэлт бүртгэгдсэн хэвээр */ }
}

/* ---------- Нягтлангийн ГАЗ бүртгэлийг нэг удаа оруулах ----------
 * Apps Script дээр дээд талын функцийн жагсаалтаас importGazLedger сонгоод Run дарна.
 * Дахин ажиллуулсан ч өмнө орсон мөрийг давхар оруулахгүй.
 */
var GAZ_LEDGER_ID = '1ZVFJruHK5JlnTojWCRNak5XFM9_8hZK4LeY2ovN1eWc';
var GAZ_BY = 'Санаа';
function importGazLedger() {
  var src = SpreadsheetApp.openById(GAZ_LEDGER_ID);
  var res = buildGazRows_(src);
  var sh = sh_('Хүсэлт');
  var existing = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { existing[String(r[0])] = 1; });
  var rows = res.rows.filter(function (r) { return !existing[r[0]]; });
  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  var msg = 'Нягтлангийн бүртгэлд ' + res.rows.length + ' мөр, нийт ' + Math.round(res.total).toLocaleString() + ' ₮. Шинээр оруулсан: ' + rows.length + ' мөр.';
  Logger.log(msg);
  return msg;
}
function buildGazRows_(src) {
  var tabs = [
    { name: 'Материалын зардал', code: 'М', amt: 5 },
    { name: 'Техник зардал', code: 'Т', amt: 7 },
    { name: 'Цалин', code: 'Ц', amt: 6 },
    { name: 'Тээвэр, шатахуун', code: 'ШТ', amt: 4 },
    { name: 'Бусад зардал', code: 'Б', amt: 2 }
  ];
  var out = [], total = 0;
  tabs.forEach(function (tb) {
    var ws = src.getSheetByName(tb.name); if (!ws) return;
    var v = ws.getDataRange().getValues();
    var prev = new Date(2026, 2, 1);
    for (var i = 1; i < v.length; i++) {
      var r = v[i], amount = r[tb.amt];
      var d = gazDate_(r[0], prev); if (d) prev = d;
      if (typeof amount !== 'number' || !amount) continue;
      total += amount;
      var info = gazInfo_(tb.name, r);
      var vatYes = tb.name === 'Бусад зардал' && /^\s*т[ий]{1,3}[мт]\s*$/i.test(String(r[3] || ''));
      var id = 'GAZ-' + tb.code + '-' + (i + 1);
      var cat = gazCat_(tb.name, info.text, r);
      var what = info.what.replace(/\s+/g, ' ').trim().slice(0, 120) || tb.name;
      var note = ('Нягтлангийн ГАЗ бүртгэл: ' + tb.name + ', мөр ' + (i + 1) + (info.note ? '. ' + info.note : '')).slice(0, 300);
      out.push([id, prev, GAZ_BY, 'ГАЗ', cat, what, info.qty, info.payee, info.bank, '', amount, '', ST.PAID, 'Нягтлангийн бүртгэл', prev,
                mkDesc_({ proj: 'ГАЗ', cat: cat, what: what, qty: info.qty }), note, vatYes ? 'НӨАТ-тай' : 'НӨАТ-гүй', vatYes ? 'Тийм' : 'Үгүй', '']);
    }
  });
  return { rows: out, total: total };
}
function gazDate_(x, prev) {
  if (x instanceof Date) return x;
  if (x === '' || x === null || x === undefined) return null;
  var s = String(x).trim(), m = s.match(/^(\d{1,2})[.,\/-](\d+)/);
  if (!m) return null;
  var mon = +m[1], dd = m[2];
  if (mon < 1 || mon > 12) return null;
  var day;
  if (dd.length >= 2) day = +dd.slice(0, 2);
  else {
    var a = new Date(2026, mon - 1, +dd), b = new Date(2026, mon - 1, +dd * 10);
    day = (+dd * 10 <= 31 && Math.abs(b - prev) < Math.abs(a - prev)) ? +dd * 10 : +dd;
  }
  if (day < 1 || day > 31) return null;
  return new Date(2026, mon - 1, day);
}
function gazInfo_(tab, r) {
  var str = function (x) { return x === null || x === undefined ? '' : String(x).trim(); };
  var o = { what: '', qty: '', payee: '', bank: '', note: '', text: '' };
  if (tab === 'Материалын зардал') { o.what = str(r[1]); o.qty = (str(r[3]) + ' ' + str(r[2])).trim(); o.payee = str(r[6]); o.note = str(r[8]) || str(r[7]); }
  else if (tab === 'Техник зардал') { o.what = str(r[5]) || 'Техник'; var u = [str(r[1]) && r[1] + ' цаг', str(r[2]) && r[2] + ' өдөр', str(r[3]) && r[3] + ' цаг', str(r[4]) && r[4] + ' өдөр'].filter(function (x) { return x; }); o.qty = u.join(' '); o.note = str(r[8]); }
  else if (tab === 'Цалин') { o.what = 'Цалин: ' + str(r[1]) + (str(r[2]) ? ' (' + str(r[2]) + ')' : ''); o.payee = str(r[1]); o.note = str(r[7]); }
  else if (tab === 'Тээвэр, шатахуун') { o.what = str(r[1]) || 'Тээвэр, шатахуун'; o.qty = str(r[2]) ? r[2] + ' л' : ''; o.note = str(r[5]); }
  else { var t = str(r[1]); o.what = t; var m = t.match(/\(([A-Z][A-Z .&()]*?BANK)\s+([^()]*)\)\s*$/); if (m) { o.bank = m[1].trim(); o.payee = m[2].trim(); o.what = t.replace(m[0], '').trim(); } o.note = str(r[4]) || (str(r[3]) && !/^\s*т[ий]{1,3}[мт]\s*$/i.test(str(r[3])) && str(r[3]).toLowerCase() !== 'үгүй' ? str(r[3]) : ''); }
  o.text = (o.what + ' ' + o.payee + ' ' + o.note).toLowerCase();
  return o;
}
function gazCat_(tab) {
  return { 'Материалын зардал': 'МАТ', 'Техник зардал': 'ТЕХ', 'Цалин': 'ЦАЛ', 'Тээвэр, шатахуун': 'ТШ', 'Бусад зардал': 'БУС' }[tab] || 'БУС';
}


/* ---------- Техник, түлшний бүртгэл ----------
 * Техник хуудас: A ID | B Төсөл | C Улсын дугаар | D Нэр | E Төрөл (Механизм / Тээврийн хэрэгсэл) | F Норм | G Зураг | H Идэвхтэй | I Тэмдэглэл | J Бүртгэсэн
 *   Норм: механизм бол л/мот.цаг, тээврийн хэрэгсэл бол л/100км. Хоосон бол тухайн техникийн өөрийн дунджаар харьцуулна.
 * Түлш хуудас:   A ID | B Огноо | C Олгосон | D Төсөл | E Техник ID | F Литр | G Заалт (мот/цаг эсвэл км) | H Заалтын зураг | I Тэмдэглэл
 * Зарцуулалт = энэ удаа хийсэн литр / (энэ удаагийн заалт - өмнөх удаагийн заалт).
 */
var ECOLS = ['id','proj','plate','name','kind','norm','photo','active','note','created','own','rate','rateType','monthDays','owner','ownerPhone','bank','acct','fuelBy','start'];
var ACOLS = ['id','date','eq','proj','status','by','note','ts'];
var ATT = ['Ажилласан', 'Хагас өдөр', 'Сул зогссон', 'Эвдэрсэн', 'Амралт, цаг агаар'];
var FCOLS = ['id','date','by','proj','eq','liters','meter','photo','note'];
var KINDS = ['Механизм', 'Тээврийн хэрэгсэл'];
var EQUIP_FOLDER = 'Аасгалдай техник', FUEL_FOLDER = 'Аасгалдай түлш';
var FUEL_TOL = 0.15; // нормоос ±15% бол хэвийн
function ensureSheet_(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function eqSheet_() {
  var sh = ensureSheet_('Техник', ['ID','Төсөл','Улсын дугаар','Нэр','Төрөл','Норм','Зураг','Идэвхтэй','Тэмдэглэл','Бүртгэсэн','Эзэмшил','Түрээсийн үнэ','Үнийн нэгж','Сарын ажлын өдөр','Эзэмшигч','Эзэмшигчийн утас','Банк','Данс','Түлш хариуцах','Ирсэн огноо']);
  if (sh.getLastColumn() < ECOLS.length) sh.getRange(1, 11, 1, 10).setValues([['Эзэмшил','Түрээсийн үнэ','Үнийн нэгж','Сарын ажлын өдөр','Эзэмшигч','Эзэмшигчийн утас','Банк','Данс','Түлш хариуцах','Ирсэн огноо']]).setFontWeight('bold');
  return sh;
}
function attSheet_() { return ensureSheet_('Ирц', ['ID','Огноо','Техник ID','Төсөл','Төлөв','Тэмдэглэсэн','Тэмдэглэл','Хугацаа']); }
function ymd_(d) { return d instanceof Date ? Utilities.formatDate(d, 'Asia/Ulaanbaatar', 'yyyy-MM-dd') : String(d || '').slice(0, 10); }
function atts_() {
  var sh = attSheet_(); if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ACOLS.length).getValues().filter(function (r) { return r[0]; }).map(function (r) {
    return { id: String(r[0]), date: ymd_(r[1]), eq: String(r[2]), proj: String(r[3]), status: String(r[4]), by: String(r[5] || ''), note: String(r[6] || '') };
  });
}
function mark_(me, boss, a) {
  var eq = equips_().filter(function (x) { return x.id === a.eq; })[0];
  if (!eq) throw new Error('Техник олдсонгүй');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(a.date))) throw new Error('Огноо буруу');
  var today = ymd_(new Date());
  if (a.date > today) throw new Error('Ирээдүйн өдрийг тэмдэглэх боломжгүй');
  if (!boss) { var lim = ymd_(new Date(Date.now() - 3 * 864e5)); if (a.date < lim) throw new Error('3 хоногоос өмнөх өдрийг нягтлан, захирал л засна'); }
  if (a.status && ATT.indexOf(a.status) < 0) throw new Error('Буруу төлөв');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = attSheet_(), v = sh.getDataRange().getValues(), idx = -1;
    for (var i = 1; i < v.length; i++) if (String(v[i][2]) === eq.id && ymd_(v[i][1]) === a.date) idx = i;
    if (!a.status) { if (idx > 0) sh.deleteRow(idx + 1); return; }
    var p = a.date.split('-');
    var row = [idx > 0 ? v[idx][0] : 'A' + a.date.replace(/-/g, '').slice(2) + '-' + eq.id, new Date(+p[0], +p[1] - 1, +p[2]), eq.id, eq.proj, a.status, me.name, String(a.note || '').trim(), new Date()];
    if (idx > 0) sh.getRange(idx + 1, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  } finally { lock.releaseLock(); }
}
function fuelSheet_() { return ensureSheet_('Түлш', ['ID','Огноо','Олгосон','Төсөл','Техник ID','Литр','Заалт','Заалтын зураг','Тэмдэглэл']); }
function equips_() {
  var sh = eqSheet_(); if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ECOLS.length).getValues().filter(function (r) { return r[0]; }).map(function (r) {
    var o = {}; ECOLS.forEach(function (c, i) { o[c] = r[i]; });
    o.id = String(o.id); o.proj = String(o.proj); o.plate = String(o.plate || ''); o.name = String(o.name || '');
    o.kind = KINDS.indexOf(String(o.kind)) >= 0 ? String(o.kind) : KINDS[0];
    o.norm = num_(o.norm); o.photo = String(o.photo || ''); o.note = String(o.note || '');
    o.active = !(o.active === false || String(o.active).toUpperCase() === 'FALSE');
    o.own = String(o.own) === 'Түрээсийн' ? 'Түрээсийн' : 'Өөрийн';
    o.rate = num_(o.rate); o.rateType = String(o.rateType) === 'Сар' ? 'Сар' : 'Өдөр'; o.monthDays = num_(o.monthDays) || 26;
    o.owner = String(o.owner || ''); o.ownerPhone = String(o.ownerPhone || ''); o.bank = String(o.bank || ''); o.acct = String(o.acct || '');
    o.fuelBy = String(o.fuelBy) === 'Түрээслэгч' ? 'Түрээслэгч' : 'Бид'; o.start = ymd_(o.start); o.created = '';
    return o;
  });
}
function fuelsRaw_() {
  var sh = fuelSheet_(); if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, FCOLS.length).getValues().filter(function (r) { return r[0]; }).map(function (r) {
    var o = {}; FCOLS.forEach(function (c, i) { o[c] = r[i]; });
    o.id = String(o.id); o.date = fmtDate_(o.date); o.by = String(o.by || ''); o.proj = String(o.proj); o.eq = String(o.eq);
    o.liters = num_(o.liters); o.meter = num_(o.meter); o.photo = String(o.photo || ''); o.note = String(o.note || '');
    return o;
  });
}
/* Нэг техникийн бүртгэлүүдийг заалтаар эрэмбэлээд зарцуулалт, үнэлгээг тооцно */
function fuelCalc_(eq, list) {
  list = list.slice().sort(function (a, b) { return a.meter - b.meter || (a.date < b.date ? -1 : 1); });
  var per = eq.kind === KINDS[1] ? 100 : 1, rates = [];
  list.forEach(function (f, i) {
    f.delta = 0; f.rate = 0; f.base = 0; f.mark = i === 0 ? 'first' : '';
    if (i === 0) return;
    f.delta = Math.round((f.meter - list[i - 1].meter) * 10) / 10;
    if (!(f.delta > 0)) { f.mark = 'bad'; return; }
    f.rate = Math.round(f.liters / f.delta * per * 100) / 100;
    f.base = eq.norm > 0 ? eq.norm : (rates.length >= 2 ? rates.reduce(function (a, x) { return a + x; }, 0) / rates.length : 0);
    if (!f.base) f.mark = 'learn';
    else f.mark = f.rate > f.base * (1 + FUEL_TOL) ? 'high' : f.rate < f.base * (1 - FUEL_TOL) ? 'low' : 'ok';
    rates.push(f.rate);
  });
  var valid = list.filter(function (f) { return f.rate > 0; });
  var used = valid.reduce(function (a, f) { return a + f.liters; }, 0), work = valid.reduce(function (a, f) { return a + f.delta; }, 0);
  return {
    list: list,
    liters: list.reduce(function (a, f) { return a + f.liters; }, 0),
    avg: work > 0 ? Math.round(used / work * per * 100) / 100 : 0,
    work: work,
    lastMeter: list.length ? list[list.length - 1].meter : null,
    last: list.length ? list[list.length - 1] : null,
    high: valid.filter(function (f) { return f.mark === 'high'; }).length
  };
}
function fuelData_(me, boss) {
  var eqs = equips_(), fuels = fuelsRaw_();
  var out = eqs.map(function (e) {
    var c = fuelCalc_(e, fuels.filter(function (f) { return f.eq === e.id; }));
    var entries = c.list.map(function (f) {
      return { id: f.id, date: f.date, by: f.by, liters: f.liters, meter: f.meter, delta: f.delta, rate: f.rate, base: Math.round(f.base * 100) / 100, mark: f.mark, note: f.note, hasPhoto: !!f.photo };
    });
    if (!boss && !isMech_(me) && !isStore_(me) && !isPM_(me)) entries = entries.filter(function (f) { return f.by === me.name; });
    return { id: e.id, proj: e.proj, plate: e.plate, name: e.name, kind: e.kind, norm: e.norm, note: e.note, active: e.active, hasPhoto: !!e.photo,
             own: e.own, rate: boss ? e.rate : 0, rateType: e.rateType, monthDays: e.monthDays, owner: e.owner, ownerPhone: boss ? e.ownerPhone : '', bank: boss ? e.bank : '', acct: boss ? e.acct : '', fuelBy: e.fuelBy, start: e.start,
             liters: c.liters, avg: c.avg, work: c.work, lastMeter: c.lastMeter, lastMark: c.last ? c.last.mark : '', lastRate: c.last ? c.last.rate : 0, high: c.high,
             entries: entries.slice(-200) };
  });
  var lim = ymd_(new Date(Date.now() - 400 * 864e5));
  var att = atts_().filter(function (a) { return a.date >= lim; }).map(function (a) { return [a.date, a.eq, a.status, a.by, a.note]; });
  return { equips: out, kinds: KINDS, tol: FUEL_TOL, att: att, attTypes: ATT, today: ymd_(new Date()) };
}
function saveEquip_(e, photo) {
  var plate = String(e.plate || '').trim().toUpperCase(), name = String(e.name || '').trim();
  if (!e.proj || !plate || !name || KINDS.indexOf(e.kind) < 0) throw new Error('Төсөл, улсын дугаар, нэр, төрлийг оруулна уу');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = eqSheet_(), v = sh.getDataRange().getValues(), idx = -1;
    for (var i = 1; i < v.length; i++) {
      if (e.id && String(v[i][0]) === e.id) idx = i;
      else if (String(v[i][2]).trim().toUpperCase() === plate && String(v[i][1]) === e.proj) throw new Error('Энэ төсөлд ' + plate + ' дугаартай техник бүртгэлтэй байна');
    }
    var id = idx > 0 ? String(v[idx][0]) : 'T' + Utilities.formatDate(new Date(), 'Asia/Ulaanbaatar', 'yyMMddHHmmss');
    var photoId = idx > 0 ? String(v[idx][6] || '') : '';
    if (photo && photo.indexOf('data:image') === 0) photoId = savePhoto_(EQUIP_FOLDER, plate + '_' + id + '.jpg', photo);
    var rent = e.own === 'Түрээсийн';
    if (rent && !(num_(e.rate) > 0)) throw new Error('Түрээсийн үнээ оруулна уу');
    var sd = String(e.start || '').split('-');
    var row = [id, e.proj, plate, name, e.kind, num_(e.norm) || '', photoId, e.active !== false, String(e.note || '').trim(), idx > 0 ? v[idx][9] : new Date(),
               rent ? 'Түрээсийн' : 'Өөрийн', rent ? num_(e.rate) : '', rent ? (e.rateType === 'Сар' ? 'Сар' : 'Өдөр') : '', rent && e.rateType === 'Сар' ? (num_(e.monthDays) || 26) : '',
               rent ? String(e.owner || '').trim() : '', rent ? "'" + String(e.ownerPhone || '').replace(/[^\d+]/g, '') : '', rent ? String(e.bank || '') : '', rent ? "'" + String(e.acct || '').replace(/\D/g, '') : '',
               rent ? (e.fuelBy === 'Түрээслэгч' ? 'Түрээслэгч' : 'Бид') : '', rent && sd.length === 3 ? new Date(+sd[0], +sd[1] - 1, +sd[2]) : ''];
    if (idx > 0) sh.getRange(idx + 1, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  } finally { lock.releaseLock(); }
}
function deleteEquip_(id) {
  var n = fuelsRaw_().filter(function (f) { return f.eq === id; }).length + atts_().filter(function (a) { return a.eq === id; }).length;
  if (n) throw new Error('Энэ техникт ' + n + ' түлш, ирцийн бүртгэл байгаа тул устгах боломжгүй. Оронд нь "Засах" дотроос "Идэвхтэй" тэмдгийг арилгана уу.');
  var sh = eqSheet_(), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === id) { sh.deleteRow(i + 1); return; }
  throw new Error('Техник олдсонгүй');
}
function addFuel_(me, f, photo) {
  var liters = num_(f.liters), meter = num_(f.meter);
  var eq = equips_().filter(function (x) { return x.id === f.eq; })[0];
  if (!eq || !eq.active) throw new Error('Техник олдсонгүй');
  if (!(liters > 0) || !(meter > 0)) throw new Error('Литр болон заалтыг оруулна уу');
  if (liters > 2000) throw new Error('Литр хэт их байна, шалгана уу');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var prev = fuelsRaw_().filter(function (x) { return x.eq === eq.id; }).reduce(function (m, x) { return Math.max(m, x.meter); }, 0);
    var unit = eq.kind === KINDS[1] ? 'км' : 'мот/цаг';
    if (prev && meter <= prev) throw new Error('Заалт өмнөх удаагийнхаас (' + prev + ' ' + unit + ') их байх ёстой');
    var sh = fuelSheet_(), now = new Date();
    var id = 'F' + Utilities.formatDate(now, 'Asia/Ulaanbaatar', 'yyMMdd') + '-' + sh.getLastRow();
    var photoId = (photo && photo.indexOf('data:image') === 0) ? savePhoto_(FUEL_FOLDER, id + '.jpg', photo) : '';
    sh.appendRow([id, now, me.name, eq.proj, eq.id, liters, meter, photoId, String(f.note || '').trim()]);
  } finally { lock.releaseLock(); }
}
function deleteFuel_(id) {
  var sh = fuelSheet_(), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === id) { sh.deleteRow(i + 1); return; }
  throw new Error('Олдсонгүй');
}


/* ---------- ХАБЭА: өдрийн заавар, аюултай нөхцөлийн мэдээлэл ----------
 * ХАБ заавар:     A ID | B Огноо | C Гарчиг | D Текст | E Төсөл (хоосон = бүх төсөл) | F Бичсэн | G Шинэчилсэн
 * ХАБ танилцсан:  A Заавар ID | B Нэр | C Хугацаа
 * ХАБ мэдээлэл:   A ID | B Огноо | C Илгээсэн | D Төсөл | E Төрөл | F Хаана | G GPS | H Тайлбар | I Зураг | J Төлөв | K Авсан арга хэмжээ | L Шийдвэрлэсэн | M Шийдсэн огноо
 * Заавар бичих, засах эрх зөвхөн "ХАБ инженер"-т. Захирал, админ мэдээллийг харж, төлөв өөрчилнө.
 */
var HAB_TYPES = ['Аюултай нөхцөл', 'Осолд дөхсөн тохиолдол', 'Осол гарсан'];
var HAB_ST = ['Шинэ', 'Арга хэмжээ авч байна', 'Шийдвэрлэсэн'];
var HAB_FOLDER = 'Аасгалдай ХАБЭА';
function habNoteSheet_() { return ensureSheet_('ХАБ заавар', ['ID','Огноо','Гарчиг','Текст','Төсөл','Бичсэн','Шинэчилсэн']); }
function habReadSheet_() { return ensureSheet_('ХАБ танилцсан', ['Заавар ID','Нэр','Хугацаа']); }
function habRepSheet_() { return ensureSheet_('ХАБ мэдээлэл', ['ID','Огноо','Илгээсэн','Төсөл','Төрөл','Хаана','GPS','Тайлбар','Зураг','Төлөв','Авсан арга хэмжээ','Шийдвэрлэсэн','Шийдсэн огноо']); }
function rowsOf_(sh, n) { return sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, n).getValues().filter(function (r) { return r[0]; }); }
function habNotes_() {
  return rowsOf_(habNoteSheet_(), 7).map(function (r) {
    return { id: String(r[0]), date: ymd_(r[1]), title: String(r[2] || ''), text: String(r[3] || ''), proj: String(r[4] || ''), by: String(r[5] || ''), upd: fmtDate_(r[6]) };
  });
}
function habReads_() { return rowsOf_(habReadSheet_(), 3).map(function (r) { return { id: String(r[0]), name: String(r[1]), at: fmtDate_(r[2]) }; }); }
function habReports_() {
  return rowsOf_(habRepSheet_(), 13).map(function (r) {
    return { id: String(r[0]), date: fmtDate_(r[1]), by: String(r[2] || ''), proj: String(r[3] || ''), type: String(r[4] || ''), where: String(r[5] || ''), loc: String(r[6] || ''),
             desc: String(r[7] || ''), photo: String(r[8] || ''), status: String(r[9] || HAB_ST[0]), action: String(r[10] || ''), handler: String(r[11] || ''), done: fmtDate_(r[12]) };
  });
}
function habBadge_(me) {
  var t = ymd_(new Date()), mine = {};
  habReads_().forEach(function (x) { if (x.name === me.name) mine[x.id] = 1; });
  var unread = isHab_(me) ? 0 : habNotes_().filter(function (n) { return n.date === t && !mine[n.id] && (isDir_(me) || !n.proj || canProj_(me, n.proj)); }).length;
  var out = { unread: unread };
  if (isHab_(me) || isDir_(me)) out.newReports = habReports_().filter(function (x) { return x.status === HAB_ST[0]; }).length;
  return out;
}
function habData_(me) {
  var t = ymd_(new Date()), lim = ymd_(new Date(Date.now() - 90 * 864e5)), manager = isHab_(me) || isDir_(me);
  var reads = habReads_(), notes = habNotes_().filter(function (n) { return n.date >= lim && (isHab_(me) || n.date <= t) && (manager || !n.proj || canProj_(me, n.proj)); });
  notes.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
  notes.forEach(function (n) {
    var rs = reads.filter(function (x) { return x.id === n.id; });
    n.read = rs.some(function (x) { return x.name === me.name; });
    if (manager) n.readers = rs.map(function (x) { return [x.name, x.at]; });
  });
  var reps = habReports_().filter(function (x) { return manager || x.by === me.name; }).map(function (x) {
    var o = {}; for (var k in x) o[k] = x[k]; o.hasPhoto = !!x.photo; delete o.photo; return o;
  }).reverse().slice(0, 300);
  var out = { today: t, notes: notes, reports: reps, types: HAB_TYPES, statuses: HAB_ST };
  if (manager) out.people = activeUsers_().map(function (u) { return u.name; });
  return out;
}
function habSave_(me, n) {
  var title = String(n.title || '').trim(), text = String(n.text || '').trim(), date = String(n.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !text) throw new Error('Огноо, зааврын текстээ оруулна уу');
  var sh = habNoteSheet_(), v = sh.getDataRange().getValues(), idx = -1;
  for (var i = 1; i < v.length; i++) if (n.id && String(v[i][0]) === n.id) idx = i;
  var p = date.split('-');
  var id = idx > 0 ? String(v[idx][0]) : 'H' + Utilities.formatDate(new Date(), 'Asia/Ulaanbaatar', 'yyMMddHHmmss');
  var row = [id, new Date(+p[0], +p[1] - 1, +p[2]), title || 'Өдрийн ХАБ заавар', text, String(n.proj || ''), me.name, new Date()];
  if (idx > 0) sh.getRange(idx + 1, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
}
function habDelete_(id) {
  var sh = habNoteSheet_(), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === id) { sh.deleteRow(i + 1); return; }
}
function habRead_(me, id) {
  if (habReads_().some(function (x) { return x.id === id && x.name === me.name; })) return;
  habReadSheet_().appendRow([id, me.name, new Date()]);
}
function habReport_(me, r, photo) {
  if (HAB_TYPES.indexOf(r.type) < 0 || !String(r.desc || '').trim()) throw new Error('Төрөл, тайлбараа оруулна уу');
  if (!photo || photo.indexOf('data:image') !== 0) throw new Error('Зургаар баримтжуулна уу');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = habRepSheet_(), now = new Date();
    var id = 'S' + Utilities.formatDate(now, 'Asia/Ulaanbaatar', 'yyMMdd') + '-' + sh.getLastRow();
    var photoId = savePhoto_(HAB_FOLDER, id + '.jpg', photo);
    sh.appendRow([id, now, me.name, String(r.proj || ''), r.type, String(r.where || '').trim(), locStr_(r.loc), String(r.desc).trim(), photoId, HAB_ST[0], '', '', '']);
  } finally { lock.releaseLock(); }
  var to = users_().filter(function (u) { return (isHab_(u) || isDir_(u)) && u.email; }).map(function (u) { return u.email; });
  if (to.length) try { MailApp.sendEmail(to.join(','), 'ХАБЭА: ' + r.type + ' — ' + (r.proj || ''), me.name + ' мэдээлэв:\n\n' + r.desc + (r.where ? '\nХаана: ' + r.where : '') + '\n\n' + (ScriptApp.getService().getUrl() || '')); } catch (e) {}
}
function habStatus_(me, id, status, action) {
  if (HAB_ST.indexOf(status) < 0) throw new Error('Буруу төлөв');
  var sh = habRepSheet_(), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === id) {
    sh.getRange(i + 1, 10, 1, 4).setValues([[status, String(action || '').trim(), me.name, status === HAB_ST[2] ? new Date() : '']]);
    return;
  }
  throw new Error('Олдсонгүй');
}


/* ---------- Хүний нөөц: ажилтны ирц, чөлөө, үүрэг даалгавар, цалин ----------
 * Ажлын ирц: A ID | B Огноо | C Нэр | D Төлөв | E Согтуу | F Тэмдэглэсэн | G Тэмдэглэл | H Хугацаа
 * Чөлөө:     A ID | B Илгээсэн | C Эхлэх | D Дуусах | E Шалтгаан | F Төлөв | G Шийдсэн | H Шийдсэн огноо | I Тайлбар | J Илгээсэн огноо
 * Даалгавар: A ID | B Огноо | C Өгсөн | D Хүлээн авагч | E Даалгавар | F Хугацаа | G Төсөл | H Төлөв | I Ойлгосон | J Тайлан огноо | K Тайлан | L Багц
 * Цалингийн тохиргоо (НДШ, ХХОАТ, хөнгөлөлт) Script Properties-ийн PAYCFG-д хадгалагдана; нягтлан апп-аас засна.
 */
var PST = ['Ажилласан', 'Хагас өдөр', 'Ажиллаагүй', 'Чөлөөтэй', 'Өвчтэй', 'Амралт'];
var TST = { NEW: 'Шинэ', ACK: 'Ойлгосон', DONE: 'Биелүүлсэн', NOT: 'Биелүүлээгүй' };
function patSheet_() { return ensureSheet_('Ажлын ирц', ['ID','Огноо','Нэр','Төлөв','Согтуу','Тэмдэглэсэн','Тэмдэглэл','Хугацаа']); }
function leaveSheet_() { return ensureSheet_('Чөлөө', ['ID','Илгээсэн','Эхлэх','Дуусах','Шалтгаан','Төлөв','Шийдсэн','Шийдсэн огноо','Тайлбар','Илгээсэн огноо']); }
function taskSheet_() { return ensureSheet_('Даалгавар', ['ID','Огноо','Өгсөн','Хүлээн авагч','Даалгавар','Хугацаа','Төсөл','Төлөв','Ойлгосон','Тайлан огноо','Тайлан','Багц']); }
function pats_() { return rowsOf_(patSheet_(), 8).map(function (r) { return { date: ymd_(r[1]), name: String(r[2]), status: String(r[3]), drunk: r[4] === true || String(r[4]).toUpperCase() === 'TRUE', by: String(r[5] || ''), note: String(r[6] || '') }; }); }
function leaves_() { return rowsOf_(leaveSheet_(), 10).map(function (r) { return { id: String(r[0]), by: String(r[1]), from: ymd_(r[2]), to: ymd_(r[3]), reason: String(r[4] || ''), status: String(r[5] || 'Хүлээгдэж буй'), decider: String(r[6] || ''), decided: fmtDate_(r[7]), note: String(r[8] || ''), at: fmtDate_(r[9]) }; }); }
function tasks_() { return rowsOf_(taskSheet_(), 12).map(function (r) { return { id: String(r[0]), date: fmtDate_(r[1]), from: String(r[2]), to: String(r[3]), text: String(r[4] || ''), due: ymd_(r[5]), proj: String(r[6] || ''), status: String(r[7] || TST.NEW), ack: fmtDate_(r[8]), rep: fmtDate_(r[9]), report: String(r[10] || ''), batch: String(r[11] || '') }; }); }
function payCfg_() {
  var d = { ee: 11.5, er: 12.5, pit: 10, cap: 0, credits: [[500000, 20000], [1000000, 18000], [1500000, 15000], [2000000, 13000], [2500000, 10000], [3000000, 8000]] };
  try { var c = JSON.parse(PropertiesService.getScriptProperties().getProperty('PAYCFG') || '{}'); for (var k in c) d[k] = c[k]; } catch (e) {}
  return d;
}
function visiblePeople_(me) {
  var all = activeUsers_();
  if (!(isPM_(me) || isMaster_(me)) || !me.projects) return all;
  return all.filter(function (u) { return u.name === me.name || !u.projects || u.projects.some(function (c) { return me.projects.indexOf(c) >= 0; }); });
}
function rating_(tasks, name) {
  var t = ymd_(new Date()), lim = ymd_(new Date(Date.now() - 60 * 864e5));
  var mine = tasks.filter(function (x) { return x.to === name && x.due >= lim && (x.due < t || x.status === TST.DONE || x.status === TST.NOT); });
  if (!mine.length) return { n: 0, score: null, label: 'Үнэлгээгүй' };
  var pts = 0, onTime = 0, late = 0, miss = 0;
  mine.forEach(function (x) {
    if (x.status === TST.DONE) { if (!x.rep || x.rep.slice(0, 10) <= x.due) { pts += 1; onTime++; } else { pts += 0.5; late++; } }
    else miss++;
  });
  var sc = Math.round(pts / mine.length * 100);
  return { n: mine.length, score: sc, onTime: onTime, late: late, miss: miss, label: sc >= 90 ? 'Маш сайн' : sc >= 75 ? 'Сайн' : sc >= 50 ? 'Дунд' : 'Муу' };
}
function peopleData_(me) {
  var t = ymd_(new Date()), lim = ymd_(new Date(Date.now() - 400 * 864e5)), payroll = isBoss_(me);
  var ppl = visiblePeople_(me), names = {}; ppl.forEach(function (u) { names[u.name] = 1; });
  var tasks = tasks_();
  var people = ppl.map(function (u) {
    var o = { name: u.name, role: u.role, position: u.position, projects: u.projects, phone: u.phone, rating: rating_(tasks, u.name) };
    if (payroll) { o.wageType = u.wageType; o.wage = u.wage; o.wdays = u.wdays; }
    return o;
  });
  var tlim = ymd_(new Date(Date.now() - 120 * 864e5));
  return {
    today: t, people: people, statuses: PST,
    att: pats_().filter(function (a) { return names[a.name] && a.date >= lim; }).map(function (a) { return [a.date, a.name, a.status, a.drunk ? 1 : 0, a.by, a.note]; }),
    leaves: leaves_().filter(function (l) { return names[l.by]; }).reverse().slice(0, 300),
    tasks: tasks.filter(function (x) { return x.due >= tlim && (names[x.to] || x.from === me.name); }).reverse(),
    cfg: payroll ? payCfg_() : null,
    adv: payroll ? loanList_().filter(function (l) { return l.type === 'Урьдчилгаа'; }) : null,
    canMark: isDir_(me) || isHR_(me) || isPM_(me) || isMaster_(me), canTask: isDir_(me) || isPM_(me) || isMaster_(me), canLeave: isBoss_(me) || isPM_(me), payroll: payroll
  };
}
function markP_(me, a) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(a.date))) throw new Error('Огноо буруу');
  if (a.date > ymd_(new Date())) throw new Error('Ирээдүйн өдрийг тэмдэглэх боломжгүй');
  if (a.status && PST.indexOf(a.status) < 0) throw new Error('Буруу төлөв');
  if (!visiblePeople_(me).some(function (u) { return u.name === a.name; })) throw new Error('Энэ ажилтныг тэмдэглэх эрхгүй');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = patSheet_(), v = sh.getDataRange().getValues(), idx = -1;
    for (var i = 1; i < v.length; i++) if (String(v[i][2]) === a.name && ymd_(v[i][1]) === a.date) idx = i;
    if (!a.status) { if (idx > 0) sh.deleteRow(idx + 1); return; }
    var p = a.date.split('-');
    var row = [idx > 0 ? v[idx][0] : 'P' + a.date.replace(/-/g, '').slice(2) + '-' + Utilities.getUuid().slice(0, 6), new Date(+p[0], +p[1] - 1, +p[2]), a.name, a.status, !!a.drunk, me.name, String(a.note || '').trim(), new Date()];
    if (idx > 0) sh.getRange(idx + 1, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  } finally { lock.releaseLock(); }
}
function myData_(me) {
  var t = ymd_(new Date()), lim = ymd_(new Date(Date.now() - 400 * 864e5)), tasks = tasks_();
  var att = pats_().filter(function (a) { return a.name === me.name && a.date >= lim; });
  var d7 = ymd_(new Date(Date.now() - 6 * 864e5)), d30 = ymd_(new Date(Date.now() - 29 * 864e5));
  return {
    today: t, statuses: PST, position: me.position,
    att: att.map(function (a) { return [a.date, a.name, a.status, a.drunk ? 1 : 0, a.by, a.note]; }),
    drunk7: att.filter(function (a) { return a.drunk && a.date >= d7; }).length,
    drunk30: att.filter(function (a) { return a.drunk && a.date >= d30; }).length,
    tasks: tasks.filter(function (x) { return x.to === me.name; }).reverse().slice(0, 100),
    leaves: leaves_().filter(function (l) { return l.by === me.name; }).reverse().slice(0, 50),
    rating: rating_(tasks, me.name),
    loans: loanList_().filter(function (l) { return l.party === me.name && l.type !== 'Авсан'; }).map(function (l) { return { type: l.type, amount: l.amount, date: l.date, bal: l.bal, months: l.months, txs: l.txs }; })
  };
}
function hrBadge_(me) {
  var t = ymd_(new Date()), out = {};
  var mt = tasks_().filter(function (x) { return x.to === me.name; });
  out.newTasks = mt.filter(function (x) { return x.status === TST.NEW; }).length;
  out.dueTasks = mt.filter(function (x) { return x.status === TST.ACK && x.due <= t; }).length;
  if (isBoss_(me) || isPM_(me)) {
    var names = {}; visiblePeople_(me).forEach(function (u) { names[u.name] = 1; });
    out.leaves = leaves_().filter(function (l) { return l.status === 'Хүлээгдэж буй' && names[l.by]; }).length;
  }
  return out;
}
function leaveReq_(me, lv) {
  var re = /^\d{4}-\d{2}-\d{2}$/;
  if (!re.test(String(lv.from)) || !re.test(String(lv.to)) || lv.to < lv.from) throw new Error('Огноогоо зөв сонгоно уу');
  if (!String(lv.reason || '').trim()) throw new Error('Шалтгаанаа бичнэ үү');
  var f = lv.from.split('-'), g = lv.to.split('-');
  leaveSheet_().appendRow(['L' + Utilities.formatDate(new Date(), 'Asia/Ulaanbaatar', 'yyMMddHHmmss'), me.name, new Date(+f[0], +f[1] - 1, +f[2]), new Date(+g[0], +g[1] - 1, +g[2]), String(lv.reason).trim(), 'Хүлээгдэж буй', '', '', '', new Date()]);
  var to = users_().filter(function (u) { return (isPM_(u) || u.role === 'Нягтлан' || isDir_(u)) && u.email; }).map(function (u) { return u.email; });
  if (to.length) try { MailApp.sendEmail(to.join(','), 'Чөлөөний хүсэлт: ' + me.name, me.name + ' ' + lv.from + ' — ' + lv.to + ' чөлөө хүсэв.\nШалтгаан: ' + lv.reason + '\n\n' + (ScriptApp.getService().getUrl() || '')); } catch (e) {}
}
function leaveDecide_(me, id, ok, note) {
  var sh = leaveSheet_(), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === id) {
    if (String(v[i][5]) !== 'Хүлээгдэж буй') throw new Error('Аль хэдийн шийдвэрлэгдсэн');
    if (!ok && !String(note || '').trim()) throw new Error('Татгалзах шалтгаанаа бичнэ үү');
    sh.getRange(i + 1, 6, 1, 4).setValues([[ok ? 'Зөвшөөрсөн' : 'Татгалзсан', me.name, new Date(), String(note || '').trim()]]);
    if (ok) {
      var name = String(v[i][1]), d = new Date(v[i][2]), end = new Date(v[i][3]), have = {};
      pats_().forEach(function (a) { if (a.name === name) have[a.date] = 1; });
      var rows = [];
      while (d <= end && rows.length < 62) { var ds = ymd_(d); if (!have[ds]) rows.push(['P' + ds.replace(/-/g, '').slice(2) + '-' + Utilities.getUuid().slice(0, 6), new Date(d), name, 'Чөлөөтэй', false, me.name, 'Чөлөөний хүсэлтээр', new Date()]); d = new Date(d.getTime() + 864e5); }
      if (rows.length) { var ps = patSheet_(); ps.getRange(ps.getLastRow() + 1, 1, rows.length, 8).setValues(rows); }
    }
    return;
  }
  throw new Error('Олдсонгүй');
}
function taskSend_(me, t) {
  var to = (t.to || []).filter(function (n) { return n; }), text = String(t.text || '').trim(), due = String(t.due || '');
  if (!to.length || !text || !/^\d{4}-\d{2}-\d{2}$/.test(due)) throw new Error('Хүмүүс, даалгавар, хугацаагаа оруулна уу');
  var vis = {}; visiblePeople_(me).forEach(function (u) { vis[u.name] = 1; });
  if (to.some(function (n) { return !vis[n]; })) throw new Error('Зарим хүнд даалгавар өгөх эрхгүй байна');
  var now = new Date(), batch = 'B' + Utilities.formatDate(now, 'Asia/Ulaanbaatar', 'yyMMddHHmmss'), p = due.split('-');
  var rows = to.map(function (n, i) { return [batch + '-' + (i + 1), now, me.name, n, text, new Date(+p[0], +p[1] - 1, +p[2]), String(t.proj || ''), TST.NEW, '', '', '', batch]; });
  var sh = taskSheet_(); sh.getRange(sh.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
}
function taskDelete_(me, batch) {
  var sh = taskSheet_(), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (String(v[i][11]) === batch && (String(v[i][2]) === me.name || isDir_(me))) sh.deleteRow(i + 1);
}
function taskUpd_(me, id, what, note) {
  var sh = taskSheet_(), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === id) {
    if (String(v[i][3]) !== me.name) throw new Error('Энэ даалгавар танд хамаарахгүй');
    if (what === 'ack') { if (String(v[i][7]) === TST.NEW) sh.getRange(i + 1, 8, 1, 2).setValues([[TST.ACK, new Date()]]); return; }
    if (what === 'notdone' && !String(note || '').trim()) throw new Error('Яагаад биелээгүйг бичнэ үү');
    sh.getRange(i + 1, 8).setValue(what === 'done' ? TST.DONE : TST.NOT);
    if (!v[i][8]) sh.getRange(i + 1, 9).setValue(new Date());
    sh.getRange(i + 1, 10, 1, 2).setValues([[new Date(), String(note || '').trim()]]);
    return;
  }
  throw new Error('Даалгавар олдсонгүй');
}
function saveWage_(w) {
  var u = users_().filter(function (x) { return x.name === w.name; })[0];
  if (!u) throw new Error('Хэрэглэгч олдсонгүй');
  sh_('Хэрэглэгч').getRange(u.row, 9, 1, 3).setValues([[w.wageType === 'Сар' ? 'Сар' : 'Өдөр', num_(w.wage), num_(w.wdays) || 22]]);
  dropUsers_();
}


/* ---------- Гал тогоо: тогооч хоол болсныг мэдэгдэж, хэдэн хүнд хоол хийснээ бүртгэнэ ----------
 * Хоол: A ID | B Хугацаа | C Тогооч | D Төсөл | E Хоолны цаг | F Цэс | G Хүний тоо | H Зураг
 */
var MEALS = ['Өглөөний цай', 'Өдрийн хоол', 'Оройн хоол', 'Шөнийн хоол'];
function mealSheet_() { return ensureSheet_('Хоол', ['ID','Хугацаа','Тогооч','Төсөл','Хоолны цаг','Цэс','Хүний тоо','Зураг']); }
function meals_() {
  return rowsOf_(mealSheet_(), 8).map(function (r) {
    return { id: String(r[0]), at: fmtDate_(r[1]), ts: r[1] instanceof Date ? r[1].getTime() : 0, by: String(r[2] || ''), proj: String(r[3] || ''), type: String(r[4] || ''), menu: String(r[5] || ''), count: num_(r[6]), photo: String(r[7] || '') };
  });
}
function mealData_(me) {
  var lim = Date.now() - 400 * 864e5;
  var list = meals_().filter(function (m) { return m.ts >= lim && (!isCook_(me) || m.by === me.name || canProj_(me, m.proj)); }).map(function (m) {
    return { id: m.id, at: m.at, by: m.by, proj: m.proj, type: m.type, menu: m.menu, count: m.count, hasPhoto: !!m.photo };
  }).reverse();
  return { today: ymd_(new Date()), meals: list, types: MEALS };
}
function mealAdd_(me, m, photo) {
  var count = Math.round(num_(m.count));
  if (MEALS.indexOf(m.type) < 0 || !String(m.menu || '').trim()) throw new Error('Хоолны цаг, цэсээ оруулна уу');
  if (!(count > 0) || count > 2000) throw new Error('Хэдэн хүнд хоол хийснээ оруулна уу');
  var now = new Date(), id = 'M' + Utilities.formatDate(now, 'Asia/Ulaanbaatar', 'yyMMddHHmmss');
  var photoId = (photo && photo.indexOf('data:image') === 0) ? savePhoto_('Аасгалдай хоол', id + '.jpg', photo) : '';
  mealSheet_().appendRow([id, now, me.name, String(m.proj || ''), m.type, String(m.menu).trim(), count, photoId]);
}
function mealDelete_(me, id) {
  var sh = mealSheet_(), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === id && (String(v[i][2]) === me.name || isDir_(me))) { sh.deleteRow(i + 1); return; }
}
function lastMeal_(me) {
  var lim = Date.now() - 3 * 3600e3, list = meals_().filter(function (m) { return m.ts >= lim && m.by !== me.name && (!m.proj || canProj_(me, m.proj)); });
  if (!list.length) return null;
  var m = list[list.length - 1];
  return { id: m.id, at: m.at, by: m.by, proj: m.proj, type: m.type, menu: m.menu, hasPhoto: !!m.photo };
}


/* ---------- Албан бичиг ----------
 * Албан бичиг: A Дугаар | B Он (2 орон) | C Бүтэн дугаар (123/26) | D Огноо | E Хаана | F Гарчиг | G Агуулга | H Танай огноо | I Танай дугаар
 *              J Тамгатай | K Гарын үсэг зурсан албан тушаал | L Нэр | M Гаргасан | N Хугацаа | O Төлөв
 * Тамга, гарын үсгийн зураг: Script Properties LETTER_STAMP (файлын ID). Тохиргоо: LETTER_CFG.
 * Эрх: Нягтлан, Админ, Захирал.
 */
var LCOLS = ['Дугаар','Он','Бүтэн дугаар','Огноо','Хаана','Гарчиг','Агуулга','Танай огноо','Танай дугаар','Тамгатай','Албан тушаал','Нэр','Гаргасан','Хугацаа','Төлөв'];
function letterSheet_() { return ensureSheet_('Албан бичиг', LCOLS); }
function letterCfg_() {
  var p = PropertiesService.getScriptProperties(), d = { signTitle: 'Гүйцэтгэх захирал', signName: 'Ү. Эрдэнэсанаа', start: {} };
  try { var c = JSON.parse(p.getProperty('LETTER_CFG') || '{}'); for (var k in c) d[k] = c[k]; } catch (e) {}
  d.hasStamp = !!p.getProperty('LETTER_STAMP');
  return d;
}
function letterCfgSave_(c, photo) {
  var p = PropertiesService.getScriptProperties(), cur = letterCfg_();
  var o = { signTitle: String(c.signTitle || cur.signTitle).trim(), signName: String(c.signName || cur.signName).trim(), start: cur.start || {} };
  if (c.startNo) { var yy = ymd_(new Date()).slice(2, 4), n = Math.round(num_(c.startNo)); if (n > 0) o.start[yy] = n; }
  p.setProperty('LETTER_CFG', JSON.stringify(o));
  if (photo && photo.indexOf('data:image/png') === 0) p.setProperty('LETTER_STAMP', savePhoto_('Аасгалдай тамга', 'tamga.png', photo));
}
function letters_() {
  return rowsOf_(letterSheet_(), 15).map(function (r) {
    return { n: num_(r[0]), yy: String(r[1]).length === 1 ? '0' + r[1] : String(r[1]), no: String(r[2]), date: ymd_(r[3]), to: String(r[4] || ''), title: String(r[5] || ''), body: String(r[6] || ''),
             refDate: String(r[7] || ''), refNo: String(r[8] || ''), stamp: r[9] === true || String(r[9]).toUpperCase() === 'TRUE',
             signTitle: String(r[10] || ''), signName: String(r[11] || ''), by: String(r[12] || ''), at: fmtDate_(r[13]), status: String(r[14] || '') };
  });
}
function letterNext_(yy) {
  var max = 0; letters_().forEach(function (l) { if (l.yy === yy) max = Math.max(max, l.n); });
  return Math.max(max + 1, num_((letterCfg_().start || {})[yy]) || 1);
}
function letterData_(me) {
  var t = ymd_(new Date()), yy = t.slice(2, 4);
  return { today: t, next: letterNext_(yy) + '/' + yy, cfg: letterCfg_(), letters: letters_().reverse().slice(0, 500), canCfg: isDir_(me) || me.role === 'Нягтлан' };
}
function letterSave_(me, l) {
  l = l || {};
  var to = String(l.to || '').trim(), title = String(l.title || '').trim(), body = String(l.body || '').replace(/\r/g, '').replace(/\s+$/, '').replace(/^(\s*\n)+/, ''), date = String(l.date || ymd_(new Date()));
  if (!to || !title || !body.trim()) throw new Error('Хаана, гарчиг, агуулгаа бичнэ үү');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Огноо буруу');
  var cfg = letterCfg_(), stamp = !!l.stamp;
  if (stamp && !cfg.hasStamp) throw new Error('Тамга, гарын үсгийн зураг оруулаагүй байна');
  var signTitle = String(l.signTitle || cfg.signTitle).trim(), signName = String(l.signName || cfg.signName).trim();
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = letterSheet_(), v = sh.getDataRange().getValues(), p = date.split('-'), d = new Date(+p[0], +p[1] - 1, +p[2]);
    if (l.no) {
      for (var i = 1; i < v.length; i++) if (String(v[i][2]) === String(l.no)) {
        if (String(v[i][14]) === 'Хүчингүй') throw new Error('Хүчингүй болсон албан бичгийг засах боломжгүй');
        sh.getRange(i + 1, 4, 1, 9).setValues([[d, to, title, body, String(l.refDate || ''), String(l.refNo || ''), stamp, signTitle, signName]]);
        return letters_().filter(function (x) { return x.no === String(l.no); })[0];
      }
      throw new Error('Албан бичиг олдсонгүй');
    }
    var yy = date.slice(2, 4), n = letterNext_(yy), no = n + '/' + yy;
    sh.appendRow([n, "'" + yy, no, d, to, title, body, String(l.refDate || ''), String(l.refNo || ''), stamp, signTitle, signName, me.name, new Date(), '']);
  } finally { lock.releaseLock(); }
  return letters_().filter(function (x) { return x.no === no; })[0];
}
function letterVoid_(me, no) {
  var sh = letterSheet_(), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][2]) === String(no)) { sh.getRange(i + 1, 15).setValue('Хүчингүй'); return; }
  throw new Error('Олдсонгүй');
}


/* ======================================================================
 * АЖЛЫН ГРАФИК (календарчилсан төлөвлөгөө, өдөр бүрийн гүйцэтгэл)
 * Ажлын график: A ID | B Төсөл | C Ажлын нэр | D Нэгж | E Тоо хэмжээ | F Эхлэх | G Дуусах | H Дараалал | I Бүртгэсэн | J Огноо
 * Гүйцэтгэл:    A ID | B Ажлын ID | C Төсөл | D Огноо | E Тоо хэмжээ | F Тэмдэглэл | G Бүртгэсэн | H Цаг
 * Харах: бүх нэвтэрсэн хэрэглэгч. Засах: Захирал, эсвэл тухайн төслийн "График хөтлөгч" Инженер.
 * ====================================================================== */
function schedSheet_() { return ensureSheet_('Ажлын график', ['ID','Төсөл','Ажлын нэр','Нэгж','Тоо хэмжээ','Эхлэх','Дуусах','Дараалал','Бүртгэсэн','Огноо']); }
function schedLogSheet_() { return ensureSheet_('Гүйцэтгэл', ['ID','Ажлын ID','Төсөл','Огноо','Тоо хэмжээ','Тэмдэглэл','Бүртгэсэн','Цаг']); }
function schedItems_() {
  return rowsOf_(schedSheet_(), 10).map(function (r) {
    return { id: String(r[0]), proj: String(r[1]), name: String(r[2] || ''), unit: String(r[3] || ''), qty: num_(r[4]), start: ymd_(r[5]), end: ymd_(r[6]), ord: num_(r[7]), by: String(r[8] || '') };
  });
}
function schedLogs_() {
  return rowsOf_(schedLogSheet_(), 8).map(function (r) {
    return { id: String(r[0]), item: String(r[1]), proj: String(r[2]), date: ymd_(r[3]), qty: num_(r[4]), note: String(r[5] || ''), by: String(r[6] || '') };
  });
}
function schedCan_(me, proj) { return isDir_(me) || (isEng_(me) && me.sched && canProj_(me, proj)); }
function schedDate_(x) { x = String(x || ''); if (!/^\d{4}-\d{2}-\d{2}$/.test(x)) return null; var p = x.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
function schedData_(me) {
  var us = activeUsers_();
  var projects = meta_().projects.filter(function (p) { return p[0] !== GEN; }).map(function (p) {
    var ed = us.filter(function (u) { return isEng_(u) && u.sched && canProj_(u, p[0]); })[0];
    return { code: p[0], name: p[1], editor: ed ? ed.name : '', editorPhone: ed ? ed.phone : '' };
  });
  return { today: ymd_(new Date()), projects: projects, items: schedItems_(), logs: schedLogs_(),
           canEdit: projects.filter(function (p) { return schedCan_(me, p.code); }).map(function (p) { return p.code; }) };
}
function schedItem_(me, it) {
  it = it || {};
  var proj = String(it.proj || '').trim();
  if (!schedCan_(me, proj)) throw new Error('Энэ төслийн графикийг засах эрхгүй');
  var name = String(it.name || '').trim(), unit = String(it.unit || '').trim(), qty = num_(it.qty), st = schedDate_(it.start), en = schedDate_(it.end);
  if (!name) throw new Error('Ажлын нэрийг бичнэ үү');
  if (!(qty > 0)) throw new Error('Тоо хэмжээг оруулна уу');
  if (!st || !en) throw new Error('Эхлэх, дуусах огноог оруулна уу');
  if (String(it.end) < String(it.start)) throw new Error('Дуусах огноо эхлэхээс өмнө байна');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = schedSheet_(), v = sh.getDataRange().getValues();
    if (it.id) {
      for (var i = 1; i < v.length; i++) if (String(v[i][0]) === String(it.id)) {
        if (String(v[i][1]) !== proj) throw new Error('Төсөл таарахгүй байна');
        sh.getRange(i + 1, 3, 1, 5).setValues([[name, unit, qty, st, en]]);
        if (it.ord != null && it.ord !== '') sh.getRange(i + 1, 8).setValue(num_(it.ord));
        return;
      }
      throw new Error('Ажил олдсонгүй');
    }
    var mx = 0; for (var j = 1; j < v.length; j++) if (String(v[j][1]) === proj) mx = Math.max(mx, num_(v[j][7]));
    sh.appendRow(['G' + new Date().getTime().toString(36) + Math.floor(Math.random() * 1296).toString(36), proj, name, unit, qty, st, en, mx + 1, me.name, new Date()]);
  } finally { lock.releaseLock(); }
}
function schedItemDel_(me, id) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = schedSheet_(), v = sh.getDataRange().getValues(), found = false;
    for (var i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === String(id)) {
      if (!schedCan_(me, String(v[i][1]))) throw new Error('Энэ төслийн графикийг засах эрхгүй');
      sh.deleteRow(i + 1); found = true;
    }
    if (!found) return;
    var lh = schedLogSheet_(), lv = lh.getDataRange().getValues();
    for (var k = lv.length - 1; k >= 1; k--) if (String(lv[k][1]) === String(id)) lh.deleteRow(k + 1);
  } finally { lock.releaseLock(); }
}
function schedDay_(me, proj, date, entries) {
  proj = String(proj || '');
  if (!schedCan_(me, proj)) throw new Error('Энэ төслийн графикийг засах эрхгүй');
  var d = schedDate_(date); if (!d) throw new Error('Огноо буруу');
  if (String(date) > ymd_(new Date())) throw new Error('Ирээдүйн өдрийн гүйцэтгэл оруулах боломжгүй');
  var items = {}; schedItems_().forEach(function (x) { if (x.proj === proj) items[x.id] = x; });
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = schedLogSheet_(), v = sh.getDataRange().getValues(), at = {};
    for (var i = 1; i < v.length; i++) if (v[i][0]) at[String(v[i][1]) + '|' + ymd_(v[i][3])] = i + 1;
    var del = [];
    (entries || []).forEach(function (e) {
      var it = items[String(e.item)]; if (!it) return;
      var q = e.qty === '' || e.qty == null ? 0 : num_(e.qty), note = String(e.note || '').trim(), row = at[it.id + '|' + date];
      if (q < 0) throw new Error('Тоо хэмжээ сөрөг байж болохгүй');
      if (!q && !note) { if (row) del.push(row); return; }
      if (row) sh.getRange(row, 5, 1, 4).setValues([[q, note, me.name, new Date()]]);
      else sh.appendRow(['L' + new Date().getTime().toString(36) + Math.floor(Math.random() * 46656).toString(36), it.id, proj, d, q, note, me.name, new Date()]);
    });
    del.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  } finally { lock.releaseLock(); }
}


/* ======================================================================
 * ЗЭЭЛ, УРЬДЧИЛГАА
 * Зээл:            A ID | B Төрөл (Олгосон/Урьдчилгаа/Авсан) | C Төсөл | D Харилцагч | E Дүн | F Огноо | G Сарын хүү % | H Дуусах огноо | I Суутгах сар | J Хүсэлт ID | K Бүртгэсэн | L Тэмдэглэл
 * Зээлийн гүйлгээ: A ID | B Зээл ID | C Огноо | D Төрөл | E Дүн | F Бүртгэсэн | G Хүсэлт ID | H Тэмдэглэл | I Сар (цалингийн суутгал)
 * Үлдэгдэл: Олгосон/Урьдчилгаа = дүн − (Эргэн төлөлт + Цалингийн суутгал); Авсан = дүн − Үндсэн төлбөр. Хүү нь зардал.
 * ====================================================================== */
var LTYPES = ['Олгосон', 'Урьдчилгаа', 'Авсан'];
var TXT = { BACK: 'Эргэн төлөлт', SAL: 'Цалингийн суутгал', PRIN: 'Үндсэн төлбөр', INT: 'Хүү, шимтгэл' };
function loanSheet_() { return ensureSheet_('Зээл', ['ID','Төрөл','Төсөл','Харилцагч','Дүн','Огноо','Сарын хүү %','Дуусах огноо','Суутгах сар','Хүсэлт ID','Бүртгэсэн','Тэмдэглэл']); }
function loanTxSheet_() { return ensureSheet_('Зээлийн гүйлгээ', ['ID','Зээл ID','Огноо','Төрөл','Дүн','Бүртгэсэн','Хүсэлт ID','Тэмдэглэл','Сар']); }
function loans_() {
  return rowsOf_(loanSheet_(), 12).map(function (r) {
    return { id: String(r[0]), type: String(r[1]), proj: String(r[2] || ''), party: String(r[3] || ''), amount: num_(r[4]), date: ymd_(r[5]), rate: num_(r[6]), due: ymd_(r[7]), months: num_(r[8]) || 1, req: String(r[9] || ''), by: String(r[10] || ''), note: String(r[11] || '') };
  });
}
function loanTxs_() {
  return rowsOf_(loanTxSheet_(), 9).map(function (r) {
    return { id: String(r[0]), loan: String(r[1]), date: ymd_(r[2]), type: String(r[3]), amount: num_(r[4]), by: String(r[5] || ''), req: String(r[6] || ''), note: String(r[7] || ''), ym: String(r[8] || '') };
  });
}
function loanList_() {
  var tx = loanTxs_();
  return loans_().map(function (l) {
    l.txs = tx.filter(function (t) { return t.loan === l.id; });
    var red = 0, intr = 0;
    l.txs.forEach(function (t) { if (t.type === TXT.INT) intr += t.amount; else if (l.type === 'Авсан' ? t.type === TXT.PRIN : (t.type === TXT.BACK || t.type === TXT.SAL)) red += t.amount; });
    l.paid = red; l.interest = intr; l.bal = Math.max(0, l.amount - red);
    return l;
  });
}
function loanData_() { return { today: ymd_(new Date()), loans: loanList_(), loanLimit: loanLimit_() }; }
function lid_(p) { return p + new Date().getTime().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase(); }
function loanAdd_(me, l) {
  var p = String(l.date).split('-');
  var id = lid_('Z');
  loanSheet_().appendRow([id, l.type, l.proj || GEN, String(l.party).trim(), num_(l.amount), new Date(+p[0], +p[1] - 1, +p[2]), num_(l.rate) || '', l.due ? schedDate_(l.due) : '', num_(l.months) || '', l.req || '', me.name, String(l.note || '').trim()]);
  return id;
}
function txAdd_(me, t) {
  var p = String(t.date).split('-');
  loanTxSheet_().appendRow([lid_('T'), t.loan, new Date(+p[0], +p[1] - 1, +p[2]), t.type, num_(t.amount), me.name, t.req || '', String(t.note || '').trim(), t.ym || '']);
}
function loanSave_(me, l) {
  l = l || {};
  if (LTYPES.indexOf(l.type) < 0) throw new Error('Зээлийн төрөл буруу');
  if (!String(l.party || '').trim()) throw new Error(l.type === 'Авсан' ? 'Банк / зээлдүүлэгчийн нэрийг бичнэ үү' : 'Зээлдэгчийн нэрийг бичнэ үү');
  if (!(num_(l.amount) > 0)) throw new Error('Дүнгээ оруулна уу');
  if (!schedDate_(l.date)) throw new Error('Огноо буруу');
  if (l.proj && l.proj !== GEN && !projects_().some(function (p) { return p.code === l.proj; })) throw new Error('Төсөл олдсонгүй');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { loanAdd_(me, l); } finally { lock.releaseLock(); }
}
function loanTx_(me, t) {
  t = t || {};
  var ln = loanList_().filter(function (x) { return x.id === t.loan; })[0];
  if (!ln) throw new Error('Зээл олдсонгүй');
  var ok = ln.type === 'Авсан' ? [TXT.PRIN, TXT.INT] : [TXT.BACK, TXT.SAL, TXT.INT];
  if (ok.indexOf(t.type) < 0) throw new Error('Гүйлгээний төрөл буруу');
  if (!(num_(t.amount) > 0)) throw new Error('Дүнгээ оруулна уу');
  if (!schedDate_(t.date)) throw new Error('Огноо буруу');
  if (t.type !== TXT.INT && num_(t.amount) > ln.bal + 0.5) throw new Error('Үлдэгдлээс (' + ln.bal.toLocaleString() + '₮) их дүн байна');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { txAdd_(me, t); } finally { lock.releaseLock(); }
}
function loanDel_(id, isTx) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (isTx) { var th = loanTxSheet_(), tv = th.getDataRange().getValues(); for (var i = tv.length - 1; i >= 1; i--) if (String(tv[i][0]) === id) { th.deleteRow(i + 1); return; } throw new Error('Олдсонгүй'); }
    var ln = loanList_().filter(function (x) { return x.id === id; })[0]; if (!ln) throw new Error('Олдсонгүй');
    if (ln.txs.length) throw new Error('Гүйлгээтэй зээлийг устгах боломжгүй. Эхлээд гүйлгээнүүдийг устгана уу.');
    var sh = loanSheet_(), v = sh.getDataRange().getValues();
    for (var k = v.length - 1; k >= 1; k--) if (String(v[k][0]) === id) sh.deleteRow(k + 1);
  } finally { lock.releaseLock(); }
}
function loanOnPaid_(me, r) {
  var x = {}; try { x = JSON.parse(r.link || '{}'); } catch (e) {}
  var d = ymd_(new Date());
  if (loans_().some(function (l) { return l.req === r.id; }) || loanTxs_().some(function (t) { return t.req === r.id; })) return;
  if (r.rtype === RT.ADV) loanAdd_(me, { type: 'Урьдчилгаа', proj: r.proj, party: x.emp || r.payee, amount: r.amount, date: d, months: x.months || 1, req: r.id, note: r.what });
  else if (r.rtype === RT.LOAN) loanAdd_(me, { type: 'Олгосон', proj: r.proj, party: r.payee, amount: r.amount, date: d, rate: x.rate, due: x.due, req: r.id, note: r.what });
  else if (x.loan) txAdd_(me, { loan: x.loan, date: d, type: r.rtype === RT.PRIN ? TXT.PRIN : TXT.INT, amount: r.amount, req: r.id, note: r.what });
}
function payDeduct_(me, ym, items) {
  if (!/^\d{4}-\d{2}$/.test(String(ym))) throw new Error('Сар буруу');
  var L = loanList_(), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = loanTxSheet_(), v = sh.getDataRange().getValues();
    (items || []).forEach(function (it) {
      var ln = L.filter(function (x) { return x.id === it.loan && x.type === 'Урьдчилгаа'; })[0]; if (!ln) return;
      for (var i = v.length - 1; i >= 1; i--) if (String(v[i][1]) === ln.id && String(v[i][3]) === TXT.SAL && String(v[i][8]) === ym) sh.deleteRow(i + 1);
      var a = Math.round(num_(it.amount));
      if (a > 0) txAdd_(me, { loan: ln.id, date: ym + '-' + (ym === ymd_(new Date()).slice(0, 7) ? ymd_(new Date()).slice(8) : '28'), type: TXT.SAL, amount: a, note: ym + ' сарын цалингаас', ym: ym });
      v = sh.getDataRange().getValues();
    });
  } finally { lock.releaseLock(); }
}


  // Зураг: Google Drive-ийн оронд Supabase Storage. ID нь хуучин Drive ID эсвэл "sb..." шинэ ID.
  function fileData_(id) { return id ? FILE_MARK + id : ""; }
  function savePhoto_(folderName, name, dataUrl) {
    const id = "sb" + randomUUID().replace(/-/g, "");
    const mime = dataUrl.substring(5, dataUrl.indexOf(";")) || "image/jpeg";
    ctx.uploads.push({ id, mime, b64: dataUrl.split(",")[1] || "" });
    return id;
  }
  return { handle_, auth_, users_, isBoss_, isDir_, isPM_, isHab_, isCook_, isHR_, isAdmin_, isMaster_, canProj_, canApprove_, approveLimit_, ymd_ };
}

/* ---------- хүсэлт боловсруулах ---------- */
export async function runRequest(body, store) {
  if (!body || typeof body !== "object") return { error: "bad", message: "Хүсэлт буруу" };
  if (typeof body.action === "string" && body.action.indexOf("__") === 0) return admin_(body, store);
  if (PUSH_ACTIONS.indexOf(body.action) >= 0) return pushAction_(body, store);
  for (let attempt = 0; attempt < 6; attempt++) {
    const snap = await store.snapshot();
    // Өгөгдөл шилжээгүй байхад хуучин Google Apps Script сервер рүү дамжуулна
    if (!snap["Хэрэглэгч"] && store.legacy) return await store.legacy(body);
    const ctx = makeCtx(snap);
    const app = buildApp(ctx);
    const out = app.handle_(body);
    if (ctx.book.dirty.size) {
      for (const u of ctx.uploads) await store.upload(u.id, u.mime, u.b64);
      const ok = await store.commit([...ctx.book.dirty].map((s) => ({ name: s.name, rows: s.rows, baseVer: s.entry ? s.entry.ver : 0 })));
      if (!ok) continue;
      try {
        const plan = notifyPlan_(app, body, out);
        if (plan.length) { const p = deliver_(ctx, plan, store).catch((e) => console.error("push", e)); await Promise.race([p, new Promise((r) => setTimeout(r, 6000))]); if (store.defer) store.defer(p); }
      } catch (e) { console.error("notify", e); }
    }
    if (out && typeof out.data === "string" && out.data.indexOf(FILE_MARK) === 0) {
      const f = await store.download(out.data.slice(FILE_MARK.length));
      out.data = f ? "data:" + f.mime + ";base64," + Buffer.from(f.bytes).toString("base64") : "";
    }
    return out;
  }
  return { error: "server", message: "Олон хүн зэрэг хадгалж байна. Дахин оролдоно уу." };
}

/* ---------- шилжүүлэлт, Google Sheet-тэй синк (зөвхөн нууц түлхүүртэй) ---------- */
async function admin_(b, store) {
  if (b.action === "__ping") return { ok: true, t: Date.now() };
  if (b.action === "__selftest") {
    const t0 = Date.now(), snap = await store.snapshot(), t1 = Date.now();
    await store.snapshot(); const t2 = Date.now();
    await store.upload("selftest.txt", "text/plain", Buffer.from("ok " + t0).toString("base64"));
    const f = await store.download("selftest.txt"); const t3 = Date.now();
    return { ok: !!f && Buffer.from(f.bytes).toString() === "ok " + t0, sheets: Object.keys(snap).length, dbMs: t1 - t0, db2Ms: t2 - t1, worker: WORKER, storageMs: t3 - t2 };
  }
  if (!b.key || sha256hex(b.key) !== KEY_HASH) return { error: "auth", message: "key" };
  switch (b.action) {
    case "__import": {
      // Apps Script-ийн SHA-256 нь энд тооцсонтой яг ижил эсэхийг шалгана (PIN-үүд хэвээр ажиллах баталгаа)
      if (!b.digest) return { error: "digest", message: "digest байхгүй" };
      const enc = Object.keys(ENCODERS).find((e) => hexOf_(digestBytes("sha256", b.digest.input, e)) === b.digest.hex);
      if (!enc) return { error: "digest", message: "SHA-256 кодлол тодорхойгүй: " + b.digest.hex + " " + JSON.stringify(b.digest.input) };
      const sheets = {};
      for (const n in b.sheets || {}) sheets[n] = decRows(b.sheets[n]);
      sheets.__props = Object.keys(b.props || {}).map((k) => [k, String(b.props[k])]).concat([["__DIGEST_ENC", enc]]);
      await store.replaceAll(sheets);
      const counts = {};
      for (const n in sheets) counts[n] = sheets[n].length;
      return { ok: true, counts, enc };
    }
    case "__fileList": return { ids: await store.listFiles() };
    case "__file": await store.upload(b.id, b.mime || "image/jpeg", b.b64); return { ok: true };
    case "__export": {
      const snap = await store.snapshot(), out = {};
      for (const n in snap) if (n.indexOf("__") !== 0) out[n] = encRows(snap[n].rows);
      return { sheets: out, t: Date.now() };
    }
    case "__put": {
      for (let attempt = 0; attempt < 6; attempt++) {
        const snap = await store.snapshot(), list = [];
        for (const n in b.sheets || {}) list.push({ name: n, rows: decRows(b.sheets[n]), baseVer: snap[n] ? snap[n].ver : 0 });
        if (!list.length || (await store.commit(list))) return { ok: true };
      }
      return { error: "server", message: "conflict" };
    }
    case "__mailPull": {
      for (let attempt = 0; attempt < 6; attempt++) {
        const snap = await store.snapshot(), m = snap.__mail;
        if (!m || !m.rows.length) return { mails: [] };
        if (await store.commit([{ name: "__mail", rows: [], baseVer: m.ver }])) return { mails: encRows(m.rows) };
      }
      return { mails: [] };
    }
  }
  return { error: "bad", message: "Unknown admin action" };
}
function hexOf_(bytes) { return bytes.map((x) => ("0" + (x & 255).toString(16)).slice(-2)).join(""); }

/* ---------- утасны мэдэгдэл (Web Push) ---------- */
const PUSH_ACTIONS = ["pushKey", "pushSub", "pushUnsub", "pushTest"];
const PUSH_SUBJECT = "https://app.aasgaldai.mn";
async function pushAction_(b, store) {
  for (let attempt = 0; attempt < 6; attempt++) {
    const snap = await store.snapshot(), ctx = makeCtx(snap), app = buildApp(ctx);
    const me = app.auth_(b);
    if (!me) return { error: "auth", message: "Дахин нэвтэрнэ үү" };
    const props = ctx.PropertiesService.getScriptProperties();
    let out;
    if (b.action === "pushKey") {
      let v = props.getProperty("__VAPID");
      if (!v) { v = JSON.stringify(await vapidGenerate()); props.setProperty("__VAPID", v); }
      out = { key: JSON.parse(v).pub };
    } else if (b.action === "pushSub" || b.action === "pushUnsub") {
      const sub = b.sub || {}, ep = String(sub.endpoint || "");
      if (!/^https:\/\//.test(ep)) return { error: "bad", message: "Буруу захиалга" };
      const sh = ctx.book.sys("__push"); sh._w();
      sh.rows = sh.rows.filter((r) => r && r[1] !== ep);
      if (b.action === "pushSub") {
        const k = sub.keys || {};
        if (!k.p256dh || !k.auth) return { error: "bad", message: "Түлхүүр дутуу" };
        sh.rows.push([me.name, ep, String(k.p256dh), String(k.auth), Date.now()]);
      }
      out = { ok: true };
    } else if (b.action === "pushTest") {
      const n = await deliver_(ctx, [{ to: [me.name], title: "🔔 Аасгалдай", body: "Мэдэгдэл ажиллаж байна. Сайн байна уу, " + me.name + "!", tag: "test" }], store);
      const subs = (ctx.book.getSheetByName("__push") || { rows: [] }).rows.filter((r) => r && r[0] === me.name).length;
      return { sent: n, devices: subs, results: (deliver_.last || []).map((x) => x[1] + " " + x[2] + (x[3] ? " " + x[3] : "")) };
    }
    if (ctx.book.dirty.size && !(await store.commit([...ctx.book.dirty].map((s) => ({ name: s.name, rows: s.rows, baseVer: s.entry ? s.entry.ver : 0 }))))) continue;
    return out;
  }
  return { error: "server", message: "Дахин оролдоно уу" };
}
function notifyPlan_(app, b, out) {
  if (!out || out.error) return [];
  if (b.action === "register") {
    const U0 = app.users_(), nm = out.name;
    return [{ to: U0.filter((u) => u.status !== "Хүлээгдэж буй" && (app.isHR_(u) || app.isDir_(u) || app.isAdmin_(u))).map((u) => u.name), title: "🆕 Шинэ бүртгэл", body: nm + " апп-д бүртгүүллээ. Тохиргоо → Хэрэглэгчид хэсгээс баталгаажуулна уу.", tag: "reg" }];
  }
  const me = app.auth_(b);
  if (!me) return [];
  const U = app.users_().filter((u) => u.status !== "Хүлээгдэж буй");
  const names = (f) => U.filter((u) => u.name !== me.name && f(u)).map((u) => u.name);
  const money = (n) => Math.round(Number(n) || 0).toLocaleString("en-US") + "₮";
  const cut = (t, n) => { t = String(t || "").replace(/\s+/g, " ").trim(); return t.length > n ? t.slice(0, n - 1) + "…" : t; };
  switch (b.action) {
    case "create": {
      const r = out.req; if (!r) return [];
      const big = Number(r.amount) > app.approveLimit_();
      return [{ to: names((u) => app.canApprove_(u, r) && (app.isDir_(u) || app.canProj_(u, r.proj))), title: (big ? "💰 Батлах (захирал) · " : "💰 Батлах · ") + r.proj, body: me.name + ": " + cut(r.what, 60) + " — " + money(r.amount), tag: "req-" + r.id }];
    }
    case "status": {
      const r = out.req; if (!r) return [];
      const em = { "Зөвшөөрсөн": "✅", "Төлсөн": "💸", "Буцаасан": "↩️" }[b.status] || "";
      const plan = r.by === me.name ? [] : [{ to: [r.by], title: em + " Хүсэлт: " + b.status, body: cut(r.what, 60) + " — " + money(r.amount) + (b.note ? " · " + cut(b.note, 80) : ""), tag: "req-" + r.id }];
      if (b.status === "Зөвшөөрсөн") plan.push({ to: names((u) => u.role === "Нягтлан" && app.canProj_(u, r.proj)), title: "💳 Төлөхөд бэлэн · " + r.proj, body: cut(r.what, 60) + " — " + money(r.amount) + " · баталсан " + me.name, tag: "pay-" + r.id });
      return plan;
    }
    case "taskSend":
      return [{ to: ((b.t && b.t.to) || []).filter((n) => n !== me.name), title: "📋 Шинэ даалгавар", body: me.name + ": " + cut(b.t.text, 100) + " · хугацаа " + b.t.due, tag: "task" }];
    case "leaveReq":
      return [{ to: names((u) => app.isPM_(u) || u.role === "Нягтлан" || app.isDir_(u)), title: "🗓 Чөлөөний хүсэлт", body: me.name + ": " + b.lv.from + " — " + b.lv.to + " · " + cut(b.lv.reason, 60), tag: "leave" }];
    case "leaveDecide": {
      const l = (out.leaves || []).filter((x) => x.id === b.id)[0]; if (!l || l.by === me.name) return [];
      return [{ to: [l.by], title: b.ok ? "✅ Чөлөө зөвшөөрөгдлөө" : "❌ Чөлөөний хүсэлт татгалзлаа", body: l.from + (l.to !== l.from ? " — " + l.to : "") + (b.note ? " · " + cut(b.note, 80) : ""), tag: "leave" }];
    }
    case "habReport":
      return [{ to: names((u) => app.isHab_(u) || app.isDir_(u)), title: "⚠️ ХАБЭА: " + b.rep.type, body: me.name + ": " + cut(b.rep.desc, 100), tag: "hab" }];
    case "habSave": {
      const n = b.note || {};
      if (n.id || String(n.date) !== app.ymd_(new Date())) return [];
      return [{ to: names((u) => !app.isHab_(u) && (!n.proj || app.canProj_(u, n.proj))), title: "🦺 Өнөөдрийн ХАБ заавар", body: cut(n.title || n.text, 100), tag: "habnote" }];
    }
    case "mealAdd": {
      const m = b.meal || {};
      return [{ to: names((u) => !m.proj || app.canProj_(u, m.proj)), title: "🍲 " + m.type + " бэлэн боллоо!", body: cut(m.menu, 90) + " · тогооч " + me.name, tag: "meal" }];
    }
  }
  return [];
}
// Төлөвлөгөөг холбогдох хүмүүсийн бүх төхөөрөмж рүү илгээнэ. Буцаах: амжилттай илгээсэн тоо
async function deliver_(ctx, plan, store) {
  const v = ctx.PropertiesService.getScriptProperties().getProperty("__VAPID");
  const sh = ctx.book.getSheetByName("__push");
  if (!v || !sh) return 0;
  const vapid = JSON.parse(v), jobs = [], dead = [], log = [];
  for (const p of plan) {
    const to = {};
    p.to.forEach((n) => (to[n] = 1));
    sh.rows.forEach((r) => {
      if (!r || !to[r[0]]) return;
      const host = (() => { try { return new URL(r[1]).host; } catch (_) { return "?"; } })();
      jobs.push(sendPush({ endpoint: r[1], p256dh: r[2], auth: r[3] }, vapid, { title: p.title, body: p.body, tag: p.tag || "", url: "./" }, PUSH_SUBJECT)
        .then((res) => { if (res.status === 404 || res.status === 410) dead.push(r[1]); log.push([r[0], host, res.status, res.text, Date.now()]); return res.status >= 200 && res.status < 300 ? 1 : 0; })
        .catch((e) => { log.push([r[0], host, 0, String(e && e.message || e).slice(0, 200), Date.now()]); return 0; }));
    });
  }
  const sent = (await Promise.all(jobs)).reduce((a, x) => a + x, 0);
  try {
    const snap = await store.snapshot(), e = snap.__push, lg = snap.__pushlog, list = [];
    if (dead.length && e) list.push({ name: "__push", rows: e.rows.filter((r) => dead.indexOf(r[1]) < 0), baseVer: e.ver });
    if (log.length) list.push({ name: "__pushlog", rows: (lg ? lg.rows : []).concat(log).slice(-40), baseVer: lg ? lg.ver : 0 });
    if (list.length) await store.commit(list);
  } catch (_) {}
  deliver_.last = log;
  return sent;
}
