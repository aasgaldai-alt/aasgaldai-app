// @ts-nocheck
// Аасгалдай API — Google Apps Script логикийг (Code.gs) Supabase дээр ажиллуулах орчин.
// Энэ файлыг build_api.py үүсгэдэг: Code.gs-ийг доорх __CODE__ хэсэгт хуулна. Гараар бүү засаарай.
import { createHash, randomUUID } from "node:crypto";
import { Buffer } from "node:buffer";

export const KEY_HASH = "f5e809095950b7176a2a97a76a6dcb3130c015274c6a1e585d736e8d3bc2b01d";
const TZ = "Asia/Ulaanbaatar";
const APP_URL = "https://app.aasgaldai.mn/";
const FILE_MARK = "\u0001FILE:";

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
    computeDigest(alg, s) { return Array.from(createHash(alg).update(String(s), "utf8").digest()).map((b) => (b > 127 ? b - 256 : b)); },
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
var ROLES = ['Ажилчин', 'Тогооч', 'ХАБ инженер', 'Төслийн удирдагч', 'Хүний нөөц', 'Нягтлан', 'Админ', 'Захирал'];
var VATS = ['НӨАТ-тай', 'НӨАТ-гүй'];
function isDir_(u) { return u.role === 'Захирал' || u.role === 'Админ'; }
function isCook_(u) { return u.role === 'Тогооч'; }
function isPM_(u) { return u.role === 'Төслийн удирдагч'; }
function isHR_(u) { return u.role === 'Хүний нөөц'; }
function isHab_(u) { return u.role === 'ХАБ инженер'; }
function isBoss_(u) { return u.role === 'Нягтлан' || isDir_(u); }
var COLS = ['id','date','by','proj','cat','what','qty','payee','bank','acct','amount','photo','status','decided','paid','desc','note','vat','ebarimt','phone','loc'];
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
    var me = auth_(b);
    if (!me) return { error: 'auth', message: 'Дахин нэвтэрнэ үү' };
    if (b.action === 'changePin') return changePin_(me, b.oldPin, b.newPin);
    if (me.mustChange) return { error: 'mustchange', message: 'Эхлээд шинэ PIN тохируулна уу' };
    var boss = isBoss_(me);
    var director = isDir_(me);
    switch (b.action) {
      case 'meta': return meta_();
      case 'people':
        if (!boss && !isHR_(me) && !isPM_(me)) return deny_();
        return peopleData_(me);
      case 'markP':
        if (!director && !isHR_(me) && !isPM_(me)) return deny_();
        markP_(me, b.a); return peopleData_(me);
      case 'my': return myData_(me);
      case 'meals':
        if (!isCook_(me) && !boss && !isHR_(me) && !isPM_(me)) return deny_();
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
        if (!director && !isPM_(me)) return deny_();
        taskSend_(me, b.t); return peopleData_(me);
      case 'taskDelete':
        if (!director && !isPM_(me)) return deny_();
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
        var mine = boss ? allRows : allRows.filter(function (r) { return r.by === me.name; });
        var out = { reqs: mine.slice(-500), meta: meta_() };
        if (boss) out.projects = dashboard_(allRows);
        try { out.hab = habBadge_(me); } catch (e) {}
        out.myProjects = me.projects;
        try { out.hr = hrBadge_(me); } catch (e) {}
        try { out.meal = lastMeal_(me); } catch (e) {}
        return out;
      case 'list':
        var all = rows_();
        if (!boss) all = all.filter(function (r) { return r.by === me.name; }); // ажилчин зөвхөн өөрийнхийг
        return { reqs: all.slice(-500) };
      case 'create': return { req: create_(me, b.req, b.photo) };
      case 'status':
        if (!boss) return deny_();
        return { req: setStatus_(me, b.id, b.status, b.note) };
      case 'photo':
        var r = rows_().filter(function (x) { return x.id === b.id; })[0];
        if (!r || !r.photo || (!boss && r.by !== me.name)) return { data: '' };
        return { data: fileData_(r.photo) };
      case 'dashboard':
        if (!boss) return deny_();
        return { projects: dashboard_() };
      case 'contract':
        if (!boss) return deny_();
        var p = projects_().filter(function (x) { return x.code === b.code; })[0];
        return { data: p && p.photo ? fileData_(p.photo) : '' };
      case 'saveProject':
        if (!director) return deny_();
        return { project: saveProject_(b.project, b.photo) };
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
        if (!boss && !isHR_(me)) return deny_();
        return { users: pubUsers_(boss) };
      case 'saveUser':
        if (!boss && !isHR_(me)) return deny_();
        saveUser_(me, b.user, b.isNew); return { users: pubUsers_(boss) };
      case 'deleteUser':
        if (!boss && !isHR_(me)) return deny_();
        deleteUser_(me, b.target); return { users: pubUsers_(boss) };
      case 'fuel': case 'saveEquip': case 'deleteEquip': case 'equipPhoto': case 'addFuel': case 'mark': case 'deleteFuel': case 'fuelPhoto':
        if (!boss) return deny_();
        break;
    }
    switch (b.action) {
      case 'fuel': return fuelData_(me, boss);
      case 'saveEquip':
        if (!boss) return deny_();
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
        if (!boss) return deny_();
        deleteFuel_(b.id); return fuelData_(me, boss);
      case 'fuelPhoto':
        var fr = fuelsRaw_().filter(function (x) { return x.id === b.id; })[0];
        if (!fr || !fr.photo || (!boss && fr.by !== me.name)) return { data: '' };
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
  if (sh.getMaxColumns() < 11) sh.insertColumnsAfter(sh.getMaxColumns(), 11 - sh.getMaxColumns());
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();
  var out = [];
  v.forEach(function (r, i) {
    if (!r[0]) return;
    out.push({ row: i + 2, name: String(r[0]).trim(), role: String(r[1]).trim(), pin: String(r[2]).trim(), email: String(r[3] || '').trim(),
               phone: digits_(r[4]), mustChange: r[5] === true || String(r[5]).toUpperCase() === 'TRUE', projects: projList_(r[6]),
               position: String(r[7] || '').trim(), wageType: String(r[8]) === 'Сар' ? 'Сар' : 'Өдөр', wage: num_(r[9]), wdays: num_(r[10]) || 22 });
  });
  cache.put('users', JSON.stringify(out), 300);
  return out;
}
function projList_(x) { var p = String(x || '').trim(); if (!p || p === '*') return null; return p.split(/[,\s]+/).filter(function (c) { return c; }); }
function canProj_(u, code) { return !u.projects || u.projects.indexOf(code) >= 0; }
function dropUsers_() { CacheService.getScriptCache().remove('users'); }
function pubUsers_(full) { return users_().map(function (u) { var o = { name: u.name, role: u.role, email: u.email, phone: u.phone, mustChange: u.mustChange, projects: u.projects, position: u.position };
  if (full) { o.wageType = u.wageType; o.wage = u.wage; o.wdays = u.wdays; } return o; }); }
function digits_(x) { return String(x === undefined || x === null ? '' : x).replace(/\D/g, ''); }
function hex_(bytes) { return bytes.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join(''); }
function secret_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function hashPin_(name, pin) { return 'h1:' + hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret_() + '|' + name + '|' + pin)); }
function pinOk_(u, pin) { return u.pin.indexOf('h1:') === 0 ? u.pin === hashPin_(u.name, pin) : (u.pin !== '' && u.pin === pin); }
function minLen_(role) { return role === 'Ажилчин' || role === 'Тогооч' ? 6 : 8; }
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
  if (u.pin.indexOf('h1:') !== 0) { setPin_(u, pin, true); u = users_().filter(function (x) { return x.name === u.name; })[0]; } // хуучин ил PIN-ийг нууцлах
  return { user: { name: u.name, role: u.role, token: token_(u), mustChange: u.mustChange } };
}
function auth_(b) {
  if (!b.name || !b.token) return null;
  var u = users_().filter(function (x) { return x.name === b.name; })[0];
  return u && token_(u) === b.token ? u : null;
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
  if (name === me.name && role !== me.role) throw new Error('Өөрийн эрхийг өөрчлөх боломжгүй');
  if (isNew && !tempPin) throw new Error('Түр PIN үүсээгүй байна');
  if (tempPin) checkPin_(role, tempPin);
  var all = users_(), ex = all.filter(function (x) { return x.name === name; })[0];
  if (isNew && ex) throw new Error('Ийм нэртэй хэрэглэгч байна');
  if (!isNew && !ex) throw new Error('Хэрэглэгч олдсонгүй');
  if (all.some(function (x) { return x.phone === phone && x.name !== name; })) throw new Error('Энэ утасны дугаар өөр хэрэглэгчид бүртгэлтэй байна');
  if (ex && ex.role === 'Захирал' && !isDir_(me)) throw new Error('Захирлын мэдээллийг засах эрхгүй');
  if (ex && isHR_(me) && (ex.role === 'Нягтлан' || isDir_(ex))) throw new Error('Энэ хэрэглэгчийг засах эрхгүй');
  if (ex && ex.name === me.name && tempPin) throw new Error('Өөрийн PIN-ийг "Миний PIN солих" хэсгээс солино уу');
  var sh = sh_('Хэрэглэгч');
  var pinCell = tempPin ? hashPin_(name, tempPin) : ex.pin, must = tempPin ? true : ex.mustChange;
  var pr = u.projects === '*' || u.projects === null ? '*' : (u.projects || []).map(function (c) { return String(c).trim(); }).filter(function (c) { return c; }).join(',');
  if (!pr) throw new Error('Дор хаяж нэг төсөл сонгоно уу');
  var row = [name, role, pinCell, email, "'" + phone, must, pr, String(u.position || '').trim()];
  if (!String(sh.getRange(1, 7).getValue())) sh.getRange(1, 7, 1, 5).setValues([['Төслүүд', 'Албан тушаал', 'Цалингийн төрөл', 'Цалин', 'Сарын ажлын өдөр']]);
  if (isNew) sh.appendRow(row); else sh.getRange(ex.row, 1, 1, 8).setValues([row]);
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
  var res = {
    projects: projects_().filter(function (p) { return p.active; }).map(function (p) { return [p.code, p.name]; }),
    categories: c.map(function (r) { return [r[0], r[1]]; }),
    banks: b.map(function (r) { return r[0]; }),
    locRequired: locRequired_()
  };
  cache.put('meta', JSON.stringify(res), 600);
  return res;
}
function locRequired_() { return PropertiesService.getScriptProperties().getProperty('LOC_REQUIRED') !== 'false'; }
function locStr_(l) {
  if (!l) return '';
  var lat = Number(l.lat), lng = Number(l.lng), acc = Math.round(Number(l.acc) || 0);
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180 || (!lat && !lng)) return '';
  return lat.toFixed(6) + ',' + lng.toFixed(6) + ',' + acc;
}
function dashboard_(rowsOpt) {
  var reqs = rowsOpt || rows_(), out = projects_(), incomes = incomesRaw_();
  out.forEach(function (p) {
    var s = { ok: { n: 0, sum: 0 }, paid: { n: 0, sum: 0 }, wait: { n: 0, sum: 0 }, no: { n: 0, sum: 0 }, cats: {}, vatSum: 0, vatN: 0, eboSum: 0, eboN: 0 };
    reqs.forEach(function (r) {
      if (r.proj !== p.code) return;
      var k = r.status === ST.OK ? 'ok' : r.status === ST.PAID ? 'paid' : r.status === ST.WAIT ? 'wait' : r.status === ST.NO ? 'no' : null;
      if (!k) return;
      s[k].n++; s[k].sum += r.amount;
      if (k === 'ok' || k === 'paid') { s.cats[r.cat] = (s.cats[r.cat] || 0) + r.amount; if (r.vat === 'НӨАТ-тай') { s.vatSum += r.amount; s.vatN++; } else if (r.ebarimt === 'Тийм') { s.eboSum += r.amount; s.eboN++; } }
    });
    p.stats = s; p.hasPhoto = !!p.photo; delete p.photo;
    p.incomes = incomes.filter(function (x) { return x.proj === p.code; }).map(function (x) {
      return { id: x.id, date: x.date, desc: x.desc, amount: x.amount, vat: x.vat, by: x.by, note: x.note, hasPhoto: !!x.photo };
    });
    p.incomeSum = p.incomes.reduce(function (a, x) { return a + x.amount; }, 0);
  });
  return out;
}
function saveProject_(p, photo) {
  CacheService.getScriptCache().remove('meta');
  var code = String(p.code || '').trim().toUpperCase(), name = String(p.name || '').trim();
  if (!code || !name) throw new Error('Төслийн код, нэрийг оруулна уу');
  var sh = sh_('Төсөл'), v = sh.getDataRange().getValues(), idx = -1;
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() === code) idx = i;
  if (p.isNew && idx > 0) throw new Error('Ийм кодтой төсөл байна');
  var photoId = idx > 0 ? String(v[idx][10] || '') : '';
  if (photo && photo.indexOf('data:image') === 0) photoId = savePhoto_(CONTRACT_FOLDER, code + '_geree.jpg', photo);
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
  if (!f.proj || !f.cat || !f.what || !(amount > 0) || !f.payee || !f.bank || !f.acct || VATS.indexOf(f.vat) < 0 || String(f.phone || '').replace(/\D/g, '').length < 8) throw new Error('Мэдээлэл дутуу байна');
  if (!canProj_(me, f.proj)) throw new Error('Танд ' + f.proj + ' төсөлд хүсэлт илгээх эрх алга. Нягтланд хандана уу.');
  var loc = locStr_(f.loc);
  if (!loc && !isBoss_(me) && locRequired_()) throw new Error('Байршил тогтоогдсонгүй. Утасны байршлаа (Location) асаагаад дахин илгээнэ үү.');
  var acct = String(f.acct).replace(/\D/g, ''), row;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sh_('Хүсэлт'), now = new Date();
    var id = 'R' + Utilities.formatDate(now, 'Asia/Ulaanbaatar', 'yyMMdd') + '-' + sh.getLastRow();
    var photoId = (photo && photo.indexOf('data:image') === 0) ? savePhoto_(PHOTO_FOLDER, id + '.jpg', photo) : '';
    row = [id, now, me.name, f.proj, f.cat, String(f.what).trim(), f.qty || '', String(f.payee).trim(), f.bank, "'" + acct, amount,
           photoId, ST.WAIT, '', '', mkDesc_(f), f.note || '', f.vat, (f.vat === 'НӨАТ-тай' || f.ebarimt) ? 'Тийм' : 'Үгүй', "'" + String(f.phone).replace(/[^\d+]/g, ''), loc];
    if (sh.getMaxColumns() < 21) sh.insertColumnsAfter(sh.getMaxColumns(), 21 - sh.getMaxColumns());
    if (!String(sh.getRange(1, 21).getValue())) sh.getRange(1, 21).setValue('Байршил');
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
      var cur = v[i][12];
      if (status === ST.PAID && cur !== ST.OK) throw new Error('Эхлээд зөвшөөрөх шаардлагатай');
      if ((status === ST.OK || status === ST.NO) && cur !== ST.WAIT) throw new Error('Энэ хүсэлт аль хэдийн шийдвэрлэгдсэн');
      sh.getRange(i + 1, 13).setValue(status);
      sh.getRange(i + 1, 14).setValue(me.name);
      if (status === ST.PAID) sh.getRange(i + 1, 15).setValue(new Date());
      if (note) sh.getRange(i + 1, 17).setValue(note);
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
    if (!boss) entries = entries.filter(function (f) { return f.by === me.name; });
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
  if (manager) out.people = users_().map(function (u) { return u.name; });
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
  var all = users_();
  if (!isPM_(me) || !me.projects) return all;
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
    canMark: isDir_(me) || isHR_(me) || isPM_(me), canTask: isDir_(me) || isPM_(me), canLeave: isBoss_(me) || isPM_(me), payroll: payroll
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
    rating: rating_(tasks, me.name)
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


  // Зураг: Google Drive-ийн оронд Supabase Storage. ID нь хуучин Drive ID эсвэл "sb..." шинэ ID.
  function fileData_(id) { return id ? FILE_MARK + id : ""; }
  function savePhoto_(folderName, name, dataUrl) {
    const id = "sb" + randomUUID().replace(/-/g, "");
    const mime = dataUrl.substring(5, dataUrl.indexOf(";")) || "image/jpeg";
    ctx.uploads.push({ id, mime, b64: dataUrl.split(",")[1] || "" });
    return id;
  }
  return { handle_ };
}

/* ---------- хүсэлт боловсруулах ---------- */
export async function runRequest(body, store) {
  if (!body || typeof body !== "object") return { error: "bad", message: "Хүсэлт буруу" };
  if (typeof body.action === "string" && body.action.indexOf("__") === 0) return admin_(body, store);
  for (let attempt = 0; attempt < 6; attempt++) {
    const snap = await store.snapshot();
    const ctx = makeCtx(snap);
    const out = buildApp(ctx).handle_(body);
    if (ctx.book.dirty.size) {
      for (const u of ctx.uploads) await store.upload(u.id, u.mime, u.b64);
      const ok = await store.commit([...ctx.book.dirty].map((s) => ({ name: s.name, rows: s.rows, baseVer: s.entry ? s.entry.ver : 0 })));
      if (!ok) continue;
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
    await store.upload("selftest.txt", "text/plain", Buffer.from("ok " + t0).toString("base64"));
    const f = await store.download("selftest.txt");
    return { ok: !!f && Buffer.from(f.bytes).toString() === "ok " + t0, sheets: Object.keys(snap).length, dbMs: t1 - t0, storageMs: Date.now() - t1 };
  }
  if (!b.key || sha256hex(b.key) !== KEY_HASH) return { error: "auth", message: "key" };
  switch (b.action) {
    case "__import": {
      // Apps Script-ийн SHA-256 нь энд тооцсонтой яг ижил эсэхийг шалгана (PIN-үүд хэвээр ажиллах баталгаа)
      if (!b.digest || hexOf_(makeCtx({}).Utilities.computeDigest("sha256", b.digest.input)) !== b.digest.hex) return { error: "digest", message: "SHA-256 зөрж байна" };
      const sheets = {};
      for (const n in b.sheets || {}) sheets[n] = decRows(b.sheets[n]);
      sheets.__props = Object.keys(b.props || {}).map((k) => [k, String(b.props[k])]);
      await store.replaceAll(sheets);
      const counts = {};
      for (const n in sheets) counts[n] = sheets[n].length;
      return { ok: true, counts };
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
