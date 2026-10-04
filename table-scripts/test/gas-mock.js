// Небольшая имитация Google Apps Script (SpreadsheetApp и друзья) для проверки Код.gs в Node.
// Хранит значения, rich text, проверки данных, правила цветов; форматирование — пустые операции.
'use strict';

const calls = { mail: [], toasts: [], alerts: [], menus: [], created: [], triggers: [] };

function colToNum(s) { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }

// ---------- Rich text ----------
class RichText {
  constructor(text, links, styles) { this.text = text; this.links = links || []; this.styles = styles || []; }
  getText() { return this.text; }
  getLinkUrl() {
    if (this.links.length === 1 && this.links[0].start === 0 && this.links[0].end === this.text.length) return this.links[0].url;
    return null;
  }
  getRuns() {
    const cuts = new Set([0, this.text.length]);
    this.links.forEach(l => { cuts.add(l.start); cuts.add(l.end); });
    const pts = [...cuts].sort((a, b) => a - b);
    const runs = [];
    for (let i = 0; i < pts.length - 1; i++) {
      const s = pts[i], e = pts[i + 1];
      const l = this.links.find(x => x.start <= s && x.end >= e);
      runs.push({ getLinkUrl: () => (l ? l.url : null), getText: () => this.text.slice(s, e) });
    }
    return runs;
  }
}
class RichBuilder {
  constructor() { this.text = ''; this.links = []; this.styles = []; }
  setText(t) { this.text = String(t); return this; }
  setLinkUrl(a, b, c) {
    if (c === undefined) this.links = [{ start: 0, end: this.text.length, url: a }];
    else {
      if (a < 0 || b > this.text.length || a >= b) throw new Error(`setLinkUrl: bad range ${a}-${b} for "${this.text}"`);
      this.links.push({ start: a, end: b, url: c });
    }
    return this;
  }
  setTextStyle(a, b, st) {
    if (b === undefined) return this;
    if (a < 0 || b > this.text.length || a > b) throw new Error(`setTextStyle: bad range ${a}-${b} for "${this.text}"`);
    this.styles.push({ a, b, st }); return this;
  }
  build() { return new RichText(this.text, this.links, this.styles); }
}

// Цепочка без эффекта: любые методы оформления возвращают сам объект
function chain(obj) {
  return new Proxy(obj, {
    get(t, p) {
      if (p in t) return t[p];
      if (typeof p === 'symbol') return undefined;
      return () => chain(t);
    }
  });
}

class Validation { constructor(kind, arg) { this.kind = kind; this.arg = arg; } }
class ValidationBuilder {
  requireValueInRange(r) { this.v = new Validation('range', r.a1()); return this; }
  requireValueInList(list) { this.v = new Validation('list', list); return this; }
  requireDate() { this.v = new Validation('date'); return this; }
  setAllowInvalid() { return this; }
  build() { return this.v; }
}
class CfBuilder {
  constructor() { this.rule = {}; }
  whenTextEqualTo(v) { this.rule.eq = v; return this; }
  whenTextStartsWith(v) { this.rule.starts = v; return this; }
  whenFormulaSatisfied(f) { this.rule.formula = f; return this; }
  setBackground(c) { this.rule.bg = c; return this; }
  setFontColor(c) { this.rule.fg = c; return this; }
  setBold() { return this; }
  setRanges(r) { this.rule.ranges = r.map(x => x.a1()); return this; }
  build() { return this.rule; }
}

// ---------- Sheet / Range ----------
class Sheet {
  constructor(ss, name, rows, maxRows, maxCols) {
    this.ss = ss; this.name = name;
    this.maxRows = Math.max(maxRows || 1000, rows.length);
    this.maxCols = Math.max(maxCols || 26, ...rows.map(r => r.length), 1);
    this.cells = rows.map(r => r.map(v => ({ v })));
    this.cf = []; this.filter = null; this.frozenRows = 0; this.frozenCols = 0; this.groups = {}; this.merges = [];
    this.mergeRects = []; // [row, col, numRows, numCols]
  }
  getSheetId() { return this.ss.sheets.indexOf(this) + 100; }
  cell(r, c) {
    while (this.cells.length < r) this.cells.push([]);
    const row = this.cells[r - 1];
    while (row.length < c) row.push({ v: '' });
    return row[c - 1];
  }
  peek(r, c) { const row = this.cells[r - 1]; return row && row[c - 1] ? row[c - 1] : { v: '' }; }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  getLastRow() {
    for (let r = this.cells.length; r >= 1; r--) if (this.cells[r - 1].some(c => c.v !== '' && c.v !== null && c.v !== undefined)) return r;
    return 0;
  }
  getLastColumn() { let m = 0; this.cells.forEach(row => row.forEach((c, i) => { if (c.v !== '' && c.v != null) m = Math.max(m, i + 1); })); return m; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') return this.rangeA1(a);
    return new Range(this, a, b, c || 1, d || 1);
  }
  rangeA1(a1) {
    const m = a1.match(/^([A-Z]+)(\d*)(?::([A-Z]+)(\d*))?$/);
    if (!m) throw new Error('bad a1 ' + a1);
    const c1 = colToNum(m[1]), r1 = m[2] ? Number(m[2]) : 1;
    const c2 = m[3] ? colToNum(m[3]) : c1;
    const r2 = m[3] ? (m[4] ? Number(m[4]) : this.maxRows) : r1;
    return new Range(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  }
  appendRow(values) { const r = this.getLastRow() + 1; values.forEach((v, i) => { this.cell(r, i + 1).v = v; }); return this; }
  getDataRange() { return new Range(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  insertRowAfter(r) {
    // Как в Таблицах: новая строка наследует оформление строки выше (здесь — пустая)
    this.cells.splice(r, 0, []); this.maxRows++;
    this.cf.forEach(rule => { rule.ranges = rule.ranges.map(x => x); });
    return this;
  }
  deleteRow(r) { this.cells.splice(r - 1, 1); this.maxRows--; return this; }
  deleteRows(r, n) { this.cells.splice(r - 1, n); this.maxRows -= n; return this; }
  insertRowsAfter(r, n) { for (let i = 0; i < n; i++) this.cells.splice(r, 0, []); this.maxRows += n; return this; }
  insertColumnsAfter(c, n) { this.maxCols += n; return this; }
  deleteColumns(c, n) { this.maxCols -= n; this.cells.forEach(row => row.splice(c - 1, n)); return this; }
  clear() { this.cells = []; return this; }
  clearConditionalFormatRules() { this.cf = []; }
  getConditionalFormatRules() { return this.cf; }
  setConditionalFormatRules(r) { this.cf = r; }
  getFilter() { return this.filter; }
  setFrozenRows(n) { this.frozenRows = n; }
  setFrozenColumns(n) { this.frozenCols = n; }
  getColumnGroupDepth(c) { return this.groups[c] ? 1 : 0; }
  getColumnGroup(c) { return chain({ collapse: () => { this.groups[c + ':collapsed'] = true; } }); }
  setColumnWidth(c, w) { (this.colW = this.colW || {})[c] = w; return this; }
  setRowHeight(r, h) { (this.rowH = this.rowH || {})[r] = h; return this; }
  autoResizeRows() { return this; }
  autoResizeColumns() { return this; } setHiddenGridlines() { return this; }
}

class Range {
  constructor(sh, r, c, nr, nc) {
    if (r < 1 || c < 1 || nr < 1 || nc < 1) throw new Error(`Range out of bounds r=${r} c=${c} nr=${nr} nc=${nc} (${sh.name})`);
    if (r + nr - 1 > sh.maxRows) throw new Error(`Range beyond max rows: ${r}+${nr} > ${sh.maxRows} (${sh.name})`);
    if (c + nc - 1 > sh.maxCols) throw new Error(`Range beyond max columns: ${c}+${nc} > ${sh.maxCols} (${sh.name})`);
    Object.assign(this, { sh, r, c, nr, nc });
    return chain(this);
  }
  a1() {
    const L = n => { let s = ''; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
    return `${this.sh.name}!${L(this.c)}${this.r}:${L(this.c + this.nc - 1)}${this.r + this.nr - 1}`;
  }
  each(fn) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) fn(this.r + i, this.c + j, i, j); }
  getSheet() { return this.sh; }
  getRow() { return this.r; }
  getNumRows() { return this.nr; }
  getNumColumns() { return this.nc; }
  getColumn() { return this.c; }
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.peek(this.r + i, this.c + j).v ?? ''); out.push(row); } return out; }
  getValue() { return this.sh.peek(this.r, this.c).v ?? ''; }
  setValues(v) {
    if (v.length !== this.nr || v.some(r => r.length !== this.nc)) throw new Error(`setValues size mismatch: got ${v.length}x${v[0] && v[0].length}, range ${this.nr}x${this.nc}`);
    this.each((r, c, i, j) => { const cell = this.sh.cell(r, c); cell.v = v[i][j] === undefined ? '' : v[i][j]; cell.rt = null; });
    return this;
  }
  setValue(v) { this.each((r, c) => { const cell = this.sh.cell(r, c); cell.v = v; cell.rt = null; }); return this; }
  getRichTextValue() {
    const cell = this.sh.peek(this.r, this.c);
    if (cell.rt) return cell.rt;
    if (cell.v instanceof Date || typeof cell.v === 'number') return null;
    return new RichText(String(cell.v ?? ''), []);
  }
  getRichTextValues() { const out = []; for (let i = 0; i < this.nr; i++) out.push([new Range(this.sh, this.r + i, this.c, 1, 1).getRichTextValue()]); return out; }
  setRichTextValue(rt) { const cell = this.sh.cell(this.r, this.c); cell.v = rt.getText(); cell.rt = rt; return this; }
  setRichTextValues(v) { v.forEach((row, i) => { const cell = this.sh.cell(this.r + i, this.c); cell.v = row[0].getText(); cell.rt = row[0]; }); return this; }
  getDataValidation() { return this.sh.peek(this.r, this.c).dv || null; }
  setDataValidation(dv) { this.each((r, c) => { this.sh.cell(r, c).dv = dv; }); return this; }
  clearDataValidations() { this.each((r, c) => { const cell = this.sh.peek(r, c); if (cell) delete cell.dv; }); return this; }
  copyTo(dst, type) {
    if (type === 'PASTE_DATA_VALIDATION') {
      for (let j = 0; j < this.nc; j++) { const dv = this.sh.peek(this.r, this.c + j).dv; const cell = dst.sh.cell(dst.r, dst.c + j); if (dv) cell.dv = dv; else delete cell.dv; }
    }
    return this;
  }
  getMergedRanges() {
    return this.sh.mergeRects
      .filter(([r, c, nr, nc]) => r >= this.r && c >= this.c && r + nr <= this.r + this.nr && c + nc <= this.c + this.nc)
      .map(([r, c, nr, nc]) => new Range(this.sh, r, c, nr, nc));
  }
  sort(spec) {
    // Устойчивая сортировка строк диапазона по колонке (как в Таблицах): даты/числа, пустые — в конце
    const col = spec.column - this.c, asc = spec.ascending !== false;
    const rows = [];
    for (let i = 0; i < this.nr; i++) rows.push(this.sh.cells[this.r - 1 + i] || []);
    const key = row => { const v = row[col] ? row[col].v : ''; return v && typeof v.getTime === 'function' ? v.getTime() : (v === '' || v == null ? null : v); };
    const sorted = rows.map((row, i) => ({ row, i })).sort((a, b) => {
      const ka = key(a.row), kb = key(b.row);
      if (ka === null && kb === null) return a.i - b.i;
      if (ka === null) return 1;
      if (kb === null) return -1;
      if (typeof ka !== typeof kb) return typeof ka === 'number' ? -1 : 1;
      return (ka < kb ? -1 : ka > kb ? 1 : 0) * (asc ? 1 : -1) || a.i - b.i;
    });
    sorted.forEach((x, i) => { this.sh.cells[this.r - 1 + i] = x.row; });
    return this;
  }
  createFilter() { const rng = this; this.sh.filter = chain({ criteria: {}, getRange: () => rng, remove: () => { this.sh.filter = null; }, setColumnFilterCriteria: (col, cr) => { this.sh.filter.criteria[col] = cr; } }); return this.sh.filter; }
  protect() { return chain({}); }
  merge() { this.sh.merges.push(this.a1()); (this.sh.mergeList = this.sh.mergeList || []).push([this.r, this.c, this.nr, this.nc]); return this; }
  fmt(k, v) { this.each((r, c) => { const cell = this.sh.cell(r, c); (cell.f = cell.f || {})[k] = v; }); return this; }
  setBackground(v) { return this.fmt('bg', v); }
  setFontColor(v) { return this.fmt('color', v); }
  setFontWeight(v) { return this.fmt('bold', v === 'bold'); }
  setFontSize(v) { return this.fmt('size', v); }
  setHorizontalAlignment(v) { return this.fmt('align', v); }
  setBorder(t, l, b, rr, v, h, color) {
    if (t && l) { this.fmt('card', true); (this.sh.cards = this.sh.cards || []).push([this.r, this.c, this.nr, this.nc]); }
    else if (b) this.fmt('underline', true);
    return this;
  }
  breakApart() { this.sh.merges = []; return this; }
  shiftColumnGroupDepth() { this.sh.groups[this.c] = true; return this; }
}

class Spreadsheet {
  constructor(fixture, id) {
    this.id = id || 'MAIN'; this.sheets = [];
    Object.entries(fixture || {}).forEach(([name, s]) => {
      const sh = new Sheet(this, name, s.rows, s.maxRows, s.maxCols);
      sh.mergeRects = s.merges || [];
      this.sheets.push(sh);
    });
  }
  getId() { return this.id; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id + '/edit'; }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  getSheets() { return this.sheets; }
  insertSheet(name, idx) { const sh = new Sheet(this, name, [], 1000, 26); this.sheets.splice(idx ?? this.sheets.length, 0, sh); return sh; }
  toast(msg) { calls.toasts.push(msg); }
}

function reviveFixture(fx) {
  const revive = v => (v && typeof v === 'object' && v.__date ? new Date(v.__date) : v);
  const out = {};
  Object.entries(fx).forEach(([k, s]) => { out[k] = { ...s, rows: s.rows.map(r => r.map(revive)) }; });
  return out;
}

function pad(n) { return String(n).padStart(2, '0'); }
function formatDate(d, tz, p) {
  return p.replace('yyyy', d.getFullYear()).replace('yy', String(d.getFullYear()).slice(2))
    .replace('MM', pad(d.getMonth() + 1)).replace('dd', pad(d.getDate())).replace('HH', pad(d.getHours())).replace('mm', pad(d.getMinutes()));
}

function makeContext(fixture, opts = {}) {
  const ss = new Spreadsheet(reviveFixture(fixture));
  const others = {};
  Object.entries(opts.others || {}).forEach(([id, fx]) => { others[id] = new Spreadsheet(reviveFixture(fx), id); });
  let userEmail = opts.email || '';
  const docProps = {}, userProps = {}, cache = {};
  const ui = {
    createMenu: name => { const m = { name, items: [], addItem(l, f) { this.items.push([l, f]); return this; }, addSeparator() { return this; }, addSubMenu(s) { this.items.push(['sub:' + s.name, s.items]); return this; }, addToUi() { calls.menus.push(this); } }; return m; },
    alert: (...a) => { calls.alerts.push(a.length > 1 ? a[1] : a[0]); return 'YES'; },
    prompt: () => ({ getSelectedButton: () => 'OK', getResponseText: () => ctx.__prompt || opts.prompt || '' }),
    Button: { OK: 'OK', YES: 'YES' }, ButtonSet: { OK_CANCEL: 1, YES_NO: 2 },
    showModalDialog: () => {}, showSidebar: () => {}
  };
  const ctx = {
    console,
    SpreadsheetApp: {
      getActive: () => ss, getActiveSpreadsheet: () => ss, getUi: () => ui, flush: () => {},
      openById: id => others[id] || ss,
      create: title => { const n = new Spreadsheet({ 'Лист1': { rows: [], maxRows: 1000, maxCols: 26 } }, 'NEW' + calls.created.length); n.title = title; calls.created.push(n); return n; },
      newRichTextValue: () => new RichBuilder(),
      newTextStyle: () => chain({ build() { return {}; } }),
      newDataValidation: () => new ValidationBuilder(),
      newConditionalFormatRule: () => new CfBuilder(),
      newFilterCriteria: () => chain({ setHiddenValues(v) { this.hidden = v; return this; }, whenTextContains(v) { this.contains = v; return this; }, build() { return this; } }),
      CopyPasteType: { PASTE_FORMAT: 'PASTE_FORMAT', PASTE_DATA_VALIDATION: 'PASTE_DATA_VALIDATION' },
      WrapStrategy: { CLIP: 'CLIP', WRAP: 'WRAP' }, BorderStyle: { SOLID: 'SOLID' }
    },
    DocumentApp: { create: () => { throw new Error('DocumentApp not mocked'); }, ParagraphHeading: {} },
    Session: { getScriptTimeZone: () => 'Europe/Moscow', getActiveUser: () => ({ getEmail: () => userEmail }) },
    Utilities: { formatDate, getUuid: () => 'uuid-1234' },
    LockService: { getScriptLock: () => ({ waitLock() { if (ctx.__lockHeld) throw new Error('deadlock: lock already held'); ctx.__lockHeld = true; }, releaseLock() { ctx.__lockHeld = false; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => 'T', setProperty() {} }),
      getDocumentProperties: () => ({ getProperty: k => docProps[k] || null, setProperty: (k, v) => { docProps[k] = v; } }),
      getUserProperties: () => ({ getProperty: k => (k in userProps ? userProps[k] : null), setProperties: o => Object.assign(userProps, o),
        deleteProperty: k => { delete userProps[k]; } }) },
    CacheService: { getDocumentCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; } }) },
    MailApp: { sendEmail: m => calls.mail.push(m) },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: h => chain({ timeBased() { return this; }, create() { calls.triggers.push(h); } }), WeekDay: { MONDAY: 1 } },
    HtmlService: { createHtmlOutputFromFile: n => { throw new Error('нет HTML-файла ' + n); }, createTemplateFromFile: n => { throw new Error('нет HTML-файла ' + n); }, createTemplate: src => ({ src, evaluate: () => chain({}) }),
      createHtmlOutput: html => { calls.html = (calls.html || []).concat(html); return chain({}); } },
    ContentService: { MimeType: { JSON: 1 }, createTextOutput: s => ({ setMimeType: () => s }) },
    Logger: console,
    __setEmail: e => { userEmail = e; },
    __clearCache: () => { Object.keys(cache).forEach(k => delete cache[k]); }
  };
  return { ctx, ss, calls, others };
}

module.exports = { makeContext, calls, RichBuilder };
