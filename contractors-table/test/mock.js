// Имитация Google Apps Script для проверки Расширение.gs в Node: один лист с объединёнными ячейками,
// проверкой данных (как в настоящей таблице — чужое значение не примет), вставкой и удалением строк.
// Часовой пояс таблицы — Москва; у «скрипта» нарочно другой, чтобы поймать даты, посчитанные не по таблице.
'use strict';

const TABLE_TZ = 'Europe/Moscow';
const MSK = 3 * 3600 * 1000;

class RichText {
  constructor(text, links) { this.text = text; this.links = links || []; }
  getText() { return this.text; }
  getLinkUrl() {
    return this.links.length === 1 && this.links[0].start === 0 && this.links[0].end === this.text.length ? this.links[0].url : null;
  }
  getRuns() {
    const cuts = new Set([0, this.text.length]);
    this.links.forEach(l => { cuts.add(l.start); cuts.add(l.end); });
    const pts = [...cuts].sort((a, b) => a - b), runs = [];
    for (let i = 0; i < pts.length - 1; i++) {
      const s = pts[i], e = pts[i + 1], l = this.links.find(x => x.start <= s && x.end >= e);
      runs.push({ getLinkUrl: () => (l ? l.url : null), getText: () => this.text.slice(s, e) });
    }
    return runs;
  }
}
class RichBuilder {
  constructor() { this.text = ''; this.links = []; }
  setText(t) { this.text = String(t); return this; }
  setLinkUrl(a, b, c) {
    if (c === undefined) this.links = [{ start: 0, end: this.text.length, url: a }];
    else {
      if (a < 0 || b > this.text.length || a >= b) throw new Error(`setLinkUrl: bad range ${a}-${b} for "${this.text}"`);
      this.links.push({ start: a, end: b, url: c });
    }
    return this;
  }
  build() { return new RichText(this.text, this.links); }
}

class Sheet {
  /** fx: { rows: [[…]], links: {'r:c': url}, merges: [[r,c,nr,nc]], lists: {col: [..]}, maxRows } — r с 1 (шапка = 1). */
  constructor(ss, name, fx) {
    this.ss = ss; this.name = name; this.id = 1338768414;
    this.cells = fx.rows.map(r => r.map(v => ({ v: v === null || v === undefined ? '' : v })));
    Object.entries(fx.links || {}).forEach(([k, url]) => {
      const [r, c] = k.split(':').map(Number), cell = this.cell(r, c);
      cell.rt = new RichText(String(cell.v), [{ start: 0, end: String(cell.v).length, url }]);
    });
    this.merges = (fx.merges || []).map(m => m.slice());
    this.lists = fx.lists || {};   // проверка данных по колонкам — у всех строк под шапкой
    this.multi = fx.multi || [];   // колонки, где можно выбрать несколько значений
    this.maxRows = Math.max(fx.maxRows || 0, this.cells.length);
    this.maxCols = fx.maxCols || 30;
    this.ops = [];                 // что меняли: для проверки «тронули только нужное»
  }
  cell(r, c) {
    while (this.cells.length < r) this.cells.push([]);
    const row = this.cells[r - 1];
    while (row.length < c) row.push({ v: '' });
    return row[c - 1];
  }
  peek(r, c) { const row = this.cells[r - 1]; return row && row[c - 1] ? row[c - 1] : { v: '' }; }
  getName() { return this.name; }
  getSheetId() { return this.id; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxCols; }
  getLastRow() {
    for (let r = this.cells.length; r >= 1; r--) if (this.cells[r - 1].some(c => c && c.v !== '' && c.v != null)) return r;
    return 0;
  }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr || 1, nc || 1); }
  mergeAt(r, c) { return this.merges.find(m => r >= m[0] && r < m[0] + m[2] && c >= m[1] && c < m[1] + m[3]) || null; }
  insertRowAfter(r) {
    if (r < 1 || r > this.maxRows) throw new Error('insertRowAfter: нет строки ' + r);
    while (this.cells.length < r) this.cells.push([]);
    this.cells.splice(r, 0, []); this.maxRows++;
    // как в Таблицах: объединение, которое проходит через место вставки, растягивается; то, что ниже, сдвигается
    this.merges.forEach(m => { if (m[0] > r) m[0]++; else if (m[0] + m[2] - 1 > r) m[2]++; });
    this.ops.push(['insertRow', r + 1]);
    return this;
  }
  deleteRow(r) {
    this.cells.splice(r - 1, 1); this.maxRows--;
    this.merges = this.merges.filter(m => {
      if (m[0] > r) { m[0]--; return true; }
      if (m[0] + m[2] - 1 >= r) { m[2]--; return m[2] > 1 || m[3] > 1; }
      return true;
    });
    this.ops.push(['deleteRow', r]);
    return this;
  }
}

class Range {
  constructor(sh, r, c, nr, nc) {
    if (r < 1 || c < 1 || nr < 1 || nc < 1) throw new Error(`Диапазон вне листа: r=${r} c=${c} nr=${nr} nc=${nc}`);
    if (r + nr - 1 > sh.maxRows) throw new Error(`Диапазон ниже последней строки листа: ${r}+${nr} > ${sh.maxRows}`);
    Object.assign(this, { sh, r, c, nr, nc });
  }
  each(fn) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) fn(this.r + i, this.c + j, i, j); }
  getRow() { return this.r; } getColumn() { return this.c; } getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.peek(this.r + i, this.c + j).v); out.push(row); }
    return out;
  }
  getValue() { return this.sh.peek(this.r, this.c).v; }
  /** Запись в ячейку: проверка данных строгая, как в таблице; в нижнюю ячейку объединения писать нельзя. */
  write(r, c, v, rt) {
    const m = this.sh.mergeAt(r, c);
    if (m && (m[0] !== r || m[1] !== c) && v !== '') throw new Error(`Запись в объединённую ячейку не с верхней строки: R${r}C${c}`);
    const list = r > 1 && this.sh.lists[c];
    if (list && v !== '' && typeof v === 'string') {
      const items = this.sh.multi.indexOf(c) !== -1 ? v.split(', ') : [v];
      const bad = items.filter(x => list.indexOf(x) === -1);
      if (bad.length) throw new Error(`Данные в ячейке R${r}C${c} не соответствуют правилам проверки данных: «${bad[0]}»`);
    }
    const cell = this.sh.cell(r, c);
    if (cell.v === v && !rt && !cell.rt) return;
    cell.v = v; cell.rt = rt || null; delete cell.formula;
    this.sh.ops.push(['set', r, c, v instanceof Date ? v.toISOString() : v]);
  }
  setValues(v) {
    if (v.length !== this.nr || v.some(row => row.length !== this.nc)) throw new Error('setValues: размер не совпал с диапазоном');
    this.each((r, c, i, j) => this.write(r, c, v[i][j] === undefined || v[i][j] === null ? '' : v[i][j]));
    return this;
  }
  setValue(v) { this.each((r, c) => this.write(r, c, v === null || v === undefined ? '' : v)); return this; }
  getRichTextValue() {
    const cell = this.sh.peek(this.r, this.c);
    if (cell.rt) return cell.rt;
    if (cell.v instanceof Date || typeof cell.v === 'number') return null;
    return new RichText(String(cell.v), []);
  }
  setRichTextValue(rt) { this.write(this.r, this.c, rt.getText(), rt); return this; }
  setFormula(f) { const cell = this.sh.cell(this.r, this.c); cell.formula = f; this.sh.ops.push(['formula', this.r, this.c, f]); return this; }
  getFormula() { return this.sh.peek(this.r, this.c).formula || ''; }
  setNumberFormat(f) { this.each((r, c) => { this.sh.cell(r, c).num = f; }); return this; }
  setNote(n) { this.sh.cell(this.r, this.c).note = String(n); this.sh.ops.push(['note', this.r, this.c]); return this; }
  getNote() { return this.sh.peek(this.r, this.c).note || ''; }
  getDataValidation() {
    const list = this.r > 1 && this.sh.lists[this.c];
    const own = this.sh.peek(this.r, this.c).dv;
    if (own === null || !list) return null; // null — у ячейки проверку сняли
    return { getCriteriaType: () => 'VALUE_IN_LIST', getCriteriaValues: () => [list.slice(), true] };
  }
  /** Объединения, которые хотя бы одной ячейкой попадают в диапазон. */
  getMergedRanges() {
    return this.sh.merges
      .filter(([r, c, nr, nc]) => r < this.r + this.nr && r + nr > this.r && c < this.c + this.nc && c + nc > this.c)
      .map(([r, c, nr, nc]) => new Range(this.sh, r, c, nr, nc));
  }
  isPartOfMerge() { return this.getMergedRanges().length > 0; }
  mergeVertically() {
    for (let j = 0; j < this.nc; j++) {
      const c = this.c + j;
      this.sh.merges = this.sh.merges.filter(([r, mc, nr, nc]) => !(mc <= c && mc + nc > c && r < this.r + this.nr && r + nr > this.r));
      for (let i = 1; i < this.nr; i++) { const cell = this.sh.cell(this.r + i, c); cell.v = ''; cell.rt = null; } // остаётся верхнее значение
      this.sh.merges.push([this.r, c, this.nr, 1]);
      this.sh.ops.push(['merge', this.r, c, this.nr]);
    }
    return this;
  }
  copyTo(dst, type) {
    this.sh.ops.push(['copy', type, this.r, dst.r]);
    if (type === 'PASTE_FORMAT' && this.isPartOfMerge()) throw new Error('копирование формата со строки с объединёнными ячейками');
    return this;
  }
}

class Spreadsheet {
  constructor(sheets) { this.sheets = Object.entries(sheets).map(([name, fx]) => new Sheet(this, name, fx)); }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/OLDTABLEID0123456789abcdefghij/edit'; }
  getSpreadsheetTimeZone() { return TABLE_TZ; }
}

function formatDate(d, tz, p) {
  if (tz !== TABLE_TZ) throw new Error('Дата посчитана не по часовому поясу таблицы: ' + tz);
  const x = new Date(d.getTime() + MSK), pad = n => String(n).padStart(2, '0');
  return p.replace('yyyy', x.getUTCFullYear()).replace('yy', String(x.getUTCFullYear()).slice(2)).replace('MM', pad(x.getUTCMonth() + 1))
    .replace('dd', pad(x.getUTCDate())).replace('HH', pad(x.getUTCHours())).replace('mm', pad(x.getUTCMinutes()));
}
function parseDate(s, tz, p) {
  if (tz !== TABLE_TZ || p !== 'yyyy-MM-dd') throw new Error('parseDate: ждала часовой пояс таблицы и yyyy-MM-dd');
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) - MSK);
}
/** Полночь по Москве — так даты лежат в таблице. */
const mskDate = iso => parseDate(iso, TABLE_TZ, 'yyyy-MM-dd');

function reviveFixture(fx) {
  const revive = v => (v && typeof v === 'object' && v.__date ? mskDate(v.__date) : v);
  return { ...fx, rows: fx.rows.map(r => r.map(revive)) };
}

function makeContext(fixture, opts = {}) {
  const ss = new Spreadsheet({ 'Localization Misc': reviveFixture(fixture) });
  const cache = {}, props = { SMETA_TOKEN: 'T' };
  let lockHeld = false;
  const ctx = {
    console: { log() {} },
    SpreadsheetApp: {
      openById: id => { ctx.__opened = id; return ss; },
      getActiveSpreadsheet: () => null, getActive: () => null, // отдельный проект: «своей» таблицы нет
      flush() {},
      newRichTextValue: () => new RichBuilder(),
      CopyPasteType: { PASTE_FORMAT: 'PASTE_FORMAT', PASTE_DATA_VALIDATION: 'PASTE_DATA_VALIDATION' },
      DataValidationCriteria: { VALUE_IN_LIST: 'VALUE_IN_LIST', VALUE_IN_RANGE: 'VALUE_IN_RANGE' }
    },
    Session: { getScriptTimeZone: () => 'America/New_York' },
    Utilities: { formatDate, parseDate, getUuid: () => 'uuid-1234' },
    LockService: { getScriptLock: () => ({ waitLock() { if (lockHeld) throw new Error('блокировка уже взята'); lockHeld = true; }, releaseLock() { lockHeld = false; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    CacheService: { getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; } }) },
    ContentService: { MimeType: { JSON: 1 }, createTextOutput: s => ({ setMimeType: () => s, text: s }) },
    Logger: { log: (...a) => { (ctx.__log = ctx.__log || []).push(a.join(' ')); } },
    Date: opts.now ? class extends Date { constructor(...a) { if (a.length) super(...a); else super(opts.now); } static now() { return new Date(opts.now).getTime(); } } : Date,
    __clearCache: () => { Object.keys(cache).forEach(k => delete cache[k]); }
  };
  return { ctx, ss, sheet: ss.sheets[0] };
}

module.exports = { makeContext, mskDate, formatDate, TABLE_TZ };
