/*************************************************************
 * Расширение.gs — приёмник расширения «Задачи и сметы» для таблицы подрядчиков
 * (лист «Localization Misc»).
 *
 * Это ОТДЕЛЬНЫЙ проект Apps Script (script.google.com → «Создать проект»): в саму таблицу
 * ничего не добавляется, её скрипты и меню не меняются. Скрипт открывает таблицу по ссылке
 * и пишет в неё от имени того, кто его развернул.
 *
 * Что делает: новая задача с номером (LIT-26-…), «Мои задачи», правка, статус, запись сметы.
 * Что НЕ делает: не красит, не сортирует, не трогает чужие строки — меняет только те ячейки,
 * которые поменяли в расширении. Колонки сдвинулись — запись останавливается с понятной ошибкой.
 *
 * Установка — README.md рядом с этим файлом.
 *************************************************************/

// Ссылка на таблицу подрядчиков (из адресной строки браузера) — впишите между кавычками.
const TABLE_LINK = '';

const SHEET_NAME = 'Localization Misc';
const SCRIPT_VERSION = '12 ноября (таблица подрядчиков: задачи, «Мои», правка, сметы)';

// Колонки листа A..S
const COL = {
  ID: 1, PRODUCT: 2, DATE: 3, SUBJECT: 4, TICKET: 5, BAND: 6, CUSTOMER: 7, LANGS: 8, DEADLINE: 9, DUE: 10,
  STATUS: 11, ESTIMATE: 12, TOTAL: 13, CONTRACTOR: 14, MANAGER: 15, DELIVERY: 16, SP: 17, COMMENT: 18, COMPLAINTS: 19
};
const NCOLS = 19;

// Шапка, по которой проверяем, что колонки на месте (начало названия, без учёта регистра)
const HEADER_CHECK = [
  [COL.PRODUCT, 'сторона'], [COL.DATE, 'дата'], [COL.SUBJECT, 'тема письма'], [COL.TICKET, 'задача'],
  [COL.BAND, 'тема в б'], [COL.CUSTOMER, 'ник'], [COL.LANGS, 'языки'], [COL.DEADLINE, 'дедлайн'], [COL.DUE, 'дедлайн'],
  [COL.STATUS, 'статус'], [COL.ESTIMATE, 'смета'], [COL.TOTAL, 'итог'], [COL.CONTRACTOR, 'подрядчик'],
  [COL.MANAGER, 'менеджер'], [COL.DELIVERY, 'статус отдачи'], [COL.SP, 'sp']
];

const CLOSED_STATUSES = ['Отдано', 'Отменено'];

// Префиксы номеров — те же, что у автономера в самой таблице
const CONTRACTOR_PREFIX = {
  'LogrusIT': 'LIT', 'LogrusGlobal': 'LoG', 'Бюро переводов': 'BP', 'JanusWW': 'JWW', 'Awatera': 'AWT'
};

// Коды языков для темы письма
const LANG_CODES = {
  'Азербайджанский': 'az', 'Английский': 'en', 'Амхарский': 'am', 'Армянский': 'hy',
  'Грузинский': 'ka', 'Казахский': 'kk', 'Китайский': 'zhs', 'Кыргызский': 'ky',
  'Русский': 'ru', 'Таджикский': 'tg', 'Узбекский': 'uz', 'Суахили': 'sw',
  'Белорусский': 'be', 'Иврит': 'he', 'Корейский': 'ko', 'Турецкий': 'tr', 'Арабский': 'ar', 'Французский': 'fr'
};
const RARE_LANGS = ['Белорусский', 'Иврит', 'Корейский', 'Турецкий', 'Арабский', 'Французский'];

// Объединённые ячейки читаем по-настоящему только у последних строк (там живые задачи): у всего листа это долго.
// Выше — по виду: строка без номера и без темы продолжает заказ над ней.
const TAIL_ROWS = 400;
const DUP_SCAN_ROWS = 800;
const MY_TASKS_LIMIT = 100;
const LISTS_CACHE_SEC = 600;

// ==================== ПЕРВЫЙ ЗАПУСК ====================

/** Запустите один раз из редактора: создаст токен и покажет его в журнале выполнения. */
function setupToken() {
  const token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('SMETA_TOKEN', token);
  Logger.log('Токен для расширения: ' + token);
}

/** Запустите из редактора, чтобы проверить доступ: напишет, что скрипт видит в таблице. Ничего не меняет. */
function checkTable() {
  const sh = sheet_();
  const raw = readRaw_(sh);
  checkHeader_(sh);
  const lists = lists_(sh, raw, true);
  Logger.log('Лист «' + SHEET_NAME + '» найден, строк с задачами: ' + raw.length);
  Logger.log('Последняя строка: ' + (raw.length + 1) + ' — ' + str_(raw.length ? raw[raw.length - 1][COL.ID - 1] : ''));
  Logger.log('Подрядчики: ' + lists.contractors.join(', '));
  Logger.log('Менеджеры: ' + lists.managers.join(', '));
  Logger.log('Следующие номера: ' + JSON.stringify(nextIds_(raw, lists)));
}

// ==================== ВХОД ====================

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Некорректный запрос' });
  }
  const token = PropertiesService.getScriptProperties().getProperty('SMETA_TOKEN');
  if (!token || req.token !== token) return json_({ ok: false, error: 'Неверный токен' });

  // Чтение — без очереди
  try {
    const read = READ_ACTIONS[req.action];
    if (read) return json_(read(req));
    if (LATER_ACTIONS.indexOf(req.action) !== -1) {
      return json_({ ok: false, error: 'В таблице подрядчиков это пока не подключено — только задачи, «Мои» и сметы' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }

  // Запись — по одному, чтобы два человека не получили один номер
  const write = WRITE_ACTIONS[req.action];
  if (!write) return json_({ ok: false, error: 'Неизвестное действие' });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return json_({ ok: false, error: 'Таблица занята другим запросом, попробуйте ещё раз через минуту' });
  }
  try {
    const res = write(req);
    console.log(req.action + ' · ' + str_(req.user) + ' · ' + str_(res && (res.id || res.row) || req.id || req.task && req.task.id || req.task));
    return json_(res);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    SpreadsheetApp.flush(); // записать номер до того, как следующий его прочитает
    lock.releaseLock();
  }
}

/** Проверка: открыть адрес веб-приложения в браузере — должно написать, что скрипт работает. */
function doGet() {
  return ContentService.createTextOutput('Скрипт работает. Адрес правильный — вставьте его в настройки расширения.\n\nВерсия: ' + SCRIPT_VERSION);
}

const READ_ACTIONS = {
  managers: () => { const t = table_(); return { ok: true, managers: t.lists.managers }; },
  taskForm: () => taskForm_(),
  previewTaskId: req => { const t = table_(); const c = str_(req.contractor); return { ok: true, id: c ? nextId_(t.raw, c) : '' }; },
  previewTaskIds: () => { const t = table_(); return { ok: true, ids: nextIds_(t.raw, t.lists) }; },
  checkDuplicates: req => ({ ok: true, duplicates: findDuplicates_(req.task || {}) }),
  myTasks: req => Object.assign({ ok: true }, myTasks_(req.manager || '*', req.filter)),
  searchTasks: req => ({ ok: true, tasks: searchTasks_(str_(req.query)) }),
  getTask: req => getTask_(req),
  lookup: req => smetaLookup_(req)
};

const WRITE_ACTIONS = {
  addTask: req => addTask_(req),
  addTaskRow: req => addTaskRow_(req),
  saveTask: req => saveTask_(req),
  setStatus: req => setCell_(req, COL.STATUS, req.status),
  setDelivery: req => setCell_(req, COL.DELIVERY, req.value),
  deleteTask: req => deleteTask_(req),
  write: req => smetaWrite_(req)
};

// Отчёты и письма по расписанию — следующим шагом
const LATER_ACTIONS = ['reportLists', 'managerReport', 'trackerReport', 'customReport', 'moneyPreview', 'actCompare',
  'exportXlsx', 'mailStatus', 'mailTest', 'mailToggle', 'mailEmail'];

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ==================== ТАБЛИЦА ====================

var SS_ = null;
function ss_() {
  if (SS_) return SS_;
  const m = String(TABLE_LINK).match(/\/d\/([a-zA-Z0-9_-]{20,})/) || String(TABLE_LINK).trim().match(/^([a-zA-Z0-9_-]{20,})$/);
  if (!m) throw new Error('В скрипте не вписана ссылка на таблицу (TABLE_LINK в начале файла)');
  SS_ = SpreadsheetApp.openById(m[1]);
  return SS_;
}

function sheet_() {
  const sh = ss_().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('В таблице нет листа «' + SHEET_NAME + '»');
  return sh;
}

function tz_() {
  return ss_().getSpreadsheetTimeZone();
}

/** Колонки на месте? Если в лист вставили или убрали колонку, писать нельзя — попадём не туда. */
function checkHeader_(sh) {
  const head = sh.getRange(1, 1, 1, NCOLS).getValues()[0].map(h => str_(h).toLowerCase());
  HEADER_CHECK.forEach(pair => {
    if (head[pair[0] - 1].indexOf(pair[1]) !== 0) {
      throw new Error('В таблице изменились колонки (в ' + colLetter_(pair[0]) + ' ждала «' + pair[1] + '…», а там «' + head[pair[0] - 1] +
        '»). Расширение остановлено, чтобы ничего не испортить, — напишите тому, кто его настраивал.');
    }
  });
  return head;
}

function colLetter_(c) {
  return String.fromCharCode(64 + c);
}

/** Строки с задачами как есть (без шапки). Пустой хвост обрезан: ниже данных в листе — заготовки строк. */
function readRaw_(sh) {
  const n = Math.max(sh.getLastRow() - 1, 0);
  const raw = n ? sh.getRange(2, 1, n, NCOLS).getValues() : [];
  let end = raw.length;
  while (end && raw[end - 1].every(v => v === '' || v === null)) end--;
  raw.length = end;
  return raw;
}

/**
 * Лист целиком для одного запроса: raw — как в ячейках; rows — с «протянутыми» объединёнными ячейками;
 * anchor['i:c'] — индекс верхней строки объединения для ячейки (строка i, колонка c); head[i] — первая строка заказа.
 */
function table_(opts) {
  const sh = sheet_();
  const raw = readRaw_(sh);
  const rows = raw.map(r => r.slice());
  const anchor = {};
  const head = rows.map((r, i) => i);
  const from = Math.max(0, raw.length - TAIL_ROWS);

  // 1) Последние строки — по настоящим объединениям
  if (raw.length) {
    sh.getRange(from + 2, 1, raw.length - from, NCOLS).getMergedRanges().forEach(m => {
      const nr = m.getNumRows();
      if (nr < 2) return;
      const top = m.getRow() - 2, c0 = m.getColumn(), c1 = Math.min(c0 + m.getNumColumns() - 1, NCOLS);
      if (top < 0) return;
      for (let c = c0; c <= c1; c++) {
        for (let i = top + 1; i < top + nr && i < rows.length; i++) {
          rows[i][c - 1] = rows[top][c - 1];
          anchor[i + ':' + c] = top;
          if (c === COL.ID) head[i] = head[top];
        }
      }
    });
  }
  // 2) Выше — по виду: без номера, а тема пустая или та же, что строкой выше → продолжение заказа над ней
  const INHERIT = [COL.ID, COL.SUBJECT, COL.DATE, COL.TICKET, COL.CONTRACTOR, COL.MANAGER, COL.PRODUCT];
  for (let i = 1; i < from; i++) {
    const r = raw[i], up = rows[i - 1];
    if (r[COL.ID - 1] !== '' || !up[COL.SUBJECT - 1] || up[COL.ID - 1] === '') continue;
    if (r[COL.SUBJECT - 1] !== '' && r[COL.SUBJECT - 1] !== up[COL.SUBJECT - 1]) continue;
    if (r.every(v => v === '' || v === null)) continue;
    INHERIT.forEach(c => { if (rows[i][c - 1] === '') rows[i][c - 1] = up[c - 1]; });
    head[i] = head[i - 1];
  }
  const t = { sh: sh, raw: raw, rows: rows, anchor: anchor, head: head };
  t.lists = lists_(sh, raw, opts && opts.freshLists);
  return t;
}

/**
 * Точные объединения одной строки (один быстрый запрос): выше «хвоста» листа table_ их не читает.
 * Заодно подтягивает значения объединённых ячеек с верхней строки.
 */
function anchorsFor_(t, i) {
  t.sh.getRange(i + 2, 1, 1, NCOLS).getMergedRanges().forEach(m => {
    const top = m.getRow() - 2;
    if (m.getNumRows() < 2 || top === i || top < 0) return;
    const c1 = Math.min(m.getColumn() + m.getNumColumns() - 1, NCOLS);
    for (let c = m.getColumn(); c <= c1; c++) {
      t.anchor[i + ':' + c] = top;
      if (t.raw[i][c - 1] === '') t.rows[i][c - 1] = t.rows[top][c - 1];
    }
  });
}

/** Строка листа, куда писать ячейку (строка i с 0, колонка c): у объединённой — верхняя ячейка. */
function targetRow_(t, i, c) {
  const a = t.anchor[i + ':' + c];
  return (a === undefined ? i : a) + 2;
}

// ==================== СПРАВОЧНИКИ ====================

/** Выпадающие списки — те же, что в ячейках таблицы (проверка данных). Запоминаются на 10 минут. */
function lists_(sh, raw, fresh) {
  let cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) { /* без памяти */ }
  const key = 'lists:v2:' + sh.getSheetId();
  if (cache && !fresh) {
    const hit = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch (e) { /* читаем заново */ } }
  }
  const header = sh.getRange(1, 1, 1, NCOLS).getValues()[0].map(str_);
  const last = raw.length + 1; // последняя строка с задачей
  const tryRows = [last, last + 1, last - 1, last - 2].filter(r => r >= 2 && r <= sh.getMaxRows());
  const listOf = col => {
    for (let k = 0; k < tryRows.length; k++) {
      const items = validationList_(sh.getRange(tryRows[k], col));
      if (items.length) return clean(items, col);
    }
    // проверки данных нет — берём то, что чаще всего встречается в последних строках
    const seen = {};
    raw.slice(-600).forEach(r => String(r[col - 1] || '').split(col === COL.LANGS ? /\s*,\s*/ : /\n{99}/).forEach(v => { v = v.trim(); if (v) seen[v] = (seen[v] || 0) + 1; }));
    return clean(Object.keys(seen).sort((a, b) => seen[b] - seen[a]), col);
  };
  const clean = (items, col) => {
    const out = [];
    items.forEach(v => {
      v = str_(v);
      if (!v || v === header[col - 1] || out.indexOf(v) !== -1) return; // в списках первым пунктом стоит название колонки
      out.push(v);
    });
    return out;
  };
  const lists = {
    products: listOf(COL.PRODUCT), langs: listOf(COL.LANGS).filter(l => l.indexOf(',') === -1), deadlines: listOf(COL.DEADLINE),
    statuses: listOf(COL.STATUS), contractors: listOf(COL.CONTRACTOR), managers: listOf(COL.MANAGER),
    deliveryStatuses: listOf(COL.DELIVERY)
  };
  lists.langCodes = {};
  Object.keys(LANG_CODES).forEach(l => { lists.langCodes[l] = LANG_CODES[l]; });
  if (cache) { try { cache.put(key, JSON.stringify(lists), LISTS_CACHE_SEC); } catch (e) { /* не страшно */ } }
  return lists;
}

function validationList_(range) {
  const dv = range.getDataValidation();
  if (!dv) return [];
  const type = dv.getCriteriaType(), vals = dv.getCriteriaValues();
  const T = SpreadsheetApp.DataValidationCriteria;
  if (type === T.VALUE_IN_LIST) return (vals[0] || []).map(String);
  if (type === T.VALUE_IN_RANGE && vals[0]) return [].concat.apply([], vals[0].getValues()).map(String);
  return [];
}

function taskForm_() {
  const t = table_();
  const l = t.lists;
  const shtat = l.langs.filter(x => /^ШТАТ/i.test(x));
  const rare = l.langs.filter(x => RARE_LANGS.indexOf(x) !== -1);
  const regular = l.langs.filter(x => shtat.indexOf(x) === -1 && rare.indexOf(x) === -1);
  return { ok: true, lists: {
    contractors: l.contractors, products: l.products, deadlines: l.deadlines, statuses: l.statuses,
    deliveryStatuses: l.deliveryStatuses, managers: l.managers,
    languages: { regular: regular, shtat: shtat, rare: rare }, langCodes: l.langCodes,
    currentManager: '', templates: [], tableUrl: tableUrl_(t.sh),
    features: { reports: false, mail: false }, // расширение прячет то, что здесь не подключено
    table: 'contractors'
  } };
}

function tableUrl_(sh, row) {
  const url = ss_().getUrl().replace(/#.*$/, '') + '#gid=' + sh.getSheetId();
  return row ? url + '&range=A' + row : url;
}

// ==================== МЕЛОЧИ ====================

function str_(v) {
  return String(v == null ? '' : v).trim();
}

function isDate_(d) {
  return d != null && typeof d.getFullYear === 'function' && !isNaN(d.getTime());
}

function fmtDate_(d, pattern) {
  return isDate_(d) ? Utilities.formatDate(d, tz_(), pattern) : '';
}

/** «2026-10-01» → полночь этого дня по часам ТАБЛИЦЫ (у отдельного скрипта свой часовой пояс — на него не смотрим). */
function parseIsoDate_(s) {
  s = str_(s);
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? Utilities.parseDate(s, tz_(), 'yyyy-MM-dd') : '';
}

function today_() {
  return parseIsoDate_(fmtDate_(new Date(), 'yyyy-MM-dd'));
}

function exactMoney_(v) {
  return Math.round(Number(v) * 1000) / 1000;
}

/** Сумма из поля: число или «3000+2500,50» → { value, formula }. */
function moneyInput_(v) {
  if (v === '' || v == null) return { value: '' };
  if (typeof v === 'number') return { value: exactMoney_(v) };
  const s = String(v).replace(/[\s  ₽]/g, '').replace(/^=/, '');
  if (!s) return { value: '' };
  const parts = s.split('+').map(p => Number(p.replace(',', '.')));
  if (!parts.length || parts.some(p => !isFinite(p) || p < 0)) throw new Error('Сумма не похожа на число: ' + v);
  const value = exactMoney_(parts.reduce((a, b) => a + b, 0));
  return parts.length > 1 ? { value: value, formula: '=' + parts.map(p => String(exactMoney_(p))).join('+') } : { value: value };
}

function moneyFormulaOf_(range) {
  const f = String(range.getFormula() || '');
  return /^=\s*[\d.]+(\s*\+\s*[\d.]+)+\s*$/.test(f) ? f.replace(/^=\s*/, '').replace(/\s+/g, '') : '';
}

function rub_(x) {
  if (x === '' || x == null || !isFinite(Number(x))) return String(x == null ? '' : x);
  const n = exactMoney_(x);
  const digits = Math.round(Math.abs(n) * 1000) % 10 ? 3 : 2;
  const parts = Math.abs(n).toFixed(digits).split('.');
  return (n < 0 ? '−' : '') + parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ',' + parts[1];
}

function moneyFormat_(x) {
  return typeof x === 'number' && Math.round(Math.abs(x) * 1000) % 10 ? '#,##0.000' : '#,##0.00';
}

/** «[LIT-8][LogrusIT][az]… Новые строчки» → «Новые строчки». */
function shortSubject_(s) {
  return String(s || '').replace(/^(\s*\[[^\]]*\])+\s*/, '').trim() || String(s || '');
}

/** «1234», «local 1234», «LOCAL-1234» → «LOCAL-1234». Другой текст (в старых строках в «Задаче» — описание) — как есть. */
function normTicket_(t) {
  const s = str_(t);
  const m = s.toUpperCase().replace(/\s+/g, '').match(/^(?:LOCAL-?)?(\d+)$/);
  return m ? 'LOCAL-' + m[1] : s;
}

function urls_(s) {
  return String(s || '').match(/https?:\/\/[^\s,;]+/g) || [];
}

/** Ссылки на Band в ячейке: по одной в строке, каждая кликабельна. */
function linksRich_(urls) {
  const text = urls.join('\n');
  let b = SpreadsheetApp.newRichTextValue().setText(text);
  let pos = 0;
  urls.forEach(u => { b = b.setLinkUrl(pos, pos + u.length, u); pos += u.length + 1; });
  return b.build();
}

/** Адрес сметы: ссылка под текстом ячейки или сам текст, если это адрес. */
function estimateUrl_(range, value) {
  const v = str_(value);
  if (/^https?:\/\/\S+$/i.test(v)) return v;
  let rt = null;
  try { rt = range.getRichTextValue(); } catch (e) { /* число или дата */ }
  if (rt) {
    const url = rt.getLinkUrl() || rt.getRuns().map(r => r.getLinkUrl()).filter(Boolean)[0];
    if (url) return url;
  }
  return /^(Смета|Ссылка на смету)( \[[^\]]*\]| \S+)?$/.test(v) ? '' : v;
}

// Сколько рабочих дней до срока → значение из списка «Дедлайн»
const DEADLINE_STEPS = [[0, 'ASAP'], [2, '1-2 дня'], [5, 'До недели'], [21, 'Больше недели'], [Infinity, 'Месяц и больше']];
function workDays_(a, b) {
  if (b <= a) return b - a;
  if (b - a > 60) return b - a;
  let n = 0;
  for (let d = a + 1; d <= b; d++) { const w = (d + 4) % 7; if (w !== 0 && w !== 6) n++; }
  return n;
}
function deadlineFromDue_(due, received, lists) {
  if (!isDate_(due)) return '';
  const day = d => { const p = fmtDate_(d, 'yyyy-MM-dd').split('-').map(Number); return Math.round(Date.UTC(p[0], p[1] - 1, p[2]) / 864e5); };
  const days = workDays_(day(isDate_(received) ? received : new Date()), day(due));
  const label = DEADLINE_STEPS.filter(step => days <= step[0])[0][1].toLowerCase();
  return lists.deadlines.filter(v => v.toLowerCase() === label)[0] || '';
}

// ==================== НОМЕР ЗАДАЧИ ====================

function prefix_(contractor) {
  return CONTRACTOR_PREFIX[contractor] || 'TASK';
}

/**
 * Следующий номер ПОДРЯДЧИК-ГГ-N: самый большой в листе + 1. Смотрим и колонку A, и номера в темах писем
 * («[LIT-26-2499]…» — на случай, если тему уже вписали, а номер в A ещё не появился).
 */
function nextId_(raw, contractor) {
  const base = prefix_(contractor) + '-' + Utilities.formatDate(new Date(), tz_(), 'yy');
  const esc = base.replace(/[-\\^$*+?.()|[\]{}]/g, '\\$&');
  const reId = new RegExp('^' + esc + '-(\\d+)$', 'i'), reSubj = new RegExp('^\\s*\\[' + esc + '-(\\d+)\\]', 'i');
  let max = 0;
  raw.forEach(r => {
    const m = str_(r[COL.ID - 1]).match(reId) || String(r[COL.SUBJECT - 1] || '').match(reSubj);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return base + '-' + (max + 1);
}

function nextIds_(raw, lists) {
  const out = {};
  lists.contractors.forEach(c => { out[c] = nextId_(raw, c); });
  return out;
}

/** Тема письма: [номер][подрядчик][коды языков][продукт] суть. */
function buildSubject_(id, task, lists) {
  const codes = (task.languages || []).map(l => lists.langCodes[l]).filter(Boolean);
  const prefix = '[' + id + ']' + (task.contractor ? '[' + task.contractor + ']' : '') +
    codes.map(c => '[' + c + ']').join('') + (task.product ? '[' + task.product + ']' : '');
  const raw = str_(task.subject);
  return raw ? prefix + ' ' + raw : prefix;
}

// ==================== НОВАЯ ЗАДАЧА ====================

function cleanTask_(t) {
  t = t || {};
  const task = {};
  ['ticket', 'contractor', 'subject', 'link', 'link2', 'date', 'product', 'customer', 'deadline', 'exactDeadline', 'status',
    'deliveryStatus', 'estimateLink', 'total', 'sp', 'manager', 'comment', 'complaints'].forEach(k => { task[k] = str_(t[k]); });
  task.languages = (Array.isArray(t.languages) ? t.languages : str_(t.languages).split(/\s*,\s*/)).map(str_).filter(Boolean);
  task.ticket = normTicket_(task.ticket);
  [task.link].concat(task.link2.split(/[\s,;]+/)).forEach(requireUrl_);
  task.bandUrls = urls_(task.link).concat(urls_(task.link2));
  return task;
}

function requireUrl_(u) {
  if (u && !/^https?:\/\//i.test(u)) throw new Error('Ссылка должна начинаться с http: ' + u);
}

function requireTicket_(ticket) {
  if (ticket && !/^LOCAL-\d+$/.test(ticket)) throw new Error('Тикет должен быть вида LOCAL-1234');
}

/**
 * Новая строка сразу под последней задачей. Именно ВСТАВЛЯЕМ строку, а не пишем в пустую:
 * если кто-то в этот момент печатает в следующей пустой строке, его строка просто сдвинется ниже.
 * Оформление и выпадающие списки — с пустой строки-заготовки под задачами (tpl — её номер уже после вставки).
 */
function insertRow_(sh, after, tpl) {
  sh.insertRowAfter(after);
  const row = after + 1;
  if (tpl && tpl <= sh.getMaxRows()) {
    const src = sh.getRange(tpl, 1, 1, NCOLS), dst = sh.getRange(row, 1, 1, NCOLS);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
  return row;
}

/** Записать строку; таблица не приняла значение (проверка данных) — убрать вставленную строку и объяснить. */
function fillRow_(sh, row, values, task) {
  try {
    sh.getRange(row, 1, 1, NCOLS).setValues([values]);
    if (task.bandUrls && task.bandUrls.length) sh.getRange(row, COL.BAND).setRichTextValue(linksRich_(task.bandUrls));
    if (task.estimateLink) {
      sh.getRange(row, COL.ESTIMATE).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(task.estimateLink).setLinkUrl(task.estimateLink).build());
    }
    if (task.totalIn && task.totalIn.formula) sh.getRange(row, COL.TOTAL).setFormula(task.totalIn.formula);
    if (values[COL.TOTAL - 1] !== '') sh.getRange(row, COL.TOTAL).setNumberFormat(moneyFormat_(values[COL.TOTAL - 1]));
  } catch (err) {
    try { sh.deleteRow(row); } catch (e2) { /* останется пустая строка */ }
    throw new Error('Таблица не приняла строку: ' + String(err && err.message || err));
  }
}

function addTask_(req) {
  const task = cleanTask_(req.task);
  if (!task.contractor) return { ok: false, error: 'Выберите подрядчика' };
  if (!task.subject) return { ok: false, error: 'Впишите тему' };
  requireTicket_(task.ticket);
  requireUrl_(task.estimateLink);

  const t = table_();
  checkHeader_(t.sh);
  const id = nextId_(t.raw, task.contractor);
  const subject = buildSubject_(id, task, t.lists);
  const date = parseIsoDate_(task.date), due = parseIsoDate_(task.exactDeadline);
  task.totalIn = moneyInput_(task.total);
  const values = [
    id, task.product, date, subject, task.ticket, task.bandUrls.join('\n'), task.customer, task.languages.join(', '),
    task.deadline || deadlineFromDue_(due, date, t.lists), due, task.status, task.estimateLink, task.totalIn.value,
    task.contractor, task.manager, task.deliveryStatus, task.sp ? Number(task.sp) : '', task.comment, task.complaints
  ];
  const row = insertRow_(t.sh, t.raw.length + 1, t.raw.length + 3); // заготовка — пустая строка, что была следующей
  fillRow_(t.sh, row, values, task);
  return { ok: true, id: id, row: row, url: tableUrl_(t.sh, row) };
}

/**
 * Ещё строка в существующий заказ — сразу под его строками, как это делают в таблице руками:
 * номер и тема объединяются на все строки заказа, у новой строки свои Band, ник, языки, сроки.
 */
function addTaskRow_(req) {
  const id = str_(req.id);
  if (!id) throw new Error('Нет номера заказа');
  const task = cleanTask_(req.task);
  requireTicket_(task.ticket);
  const t = table_();
  checkHeader_(t.sh);
  const sh = t.sh;

  let first = -1, last = -1;
  t.rows.forEach((r, i) => { if (str_(r[COL.ID - 1]) === id) { if (first < 0) first = i; last = i; } });
  if (first < 0) throw new Error('Не нашла заказ ' + id + ' — возможно, его удалили или поменяли номер');
  for (let i = first; i <= last; i++) {
    if (str_(t.rows[i][COL.ID - 1]) !== id) throw new Error('Строки заказа ' + id + ' идут не подряд — добавьте строку в таблице вручную');
  }
  const base = t.rows[first];
  const date = isDate_(base[COL.DATE - 1]) ? base[COL.DATE - 1] : parseIsoDate_(task.date);
  const due = parseIsoDate_(task.exactDeadline);

  // Тема заказа — с кодами языков всех его строк
  const langs = [];
  for (let i = first; i <= last; i++) str_(t.rows[i][COL.LANGS - 1]).split(/\s*,\s*/).forEach(l => { if (l && langs.indexOf(l) === -1) langs.push(l); });
  task.languages.forEach(l => { if (langs.indexOf(l) === -1) langs.push(l); });
  const oldSubject = str_(base[COL.SUBJECT - 1]);
  const subject = buildSubject_(id, {
    contractor: str_(base[COL.CONTRACTOR - 1]), product: str_(base[COL.PRODUCT - 1]), languages: langs,
    subject: oldSubject.replace(/^(\s*\[[^\]]*\])+\s*/, '')
  }, t.lists);

  // Номер и тему объединяем, только если у остальных строк заказа эти ячейки пустые (как при объединении руками)
  const below = c => { for (let i = first + 1; i <= last; i++) if (t.raw[i][c - 1] !== '') return false; return true; };
  const mergeId = below(COL.ID), mergeSubject = below(COL.SUBJECT);

  const values = [
    mergeId ? '' : id, task.product, date || '', mergeSubject ? '' : subject, task.ticket || str_(base[COL.TICKET - 1]),
    task.bandUrls.join('\n'), task.customer, task.languages.join(', '),
    task.deadline || deadlineFromDue_(due, date, t.lists), due, task.status, '', '',
    str_(base[COL.CONTRACTOR - 1]), task.manager || str_(base[COL.MANAGER - 1]),
    task.deliveryStatus, task.sp ? Number(task.sp) : '', task.comment, ''
  ];
  // Последняя строка заказа объединена с задачей под ней (бывает у «Задачи», Band и ника): вставленная строка
  // попала бы внутрь такого объединения, и её ячейки было бы не заполнить
  if (last + 1 < t.raw.length) {
    const cross = sh.getRange(last + 2, 1, 2, NCOLS).getMergedRanges()
      .filter(m => m.getRow() <= last + 2 && m.getRow() + m.getNumRows() - 1 >= last + 3);
    if (cross.length) {
      const cols = cross.map(m => colLetter_(m.getColumn())).join(', ');
      throw new Error('Под заказом ' + id + ' ячейки (' + cols + ') объединены со следующей задачей — добавьте строку в таблице вручную.');
    }
  }
  const row = insertRow_(sh, last + 2, t.raw.length + 3);
  fillRow_(sh, row, values, task);
  const n = row - (first + 2) + 1;
  if (mergeId) sh.getRange(first + 2, COL.ID, n, 1).mergeVertically();
  if (mergeSubject) sh.getRange(first + 2, COL.SUBJECT, n, 1).mergeVertically();
  if (subject !== oldSubject) {
    if (mergeSubject) sh.getRange(first + 2, COL.SUBJECT).setValue(subject);
    else for (let i = first; i <= last; i++) if (t.raw[i][COL.SUBJECT - 1] !== '') sh.getRange(i + 2, COL.SUBJECT).setValue(subject);
  }
  return { ok: true, id: id, row: row, subject: subject, url: tableUrl_(sh, row) };
}

// ==================== ПОИСК, «МОИ», ПРАВКА ====================

/**
 * «Отпечаток» строки — по её собственным ячейкам (Band, ник, языки, продукт). У заказа из нескольких строк
 * номер и тема одинаковые, и только по отпечатку видно, какая это строка, если строки сдвинулись.
 * Статус и сроки сюда не входят: их меняют из списка «Мои», и отпечаток при этом должен остаться тем же.
 */
function keyOf_(raw) {
  const s = [COL.BAND, COL.CUSTOMER, COL.LANGS, COL.PRODUCT].map(c => str_(raw[c - 1])).join('\u0001');
  let h = 5381;
  for (let i = 0; i < s.length; i++) h = ((h * 33) ^ s.charCodeAt(i)) >>> 0;
  return h.toString(36);
}

/**
 * Где сейчас задача: строки могли сдвинуться, пока панель была открыта, — ищем по номеру, теме и отпечатку.
 * Подходит несколько строк, а номер строки уже не тот — не гадаем, а просим открыть задачу заново.
 */
function locate_(t, row, id, origSubject, key) {
  id = str_(id);
  if (!id && origSubject == null) throw new Error('Не знаю, какую задачу искать — откройте её заново.');
  const hasSubject = origSubject != null;
  const exact = k => {
    const r = t.rows[k];
    return (!id || str_(r[COL.ID - 1]) === id) && (!hasSubject || String(r[COL.SUBJECT - 1]) === origSubject) &&
      (!key || keyOf_(t.raw[k]) === key);
  };
  const i0 = Number(row) - 2;
  if (i0 >= 0 && i0 < t.rows.length) {
    anchorsFor_(t, i0); // точные объединения этой строки: номер и тема могут быть в ячейке выше
    if (exact(i0)) return i0;
  }
  const found = [];
  t.rows.forEach((r, k) => { if (exact(k)) found.push(k); });
  if (found.length === 1) return found[0];
  if (found.length > 1) throw new Error('Строки заказа ' + id + ' сдвинулись, пока панель была открыта. Обновите список и откройте задачу заново.');
  if (id && !key) {
    // Тему могли поменять, пока окно было открыто: если строка с таким номером одна — это она
    const byId = [];
    t.rows.forEach((r, k) => { if (str_(r[COL.ID - 1]) === id) byId.push(k); });
    if (byId.length === 1) return byId[0];
  }
  throw new Error('Задача ' + (id || '«' + shortSubject_(origSubject) + '»') + ' не найдена — возможно, её изменили, сдвинули или удалили. Обновите список и откройте её заново.');
}

function searchTasks_(q) {
  q = q.toLowerCase();
  if (!q) return [];
  const t = table_();
  const out = [];
  for (let i = t.rows.length - 1; i >= 0 && out.length < 30; i--) { // новые задачи — внизу листа
    const r = t.rows[i];
    if (!r[COL.SUBJECT - 1]) continue;
    const hay = [r[COL.ID - 1], r[COL.TICKET - 1], r[COL.SUBJECT - 1], r[COL.CUSTOMER - 1], r[COL.COMMENT - 1], r[COL.COMPLAINTS - 1]]
      .map(v => String(v || '').toLowerCase()).join(' ');
    if (hay.indexOf(q) === -1) continue;
    out.push({
      row: i + 2, id: str_(r[COL.ID - 1]), ticket: str_(r[COL.TICKET - 1]), title: shortSubject_(r[COL.SUBJECT - 1]),
      date: fmtDate_(r[COL.DATE - 1], 'dd.MM'), status: str_(r[COL.STATUS - 1]), link: urls_(r[COL.BAND - 1])[0] || '', key: keyOf_(t.raw[i])
    });
  }
  return out;
}

function taskOf_(t, i) {
  anchorsFor_(t, i);
  const r = t.rows[i], sh = t.sh;
  const band = urls_(r[COL.BAND - 1]);
  const subject = String(r[COL.SUBJECT - 1] || '');
  const totalCell = sh.getRange(targetRow_(t, i, COL.TOTAL), COL.TOTAL);
  return {
    row: i + 2, id: str_(r[COL.ID - 1]), ticket: str_(r[COL.TICKET - 1]), subject: subject, origSubject: subject, key: keyOf_(t.raw[i]),
    link: band[0] || '', link2: band.slice(1).join('\n'),
    date: fmtDate_(r[COL.DATE - 1], 'yyyy-MM-dd'), product: str_(r[COL.PRODUCT - 1]), customer: str_(r[COL.CUSTOMER - 1]),
    languages: str_(r[COL.LANGS - 1]), deadline: str_(r[COL.DEADLINE - 1]), exactDeadline: fmtDate_(r[COL.DUE - 1], 'yyyy-MM-dd'),
    status: str_(r[COL.STATUS - 1]),
    estimateLink: estimateUrl_(sh.getRange(targetRow_(t, i, COL.ESTIMATE), COL.ESTIMATE), r[COL.ESTIMATE - 1]),
    total: r[COL.TOTAL - 1] === '' ? '' : r[COL.TOTAL - 1], totalFormula: moneyFormulaOf_(totalCell),
    contractor: str_(r[COL.CONTRACTOR - 1]), manager: str_(r[COL.MANAGER - 1]), deliveryStatus: str_(r[COL.DELIVERY - 1]),
    sp: r[COL.SP - 1] === '' ? '' : r[COL.SP - 1], comment: str_(r[COL.COMMENT - 1]), complaints: str_(r[COL.COMPLAINTS - 1])
  };
}

function getTask_(req) {
  const t = table_();
  const id = str_(req.id);
  const i = id || req.origSubject != null ? locate_(t, req.row, id, req.origSubject, req.key) : Number(req.row) - 2;
  if (!(i >= 0 && i < t.rows.length)) throw new Error('Нет такой строки — найдите задачу заново.');
  return { ok: true, task: taskOf_(t, i) };
}

function myFilter_(filter, status) {
  if (filter === 'done') return status === 'Отдано';
  if (filter === 'cancelled') return status === 'Отменено';
  if (filter === 'all') return true;
  return CLOSED_STATUSES.indexOf(status) === -1;
}

/** Задачи менеджера: 'open' — открытые по сроку; 'done', 'cancelled', 'all' — свежие сверху, не больше 100. */
function myTasks_(manager, filter) {
  filter = ['open', 'done', 'cancelled', 'all'].indexOf(filter) !== -1 ? filter : 'open';
  manager = manager === '*' ? '' : str_(manager);
  const t = table_();
  const l = t.lists;
  if (!manager) return { manager: '', filter: filter, tasks: [], overdue: 0, counts: {}, statuses: [], managers: l.managers };
  const today = today_().getTime();
  const tasks = [];
  const counts = { open: 0, done: 0, cancelled: 0, all: 0 };
  let overdue = 0;
  t.rows.forEach((r, i) => {
    if (!r[COL.SUBJECT - 1] || str_(r[COL.MANAGER - 1]) !== manager) return;
    const status = str_(r[COL.STATUS - 1]);
    const isOpen = CLOSED_STATUSES.indexOf(status) === -1;
    const due = r[COL.DUE - 1];
    const dueTime = isDate_(due) ? due.getTime() : null;
    const late = isOpen && dueTime !== null && dueTime < today;
    ['open', 'done', 'cancelled', 'all'].forEach(f => { if (myFilter_(f, status)) counts[f]++; });
    if (late) overdue++;
    if (!myFilter_(filter, status)) return;
    const date = r[COL.DATE - 1];
    tasks.push({
      row: i + 2, id: str_(r[COL.ID - 1]), subject: String(r[COL.SUBJECT - 1]), title: shortSubject_(r[COL.SUBJECT - 1]),
      ticket: str_(r[COL.TICKET - 1]), status: status, deadline: str_(r[COL.DEADLINE - 1]), delivery: str_(r[COL.DELIVERY - 1]),
      date: fmtDate_(date, 'dd.MM.yy'), due: fmtDate_(due, 'dd.MM'), overdue: late,
      dueToday: isOpen && dueTime !== null && dueTime >= today && dueTime < today + 86400000,
      link: urls_(r[COL.BAND - 1])[0] || '', key: keyOf_(t.raw[i]),
      sort: filter === 'open' ? (dueTime === null ? Infinity : dueTime) : -(isDate_(date) ? date.getTime() : 0) - i / 1e6
    });
  });
  tasks.sort((a, b) => a.sort - b.sort);
  tasks.forEach(x => { delete x.sort; });
  const shown = filter === 'open' ? tasks : tasks.slice(0, MY_TASKS_LIMIT);
  return { manager: manager, filter: filter, tasks: shown, more: filter === 'open' ? 0 : Math.max(0, tasks.length - MY_TASKS_LIMIT),
    counts: counts, overdue: overdue, statuses: l.statuses, managers: l.managers, deliveryStatuses: l.deliveryStatuses };
}

/** Похожие задачи среди последних: та же ссылка на Band или та же тема с тем же тикетом. */
function findDuplicates_(task) {
  const t = table_();
  const subject = shortSubject_(task.subject).toLowerCase();
  const ticket = normTicket_(task.ticket);
  const link = str_(task.link);
  const out = [];
  for (let i = t.rows.length - 1; i >= Math.max(0, t.rows.length - DUP_SCAN_ROWS) && out.length < 5; i--) {
    const r = t.rows[i];
    if (!r[COL.SUBJECT - 1]) continue;
    const sameLink = link && urls_(r[COL.BAND - 1]).indexOf(link) !== -1;
    const sameSubject = subject && shortSubject_(r[COL.SUBJECT - 1]).toLowerCase() === subject && normTicket_(r[COL.TICKET - 1]) === ticket;
    if (sameLink || sameSubject) {
      out.push({ id: str_(r[COL.ID - 1]), title: shortSubject_(r[COL.SUBJECT - 1]), date: fmtDate_(r[COL.DATE - 1], 'dd.MM.yyyy'),
        manager: str_(r[COL.MANAGER - 1]), why: sameLink ? 'та же ссылка на Band' : 'та же тема и тикет' });
    }
  }
  return out;
}

/** Одна ячейка задачи (статус, статус отдачи). */
function setCell_(req, col, value) {
  const t = table_();
  checkHeader_(t.sh);
  const i = locate_(t, req.row, req.id, req.origSubject, req.key);
  anchorsFor_(t, i);
  t.sh.getRange(targetRow_(t, i, col), col).setValue(str_(value));
  return { ok: true, row: i + 2 };
}

/**
 * Правка задачи: пишем ТОЛЬКО те ячейки, что изменились, — остальное в строке не трогаем.
 * Ячейка объединена с верхними (номер, тема, смета заказа) — правим верхнюю.
 */
function saveTask_(req) {
  const task = cleanTask_(req.task);
  const t = table_();
  checkHeader_(t.sh);
  const sh = t.sh;
  const i = locate_(t, req.task.row, req.task.id, req.task.origSubject, req.task.key);
  anchorsFor_(t, i);
  const old = t.rows[i];
  const date = parseIsoDate_(task.date), due = parseIsoDate_(task.exactDeadline);
  // «Итого» не меняли — не разбираем и не трогаем: там бывает текст («-») и суммы с длинным хвостом после запятой
  const totalKept = task.total !== '' && str_(old[COL.TOTAL - 1]) === task.total;
  const totalIn = totalKept ? { value: old[COL.TOTAL - 1] } : moneyInput_(task.total);
  const day = d => fmtDate_(d, 'yyyy-MM-dd');
  const langKey = v => str_(v).split(/\s*,\s*/).filter(Boolean).join('|');
  const same = {
    text: (a, b) => str_(a) === str_(b),
    // дата в ячейке записана текстом — в форму она не попала; пустое поле тогда значит «не меняли», а не «стереть»
    date: (a, b) => (!isDate_(a) && b === '') || (isDate_(a) ? day(a) : str_(a)) === (isDate_(b) ? day(b) : str_(b)),
    num: (a, b) => (a === '' || a == null ? '' : Number(a)) === (b === '' || b == null ? '' : Number(b))
  };
  const cell = c => sh.getRange(targetRow_(t, i, c), c);
  const changed = [];
  const put = (c, kind, value, write) => {
    if (same[kind](old[c - 1], value)) return;
    changed.push(c);
    if (write) write(cell(c)); else cell(c).setValue(value);
  };

  put(COL.PRODUCT, 'text', task.product);
  put(COL.DATE, 'date', date);
  put(COL.SUBJECT, 'text', task.subject);
  if (!same.text(old[COL.TICKET - 1], task.ticket)) requireTicket_(task.ticket); // прежний текст в «Задаче» не трогаем и не проверяем
  put(COL.TICKET, 'text', task.ticket);
  if (urls_(old[COL.BAND - 1]).join('\n') !== task.bandUrls.join('\n')) {
    changed.push(COL.BAND);
    if (task.bandUrls.length) cell(COL.BAND).setRichTextValue(linksRich_(task.bandUrls)); else cell(COL.BAND).setValue('');
  }
  put(COL.CUSTOMER, 'text', task.customer);
  if (langKey(old[COL.LANGS - 1]) !== task.languages.join('|')) put(COL.LANGS, 'text', task.languages.join(', ')); // тот же набор — не переписываем
  // «Дедлайн» (ASAP, 1-2 дня…) пересчитываем, только если поменяли даты: без правок строка остаётся как была
  const datesChanged = !same.date(old[COL.DUE - 1], due) || !same.date(old[COL.DATE - 1], date);
  put(COL.DUE, 'date', due);
  put(COL.DEADLINE, 'text', task.deadline || (datesChanged && deadlineFromDue_(due, date || old[COL.DATE - 1], t.lists)) || str_(old[COL.DEADLINE - 1]));
  put(COL.STATUS, 'text', task.status);
  const oldEstimate = estimateUrl_(cell(COL.ESTIMATE), old[COL.ESTIMATE - 1]);
  if (oldEstimate !== task.estimateLink) { // в колонке сметы бывает и текст («ОТМЕНА») — пока его не меняют, не трогаем
    requireUrl_(task.estimateLink);
    changed.push(COL.ESTIMATE);
    if (task.estimateLink) cell(COL.ESTIMATE).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(task.estimateLink).setLinkUrl(task.estimateLink).build());
    else cell(COL.ESTIMATE).setValue('');
  }
  const oldFormula = moneyFormulaOf_(cell(COL.TOTAL));
  const newFormula = totalIn.formula ? totalIn.formula.replace(/^=/, '') : '';
  // разница меньше десятой доли копейки — это не правка, а хвост после запятой в ячейке
  const sameSum = typeof old[COL.TOTAL - 1] === 'number' && typeof totalIn.value === 'number'
    ? Math.abs(old[COL.TOTAL - 1] - totalIn.value) < 0.0006 : same.num(old[COL.TOTAL - 1], totalIn.value);
  if (!totalKept && (!sameSum || (newFormula && newFormula !== oldFormula))) {
    changed.push(COL.TOTAL);
    const c = cell(COL.TOTAL);
    if (totalIn.formula) c.setFormula(totalIn.formula); else c.setValue(totalIn.value);
    if (totalIn.value !== '') c.setNumberFormat(moneyFormat_(totalIn.value));
  }
  if ((!totalKept && !sameSum) || oldEstimate !== task.estimateLink) {
    noteMoneyEdit_(cell(COL.TOTAL), old[COL.TOTAL - 1], sameSum ? old[COL.TOTAL - 1] : totalIn.value, oldEstimate, task.estimateLink);
  }
  put(COL.CONTRACTOR, 'text', task.contractor);
  put(COL.MANAGER, 'text', task.manager);
  put(COL.DELIVERY, 'text', task.deliveryStatus);
  put(COL.SP, 'num', task.sp === '' ? '' : Number(task.sp));
  put(COL.COMMENT, 'text', task.comment);
  if (req.task.complaints !== undefined) put(COL.COMPLAINTS, 'text', task.complaints);
  return { ok: true, row: i + 2, changed: changed.length };
}

/** Сумму или смету поменяли — прежние остаются в заметке к ячейке суммы (для сверки с подрядчиком). */
function noteMoneyEdit_(cell, oldTotal, newTotal, oldLink, newLink) {
  const lines = [];
  const same = (a, b) => (a === '' ? b === '' : b !== '' && Number(a) === Number(b) && isFinite(Number(a)));
  if (oldTotal !== '' && !same(oldTotal, newTotal)) {
    lines.push(newTotal === '' ? 'сумма удалена, была ' + rub_(oldTotal) + ' ₽' : 'сумма изменена: было ' + rub_(oldTotal) + ' ₽ → стало ' + rub_(newTotal) + ' ₽');
  }
  if (str_(oldLink) && str_(oldLink) !== str_(newLink)) lines.push('прежняя смета: ' + str_(oldLink));
  if (!lines.length) return;
  const line = fmtDate_(new Date(), 'dd.MM.yyyy') + ' — ' + lines.join('\n');
  const note = cell.getNote();
  cell.setNote(note ? note + '\n\n' + line : line);
}

function deleteTask_(req) {
  const t = table_();
  checkHeader_(t.sh);
  const i = locate_(t, req.row, req.id, req.origSubject, req.key);
  // Первая строка заказа из нескольких: в ней номер и тема объединённых ячеек — вместе со строкой они пропадут
  if (i + 1 < t.rows.length && t.head[i + 1] === t.head[i] && t.head[i] === i) {
    throw new Error('Это первая строка заказа из нескольких строк — в ней номер и тема всего заказа. Удалите её в таблице вручную.');
  }
  t.sh.deleteRow(i + 2);
  return { ok: true };
}

// ==================== СМЕТА ====================

/** Заказ по номеру: колонка A; потом номер в любой ячейке строки; потом номер без последней части. */
function findOrder_(t, task) {
  const id = str_(task).toUpperCase();
  if (!id || id.length > 100) throw new Error('Пустой номер задачи');
  const same = (cell, x) => str_(cell).toUpperCase() === x;
  const found = (i, matched) => ({ i: t.head[i], matched: matched });
  for (let i = t.rows.length - 1; i >= 0; i--) if (same(t.raw[i][COL.ID - 1], id)) return found(i, id);
  for (let i = t.rows.length - 1; i >= 0; i--) if (t.raw[i].some(cell => same(cell, id))) return found(i, id);
  const short = id.match(/^(.+)-\d+$/);
  if (short) for (let i = t.rows.length - 1; i >= 0; i--) if (same(t.raw[i][COL.ID - 1], short[1])) return found(i, short[1]);
  return null;
}

function smetaLookup_(req) {
  const t = table_();
  const f = findOrder_(t, req.task);
  if (!f) return { ok: true, found: false };
  anchorsFor_(t, f.i);
  const r = t.rows[f.i];
  return {
    ok: true, found: true, row: f.i + 2, matched: f.matched, subject: str_(r[COL.SUBJECT - 1]),
    contractor: str_(r[COL.CONTRACTOR - 1]), manager: str_(r[COL.MANAGER - 1]),
    link: estimateUrl_(t.sh.getRange(targetRow_(t, f.i, COL.ESTIMATE), COL.ESTIMATE), r[COL.ESTIMATE - 1]), total: r[COL.TOTAL - 1]
  };
}

function smetaWrite_(req) {
  const total = Number(req.total);
  const link = str_(req.link);
  if (!(total > 0)) return { ok: false, error: 'Сумма должна быть больше нуля' };
  if (!/^https?:\/\//i.test(link)) return { ok: false, error: 'Ссылка должна начинаться с http' };

  const t = table_();
  checkHeader_(t.sh);
  const f = findOrder_(t, req.task);
  if (!f) return { ok: false, error: 'Задача ' + req.task + ' не найдена в листе «' + SHEET_NAME + '»' };
  anchorsFor_(t, f.i);
  const r = t.rows[f.i];
  const linkCell = t.sh.getRange(targetRow_(t, f.i, COL.ESTIMATE), COL.ESTIMATE);
  const cell = t.sh.getRange(targetRow_(t, f.i, COL.TOTAL), COL.TOTAL);
  const oldLink = estimateUrl_(linkCell, r[COL.ESTIMATE - 1]);
  const oldTotal = r[COL.TOTAL - 1];
  if ((oldLink || oldTotal !== '') && !req.overwrite) return { ok: false, error: 'exists', link: oldLink, total: oldTotal };

  linkCell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(link).setLinkUrl(link).build());
  const exact = exactMoney_(total);
  cell.setValue(exact).setNumberFormat(moneyFormat_(exact));
  if (req.formula) {
    try { const m = moneyInput_(req.formula); if (m.formula && m.value === exact) cell.setFormula(m.formula); } catch (e) { /* только число */ }
  }
  // Подрядчик обновил смету — прежняя сумма и ссылка остаются в заметке к ячейке
  if (req.overwrite && (oldLink || oldTotal !== '')) {
    const changed = oldTotal === '' || Number(oldTotal) !== exact || (oldLink && oldLink !== link);
    if (changed) {
      const ver = req.version ? ' (версия ' + req.version + ')' : '';
      const sum = oldTotal === '' ? 'сумма ' + rub_(exact) + ' ₽'
        : Number(oldTotal) === exact ? 'сумма та же: ' + rub_(exact) + ' ₽'
        : 'было ' + rub_(oldTotal) + ' ₽ → стало ' + rub_(exact) + ' ₽';
      const line = fmtDate_(new Date(), 'dd.MM.yyyy') + ' — смета обновлена' + ver + ': ' + sum + (oldLink && oldLink !== link ? '\nпрежняя смета: ' + oldLink : '');
      const note = cell.getNote();
      cell.setNote(note ? note + '\n\n' + line : line);
    }
  }
  return { ok: true, row: f.i + 2, url: tableUrl_(t.sh, f.i + 2), was: req.overwrite ? oldTotal : '' };
}
