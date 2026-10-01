/*************************************************************
 * ВсёВОдном.gs — вся таблица задач в одном файле.
 *
 * Как поставить: в таблице «Расширения → Apps Script», удалить все старые файлы,
 * создать один файл-скрипт, вставить этот текст целиком, сохранить (Ctrl+S).
 * Дальше — README, шаги 4–6 (токен, развёртывание, «Оформить таблицу»).
 *
 * Собран автоматически из Код.gs, Перенос.gs, Smeta.gs и HTML-окон (tools/bundle.js).
 * Править лучше исходные файлы и пересобирать.
 *************************************************************/

/*************************************************************
 * Код.gs — таблица задач локализации.
 *
 * Листы:  «📌 Задачи (менеджеры)», «✍️ Задачи (переводчики)», «Списки»,
 *         «Справочник разделов», «📝 Журнал», две инструкции.
 * Окна:   Общее.html (стили и помощники для всех окон) + по файлу на окно.
 * Первый запуск: «⚙️ Настройки → 🎨 Оформить таблицу».
 *************************************************************/

// ==================== ЛИСТЫ И КОЛОНКИ ====================

const TASKS_SHEET = '📌 Задачи (менеджеры)';
const TRANSLATORS_SHEET = '✍️ Задачи (переводчики)';
const LISTS_SHEET = 'Списки';
const SECTIONS_SHEET = 'Справочник разделов';
const LOG_SHEET = '📝 Журнал';
const MANAGER_GUIDE_SHEET = '📋 Инструкция (менеджеры)';
const TRANSLATOR_GUIDE_SHEET = '👥 Инструкция (переводчики)';

// Лист менеджеров, колонки A..Q
const COL = {
  ID: 1, TICKET: 2, SUBJECT: 3, DATE: 4, PRODUCT: 5, CUSTOMER: 6, LANGS: 7, DEADLINE: 8, DUE: 9,
  STATUS: 10, ESTIMATE: 11, TOTAL: 12, CONTRACTOR: 13, MANAGER: 14, DELIVERY: 15, SP: 16, COMMENT: 17
};
const TASK_COLS = 17;

// Лист переводчиков, колонки A..I
const TCOL = { DATE: 1, SIDE: 2, RAZDEL: 3, TASK: 4, CUSTOMER: 5, TRANSLATOR: 6, EDITOR: 7, READY: 8, COMMENT: 9 };
const TR_COLS = 9;

const CLOSED_STATUSES = ['Отдано', 'Отменено'];
const READINESS = ['Не начато', 'В работе', 'Готово', 'На проверке'];
const LINK2_MARKER = ' (доп. ссылка)'; // прежний вид; теперь доп. ссылки — « (ссылка 2)», « (ссылка 3)»…
const LINK_MARKERS_RE = / \((?:доп\. ссылка|ссылка \d+)\)/g;

/** Тема без подписей доп. ссылок. */
function stripLinkMarkers_(s) {
  return String(s == null ? '' : s).replace(LINK_MARKERS_RE, '');
}

/** Доп. ссылки на Band: строка «по одной в строке» (или через пробел) → список адресов. */
function extraLinks_(s) {
  return String(s || '').split(/[\s,;]+/).map(x => x.trim()).filter(x => /^https?:\/\//i.test(x));
}

function getTasksSheet() {
  return findSheet_(TASKS_SHEET, n => n.indexOf('задачи') !== -1 && n.indexOf('перевод') === -1);
}

function getTranslatorsSheet() {
  return findSheet_(TRANSLATORS_SHEET, n => n.indexOf('перевод') !== -1 && n.indexOf('задачи') !== -1);
}

/** Лист по точному имени, а если его переименовали — по части имени. */
function findSheet_(name, test) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (sh) return sh;
  const found = ss.getSheets().filter(s => test(s.getName().toLowerCase()))[0];
  if (found) return found;
  throw new Error('Не нашла лист «' + name + '». Проверьте, что он есть и называется так же.');
}

function listsSheet_() {
  return SpreadsheetApp.getActive().getSheetByName(LISTS_SHEET);
}

/** Сколько строк с данными под шапкой (без лимита в 300 строк). */
function dataRowCount_(sh) {
  return Math.max(sh.getLastRow() - 1, 0);
}

function readRows_(sh, ncols) {
  const n = dataRowCount_(sh);
  return n ? sh.getRange(2, 1, n, ncols).getValues() : [];
}

// ==================== МЕНЮ ====================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 Менеджеры')
    .addItem('➕ Добавить задачу', 'showAddTaskDialog')
    .addItem('🔍 Поиск и правка', 'showSearchEditSidebar')
    .addItem('📊 Дашборд', 'showDashboard')
    .addSeparator()
    .addItem('🙋 Мои задачи', 'showMyTasks')
    .addItem('🙋 Задачи одного менеджера', 'filterByManager')
    .addItem('🙈 Скрыть закрытые (Отдано/Отменено)', 'hideClosedTasks')
    .addItem('👁 Показать все задачи', 'showAllTasks')
    .addToUi();

  ui.createMenu('👥 Переводчики EN')
    .addItem('➕ Добавить задачу переводчика', 'showAddTranslatorTaskDialog')
    .addItem('🔍 Поиск и правка', 'showSearchEditTranslatorSidebar')
    .addSeparator()
    .addItem('🙋 Мои задачи', 'showMyTranslatorTasks')
    .addItem('🙋 Показать одного переводчика', 'filterTranslator')
    .addItem('🙋 Показать одного редактора', 'filterEditor')
    .addItem('🔵 Показать: в работе', 'showInProgressTranslatorTasks')
    .addItem('🟢 Показать: сданные', 'showDoneTranslatorTasks')
    .addItem('👁 Показать все', 'showAllTranslatorTasks')
    .addToUi();

  ui.createMenu('📈 Отчёты')
    .addItem('Отчёт по менеджеру', 'showManagerReportSidebar')
    .addItem('Отчёт по переводчику', 'showTranslatorReportSidebar')
    .addItem('Кастомный отчёт', 'showCustomReportSidebar')
    .addToUi();

  ui.createMenu('⚙️ Настройки')
    .addItem('🎨 Оформить таблицу', 'setupDesign')
    .addItem('🛡 Защитить шапку и справочники', 'protectImportantRanges')
    .addItem('📥 Перенести историю из старой таблицы', 'migrateFromOldTable')
    .addSeparator()
    .addSubMenu(ui.createMenu('✉️ Еженедельная сводка')
      .addItem('Включить (по понедельникам)', 'createWeeklyDigestTrigger')
      .addItem('Тест — прислать сейчас', 'sendWeeklyDigests'))
    .addSubMenu(ui.createMenu('🔥 Горящие дедлайны')
      .addItem('Включить (каждый день)', 'createDailyDueSoonTrigger')
      .addItem('Тест — прислать сейчас', 'sendDueSoonAlerts'))
    .addSubMenu(ui.createMenu('📋 Незаполненные поля')
      .addItem('Включить (26 числа)', 'createMonthEndReminderTrigger')
      .addItem('Тест — прислать сейчас', 'sendMonthEndReminders'))
    .addToUi();

  ensureFilterExists();
}

/**
 * Правки прямо в листе (без окон):
 *  - «Списки» — сразу обновляем цвета (выпадающие списки подтягиваются сами);
 *  - лист задач — языки красятся, новая строка получает выпадающие списки и номер, правка идёт в «Журнал»;
 *  - лист переводчиков — новая строка получает выпадающие списки, правка идёт в «Журнал».
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const name = e.range.getSheet().getName();
  try {
    if (name === LISTS_SHEET) refreshColorRules_();
    else if (name === TASKS_SHEET) onTasksEdit_(e);
    else if (name === TRANSLATORS_SHEET) onTranslatorsEdit_(e);
  } catch (err) {}
}

/** Значение из события правки в читаемый вид (даты приходят числом). */
function editedText_(v, isDate) {
  if (v === undefined || v === null) return '';
  if (isDate && /^\d+(\.\d+)?$/.test(String(v))) return fmtDate_(new Date(Math.round((Number(v) - 25569) * 86400000)), 'dd.MM.yyyy');
  return String(v);
}

function onTasksEdit_(e) {
  const range = e.range, sh = range.getSheet();
  const row = range.getRow(), n = range.getNumRows();
  const col = range.getColumn(), lastCol = col + range.getNumColumns() - 1;
  if (row < 2) return;

  // Новая строка, вписанная руками, — ставим выпадающие списки
  if (!sh.getRange(row, COL.STATUS).getDataValidation()) applyTaskValidations_(sh, row, n);
  if (n > 50) return; // большая вставка — дальше не разбираем построчно

  const values = sh.getRange(row, 1, n, TASK_COLS).getValues();
  values.forEach((r, i) => {
    // Номер задачи, если его нет: как только есть подрядчик и тема
    if (!str_(r[COL.ID - 1]) && str_(r[COL.CONTRACTOR - 1]) && str_(r[COL.SUBJECT - 1])) {
      withScriptLock_(() => {
        const id = generateNextTaskId(sh, str_(r[COL.CONTRACTOR - 1]));
        sh.getRange(row + i, COL.ID).setValue(id);
        r[COL.ID - 1] = id;
        logChange('Создание задачи (в таблице)', id, 'Тема письма', '', str_(r[COL.SUBJECT - 1]));
      });
    }
    if (col <= COL.LANGS && lastCol >= COL.LANGS) colorizeLanguagesCell(sh, row + i);
    // Вставили адрес сметы — показываем коротко «Ссылка на смету номер»
    if (col <= COL.ESTIMATE && lastCol >= COL.ESTIMATE && isUrlText_(r[COL.ESTIMATE - 1])) {
      sh.getRange(row + i, COL.ESTIMATE).setRichTextValue(estimateRich_(r[COL.ESTIMATE - 1], r[COL.ID - 1]));
    }
  });

  // Одна ячейка — пишем в журнал, что было и что стало
  if (n === 1 && col === lastCol && col !== COL.ID) {
    const isDate = col === COL.DATE || col === COL.DUE;
    const before = editedText_(e.oldValue, isDate), after = cellText_(values[0][col - 1]);
    if (before !== after) logChange('Правка в таблице', str_(values[0][COL.ID - 1]), TASK_FIELD_NAMES[col], before, after);
  }
}

function onTranslatorsEdit_(e) {
  const range = e.range, sh = range.getSheet();
  const row = range.getRow(), col = range.getColumn();
  if (row < 2) return;
  if (!sh.getRange(row, TCOL.READY).getDataValidation()) applyTranslatorValidations_(sh, row, range.getNumRows());
  if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
    const r = sh.getRange(row, 1, 1, TR_COLS).getValues()[0];
    const names = ['Дата', 'Сторона', 'Раздел', 'Задача', 'Заказчик', 'Переводчик', 'Редактор', 'Готовность', 'Комментарий'];
    const before = editedText_(e.oldValue, col === TCOL.DATE), after = cellText_(r[col - 1]);
    if (before !== after) logChange('Правка в таблице (переводчик)', str_(r[TCOL.TASK - 1]), names[col - 1], before, after);
  }
}

// Старые пункты меню — теперь всё делает «Оформить таблицу».
function repaintAllRows() { setupDesign(); }
function repaintAllTranslatorRows() { setupDesign(); }
function applyAllValidations() { setupDesign(); }
function applyAllTranslatorValidations() { setupDesign(); }

// ==================== ОБЩИЕ ПОМОЩНИКИ ====================

/**
 * Окна берутся из HTML-файлов проекта. В сборке «всё в одном файле» (ВсёВОдном.gs) их текст
 * лежит в HTML_FILES — тогда отдельные HTML-файлы не нужны.
 */
function htmlSource_(name) {
  return typeof HTML_FILES !== 'undefined' && HTML_FILES[name] != null ? HTML_FILES[name] : null;
}

/** Адреса смет, вписанные текстом (в т. ч. перенесённые), → «Ссылка на смету номер» со ссылкой. */
function estimateLinksToLabels_(sh) {
  const n = dataRowCount_(sh);
  if (!n) return;
  const range = sh.getRange(2, COL.ESTIMATE, n, 1);
  const values = range.getValues();
  // Голый адрес или подпись прежнего вида («Смета [номер]») → «Ссылка на смету номер»
  if (!values.some(r => isUrlText_(r[0]) || (ESTIMATE_LABEL_RE.test(str_(r[0])) && !/^Ссылка на смету/.test(str_(r[0]))))) return;
  const ids = sh.getRange(2, COL.ID, n, 1).getValues();
  const rich = range.getRichTextValues();
  range.setRichTextValues(values.map((r, i) => {
    const url = estimateUrl_(rich[i][0], r[0]);
    if (isUrlText_(url) && (isUrlText_(r[0]) || ESTIMATE_LABEL_RE.test(str_(r[0])))) return [estimateRich_(url, ids[i][0])];
    return [rich[i][0] || SpreadsheetApp.newRichTextValue().setText(cellText_(r[0])).build()];
  }));
}

/** Подключает Общее.html в окна: <?!= include('Общее') ?> */
function include(name) {
  const src = htmlSource_(name);
  return src != null ? src : HtmlService.createHtmlOutputFromFile(name).getContent();
}

function template_(file) {
  const src = htmlSource_(file);
  return src != null ? HtmlService.createTemplate(src) : HtmlService.createTemplateFromFile(file);
}

function showDialog_(file, title, width, height) {
  const html = template_(file).evaluate().setWidth(width).setHeight(height);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

function showSidebar_(file, title) {
  const html = template_(file).evaluate().setTitle(title);
  SpreadsheetApp.getUi().showSidebar(html);
}

function tz_() {
  return Session.getScriptTimeZone();
}

function fmtDate_(d, pattern) {
  return d instanceof Date ? Utilities.formatDate(d, tz_(), pattern) : '';
}

/** «2026-10-01» → дата в часовом поясе таблицы (без сдвига на сутки). */
function parseIsoDate_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : '';
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function str_(v) {
  return String(v == null ? '' : v).trim();
}

/** «[LIT-8][LogrusIT][az]… Новые строчки» → «Новые строчки». */
function shortSubject_(s) {
  return stripLinkMarkers_(s).replace(/^(\s*\[[^\]]*\])+\s*/, '').trim() || String(s || '');
}

/** «1234», «local 1234», «LOCAL-1234» → «LOCAL-1234». */
function normTicket_(t) {
  const s = String(t || '').trim().toUpperCase().replace(/\s+/g, '');
  if (!s) return '';
  const m = s.match(/^(?:LOCAL-?)?(\d+)$/);
  return m ? 'LOCAL-' + m[1] : s;
}

// ==================== БЛОКИРОВКА ====================

// Выставляет Smeta.gs, когда уже держит блокировку (запрос из расширения).
var SCRIPT_LOCK_HELD = false;

/** Чтобы два человека одновременно не получили один номер и не правили одну строку. */
function withScriptLock_(fn) {
  if (SCRIPT_LOCK_HELD) return fn();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  SCRIPT_LOCK_HELD = true;
  try {
    return fn();
  } finally {
    SCRIPT_LOCK_HELD = false;
    lock.releaseLock();
  }
}

// ==================== ССЫЛКИ В ТЕМЕ ====================

function linksFromRich_(rt) {
  if (!rt) return { link: '', link2: '' };
  const urls = [];
  const whole = rt.getLinkUrl();
  if (whole) urls.push(whole);
  rt.getRuns().forEach(run => {
    const u = run.getLinkUrl();
    if (u && urls.indexOf(u) === -1) urls.push(u);
  });
  // link — основная (на теме), link2 — доп. ссылки по одной в строке
  const hasMarker = /\((?:доп\. ссылка|ссылка \d+)\)/.test(rt.getText());
  return { link: urls[0] || '', link2: hasMarker ? urls.slice(1).join('\n') : '' };
}

// Смета: в ячейке короткая ссылка «Ссылка на смету LIT-26-2232» вместо длинного адреса
const ESTIMATE_LABEL_RE = /^(Смета|Ссылка на смету)( \[[^\]]*\]| \S+)?$/; // и прежний вид «Смета [номер]»
const isUrlText_ = v => /^https?:\/\/\S+$/i.test(str_(v));

function estimateRich_(url, id) {
  url = str_(url);
  const label = url ? 'Ссылка на смету' + (str_(id) ? ' ' + str_(id) : '') : '';
  const b = SpreadsheetApp.newRichTextValue().setText(label);
  return (url ? b.setLinkUrl(url) : b).build();
}

/** Адрес сметы из ячейки: ссылка под «Ссылка на смету …» или сам текст, если там адрес. */
function estimateUrl_(rich, value) {
  const link = linksFromRich_(rich).link;
  if (link) return link;
  const v = str_(value);
  return ESTIMATE_LABEL_RE.test(v) ? '' : v;
}

/** Ссылки всех строк одним запросом (раньше — по запросу на строку). */
function subjectLinks_(sh, col) {
  const n = dataRowCount_(sh);
  if (!n) return [];
  return sh.getRange(2, col || COL.SUBJECT, n, 1).getRichTextValues().map(r => linksFromRich_(r[0]));
}

function getSubjectLink(sh, row) {
  return linksFromRich_(sh.getRange(row, COL.SUBJECT).getRichTextValue()).link;
}

function getSubjectLink2(sh, row) {
  return linksFromRich_(sh.getRange(row, COL.SUBJECT).getRichTextValue()).link2;
}

/** Тема со вшитой основной ссылкой и после неё «(ссылка 2)», «(ссылка 3)»… — доп. ссылки на Band. */
function buildSubjectRich_(text, link, link2) {
  text = stripLinkMarkers_(text);
  const extra = text ? extraLinks_(link2) : [];
  const marks = extra.map((u, i) => ' (ссылка ' + (i + 2) + ')');
  const full = text + marks.join('');
  let b = SpreadsheetApp.newRichTextValue().setText(full);
  if (text && link) b = b.setLinkUrl(0, text.length, link);
  let pos = text.length;
  extra.forEach((u, i) => {
    b = b.setLinkUrl(pos + 1, pos + marks[i].length, u); // без пробела перед скобкой
    pos += marks[i].length;
  });
  return b.build();
}

// ==================== СПРАВОЧНИКИ ====================

/** Все справочники с листа «Списки» одним запросом. */
function getListsData() {
  const lst = listsSheet_();
  const ncols = Math.min(15, lst.getMaxColumns()); // O может ещё не быть — его добавит «Оформить таблицу»
  const raw = lst.getRange(3, 1, 200, ncols).getValues().map(r => r.concat(Array(15 - ncols).fill(''))); // A..O
  function col(idx) {
    const items = [];
    for (let i = 0; i < raw.length; i++) {
      const v = raw[i][idx];
      if (v === '' || v === null) break;
      items.push(String(v).trim());
    }
    return items;
  }
  const contractors = col(6);
  return {
    products: col(0), langs: col(1), deadlines: col(2), statuses: col(3), managers: col(4),
    contractors: contractors, deliveryStatuses: col(7), translators: col(9),
    prefixes: contractors.map((c, i) => String(raw[i][14] || '').trim()), // O = Префикс подрядчика
    langCodes: langCodesFrom_(col(1), raw)
  };
}

// Коды языков для темы письма. Лист «Списки», колонка I «Код языка» (рядом с языком в B), главнее;
// это — запасной вариант. Новый язык: впишите его в B, а код — в I той же строки.
const LANG_CODES = {
  'Азербайджанский': 'az', 'Английский': 'en', 'Амхарский': 'am', 'Армянский': 'hy',
  'Грузинский': 'ka', 'Казахский': 'kk', 'Китайский': 'zhs', 'Кыргызский': 'ky',
  'Русский': 'ru', 'Таджикский': 'tg', 'Узбекский': 'uz', 'Суахили': 'sw',
  'Белорусский': 'be', 'Иврит': 'he', 'Корейский': 'ko', 'Турецкий': 'tr', 'Арабский': 'ar', 'Французский': 'fr'
};

function langCodesFrom_(langs, raw) {
  const codes = {};
  Object.keys(LANG_CODES).forEach(l => { codes[l] = LANG_CODES[l]; });
  langs.forEach((l, i) => {
    const c = String(raw[i][8] || '').trim(); // I = Код языка
    if (c) codes[l] = c;
  });
  return codes;
}
const RARE_LANGS = ['Белорусский', 'Иврит', 'Корейский', 'Турецкий', 'Арабский', 'Французский'];

// Префиксы номеров задач. Лист «Списки», колонка O, главнее; это — запасной вариант.
const CONTRACTOR_PREFIX = {
  'LogrusIT': 'LIT', 'LogrusGlobal': 'LoG', 'Бюро переводов': 'BP', 'JanusWW': 'JWW', 'Awatera': 'AWT'
};

function contractorPrefix_(contractor, lists) {
  lists = lists || getListsData();
  const i = lists.contractors.indexOf(contractor);
  return (i !== -1 && lists.prefixes[i]) || CONTRACTOR_PREFIX[contractor] || 'TASK';
}

// Разделы переводчиков — запасной вариант, если листа «Справочник разделов» нет.
// Им же пользуется TranslationStatus.gs.
const CONTENT_CATEGORIES = {
  "Портал продавца": ["А/Б-тесты", "Авторизация", "Админка команды транзитов", "Админка перемещения остатков", "Админка портала", "Админка справочного центра", "Админка стран", "Админка управления поставщиками", "Админки коммуникаций", "Акции: отзывы и вайбы", "Аналитика", "Анрег", "Биржа карточек", "Бренды", "Витрина магазина", "Возвраты", "Восстановление доступа", "Главная", "Данные", "Джем", "Документы", "Закреплённые отзывы", "Защита контента", "История платежей", "История платежей за сервисы", "Календарь акций", "Калькулятор прибыли", "Карточка товара", "Каталог решений для бизнеса", "Конструктор тарифов", "Лендинг", "Лендинг WB Банка", "Настройки", "Обращения правообладателей", "Онбординг", "Опросы", "Оферта продавца", "Перенос карточек товаров", "Поддержка (Диалоги)", "Подменные артикулы", "Поставки FBS", "Поставки FBW", "Почтовая рассылка", "Программа лояльности", "Профиль производителя автозапчастей", "Регистрация", "Рекомендации", "Сервис новостей", "Сервис отзывов и вопросов", "Сервис уведомлений", "Сервисы", "Системные уведомления", "Справочный центр", "Тарифы", "Тематики для поддержки", "Финансовые отчёты", "Финансовые отчёты - баланс", "Цены и скидки", "Цифровой арбитраж", "Чат с покупателями", "API", "Core", "WB для бизнеса", "WB Продвижение", "WB Legal"],
  "Магазинка": ["Конфиг", "Разное", "Платёжный шлюз", "Подборки", "Сервис заказов", "Системные пуши", "Статика", "WBX Поиск"],
  "WB Клуб": ["WB Клуб"],
  "WB ID": ["WB ID"],
  "WB ПВЗ": ["Бот-помощник", "Брендбук", "Карта WB ПВЗ", "Озвучка", "Поддержка ПВЗ", "Сайт WB ПВЗ", "WB Мой ПВЗ", "WB ПВЗ десктоп", "WB ПВЗ моб", "WB ПВЗ Core"],
  "WB Rent": ["WB Rent"],
  "Поддержка продавцов": ["Поддержка продавцов"],
  "Поддержка покупателей": ["Поддержка покупателей", "Чат-бот"],
  "Справочный центр": ["Статьи СЦ"],
  "Коммуникации": ["Баннеры", "Новость"],
  "Маркетинг": ["Маркетинговые пуши", "Прочее", "Пуши", "Рекламные баннеры", "Сторы"],
  "Контент": ["Контент", "Меню, категории, характеристики"],
  "WB Банк": ["WB Банк"],
  "WB Логистика": ["Портал логистики"],
  "WB Такси": ["WB Такси"],
  "WB Taxi": ["WB Taxi"],
  "WB Tracker": ["WB Tracker"],
  "Нестандартное": ["Нестандартное"]
};

/** Стороны и разделы с листа «Справочник разделов»: строка 2 — стороны, ниже — разделы. */
function getContentCategories() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SECTIONS_SHEET);
  if (!sh || sh.getLastRow() < 2) return CONTENT_CATEGORIES;
  const lastCol = sh.getLastColumn();
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();
  const out = {};
  for (let c = 0; c < lastCol; c++) {
    const side = String(data[0][c] || '').trim();
    if (!side) continue;
    const items = [];
    for (let r = 1; r < data.length; r++) {
      const v = String(data[r][c] || '').trim();
      if (v && items.indexOf(v) === -1) items.push(v);
    }
    out[side] = items.length ? items : [side];
  }
  return Object.keys(out).length ? out : CONTENT_CATEGORIES;
}

/** Имя менеджера по почте того, кто открыл таблицу («Списки»: E — имя, F — почта). */
function currentManager_() {
  return personByEmail_(5);
}

/** Имя переводчика по почте («Списки»: J — имя, K — почта). */
function currentTranslator_() {
  return personByEmail_(10);
}

function personByEmail_(nameCol) {
  let email = '';
  try { email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) {}
  if (!email) return '';
  const rows = listsSheet_().getRange(3, nameCol, 200, 2).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (!rows[i][0]) break;
    if (String(rows[i][1] || '').trim().toLowerCase() === email) return String(rows[i][0]).trim();
  }
  return '';
}

// ==================== ЦВЕТА ====================

const HEAD_BG = '#2D3340';
const INK = '#1F2328';
const FONT = 'Roboto'; // родной шрифт Google Таблиц

// Светлые пары [фон, текст] — продукты, менеджеры, стороны, переводчики
const BADGE_PALETTE = [
  ['#E0F2FF', '#0B5394'], ['#E2F0D9', '#38761D'], ['#FCE5CD', '#B45F06'],
  ['#EAD1F0', '#A64D79'], ['#FFF2CC', '#7F6000'], ['#D9D2E9', '#674EA7'],
  ['#F4CCCC', '#990000'], ['#D0E0E3', '#134F5C'], ['#EFEFEF', '#434343'],
  ['#D9EAD3', '#274E13'], ['#FCE8B2', '#7F6000'], ['#F9CBDB', '#741B47']
];
// Менеджеры — та же светлая палитра, но с другого места, чтобы не совпадать с продуктами
const MANAGER_PALETTE = BADGE_PALETTE.slice(11).concat(BADGE_PALETTE.slice(0, 11));
// Подрядчики — тёмные плашки с белым текстом: их не спутать с менеджерами
const CONTRACTOR_PALETTE = [
  ['#1F4E79', '#FFFFFF'], ['#2E6B3F', '#FFFFFF'], ['#8E4B10', '#FFFFFF'],
  ['#5B2C83', '#FFFFFF'], ['#455A64', '#FFFFFF'], ['#7A1F3D', '#FFFFFF']
];
const DEADLINE_COLORS = {
  'ASAP': ['#F4CCCC', '#990000'], '1-2 дня': ['#FCE5CD', '#B45F06'],
  'До недели': ['#FFF2CC', '#7F6000'], 'Больше недели': ['#D9EAD3', '#38761D'],
  'Месяц и больше': ['#D6E4F0', '#0B5394'], 'Холд': ['#EFEFEF', '#434343']
};
const STATUS_COLORS = {
  'Принято': ['#EDEDED', '#202124'], 'В работе': ['#D6E4F0', '#0B5394'],
  'Отдано': ['#D9EAD3', '#274E13'], 'Отменено': ['#F4CCCC', '#990000'], 'Холд': ['#FFF2CC', '#7F6000']
};
const DELIVERY_STATUS_COLORS = {
  'Отдано в срок': ['#D9EAD3', '#274E13'],
  'Сообщили о просрочке': ['#FFF2CC', '#7F6000'],
  'Не сообщили о просрочке': ['#F4CCCC', '#990000']
};
const READINESS_COLORS = {
  'Не начато': ['#EFEFEF', '#434343'], 'В работе': ['#D6E4F0', '#0B5394'],
  'Готово': ['#D9EAD3', '#274E13'], 'На проверке': ['#FFF2CC', '#7F6000']
};
const OVERDUE_BG = '#FDEDEA';
const OVERDUE_INK = '#B42318';

// Языки: у каждого свой цвет букв, ШТАТ — того же цвета, что и основной язык
const LANG_COLORS = {
  'Азербайджанский': '#B45F06', 'Амхарский': '#38761D', 'Английский': '#1155CC', 'Армянский': '#A64D79',
  'Грузинский': '#CC0000', 'Казахский': '#0B5394', 'Кыргызский': '#BF9000', 'Таджикский': '#674EA7',
  'Узбекский': '#134F5C', 'Китайский': '#990000', 'Русский': '#333333', 'Суахили': '#6AA84F',
  'Белорусский': '#3D85C6', 'Иврит': '#7F6000', 'Корейский': '#45818E', 'Турецкий': '#E69138',
  'Арабский': '#274E13', 'Французский': '#351C75'
};

function colorForLang(lang) {
  const base = String(lang).replace(/^ШТАТ\s+/i, '');
  const key = Object.keys(LANG_COLORS).filter(k => k.toLowerCase() === base.toLowerCase())[0];
  return LANG_COLORS[key] || '#5F6368';
}

function colorForItem(item, list, palette) {
  const i = list.indexOf(item);
  if (i === -1) return null;
  palette = palette || BADGE_PALETTE;
  return palette[i % palette.length];
}

function langsRich_(text) {
  const langs = String(text || '').split(',').map(s => s.trim()).filter(Boolean);
  const joined = langs.join(', ');
  const b = SpreadsheetApp.newRichTextValue().setText(joined);
  let pos = 0;
  langs.forEach(l => {
    b.setTextStyle(pos, pos + l.length, SpreadsheetApp.newTextStyle()
      .setForegroundColor(colorForLang(l)).setBold(true).build());
    pos += l.length + 2;
  });
  return b.build();
}

/**
 * Красит языки в ячейке «Языки» (каждый своим цветом).
 * Если в колонке включены «чипы» (выпадающий список), красить не нужно — цвет дают чипы.
 */
function colorizeLanguagesCell(sh, row) {
  const cell = sh.getRange(row, COL.LANGS);
  if (cell.getDataValidation()) return;
  const text = String(cell.getValue() || '');
  if (!text.trim()) return;
  cell.setRichTextValue(langsRich_(text));
}

/**
 * Раньше красил ячейки строки вручную. На оформленных листах цвета ставят правила
 * условного форматирования, и они меняются сами. Для листов без правил (например,
 * Translation Status) красим по-старому.
 */
function colorizeRowDirectly(sh, row, listsData) {
  if (sh.getConditionalFormatRules().length) return;
  const lists = listsData || getListsData();
  const v = sh.getRange(row, 1, 1, TASK_COLS).getValues()[0];
  function paint(col, pair) {
    if (pair) sh.getRange(row, col).setBackground(pair[0]).setFontColor(pair[1]).setFontWeight('bold');
  }
  paint(COL.PRODUCT, colorForItem(v[COL.PRODUCT - 1], lists.products));
  paint(COL.DEADLINE, DEADLINE_COLORS[v[COL.DEADLINE - 1]]);
  paint(COL.STATUS, STATUS_COLORS[v[COL.STATUS - 1]]);
  paint(COL.CONTRACTOR, colorForItem(v[COL.CONTRACTOR - 1], lists.contractors, CONTRACTOR_PALETTE));
  paint(COL.MANAGER, colorForItem(v[COL.MANAGER - 1], lists.managers, MANAGER_PALETTE));
  paint(COL.DELIVERY, DELIVERY_STATUS_COLORS[v[COL.DELIVERY - 1]]);
}

// ==================== ВЫПАДАЮЩИЕ СПИСКИ ====================

// Какие колонки берут значения из какой колонки «Списков»
const TASK_LIST_SOURCES = [
  [COL.PRODUCT, 'A'], [COL.DEADLINE, 'C'], [COL.STATUS, 'D'],
  [COL.CONTRACTOR, 'G'], [COL.MANAGER, 'E'], [COL.DELIVERY, 'H']
];

function listRule_(range) {
  return SpreadsheetApp.newDataValidation().requireValueInRange(range, true).setAllowInvalid(true).build();
}

function dateRule_() {
  return SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
}

/**
 * Выпадающие списки в строках row..row+n-1. Значения берутся прямо из «Списков»:
 * дописали нового менеджера — он сразу есть в списке.
 */
function applyTaskValidations_(sh, row, n) {
  const lst = listsSheet_();
  TASK_LIST_SOURCES.forEach(([col, letter]) => {
    sh.getRange(row, col, n, 1).setDataValidation(listRule_(lst.getRange(letter + '3:' + letter + '200')));
  });
  sh.getRange(row, COL.DATE, n, 1).setDataValidation(dateRule_());
  sh.getRange(row, COL.DUE, n, 1).setDataValidation(dateRule_());
}

/** Для одной строки (им же пользуется TranslationStatus.gs). */
function applyRowValidations(sh, row) {
  applyTaskValidations_(sh, row, 1);
}

function applyTranslatorValidations_(sh, row, n) {
  const lst = listsSheet_();
  const people = lst.getRange('J3:J200');
  sh.getRange(row, TCOL.TRANSLATOR, n, 1).setDataValidation(listRule_(people));
  sh.getRange(row, TCOL.EDITOR, n, 1).setDataValidation(listRule_(people));
  sh.getRange(row, TCOL.READY, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(READINESS, true).setAllowInvalid(true).build());
  const ref = SpreadsheetApp.getActive().getSheetByName(SECTIONS_SHEET);
  if (ref) sh.getRange(row, TCOL.SIDE, n, 1).setDataValidation(listRule_(ref.getRange(2, 1, 1, ref.getMaxColumns())));
  sh.getRange(row, TCOL.DATE, n, 1).setDataValidation(dateRule_());
}

/**
 * Новая строка сверху (под шапкой). Формат и выпадающие списки копируем со строки ниже,
 * чтобы новая строка выглядела как остальные (и сохранялись «чипы», если их включили вручную).
 */
function insertTopRow_(sh, ncols, fallback) {
  sh.insertRowAfter(1);
  if (sh.getLastRow() >= 3) {
    const src = sh.getRange(3, 1, 1, ncols);
    const dst = sh.getRange(2, 1, 1, ncols);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  } else {
    sh.getRange(2, 1, 1, ncols).setBackground(null).setFontColor(INK).setFontWeight('normal')
      .setFontFamily(FONT).setFontSize(10).setVerticalAlignment('middle').setWrap(true);
    if (sh.getName() === TASKS_SHEET) sh.getRange(2, COL.SUBJECT).setFontWeight('bold');
    fallback();
  }
  return 2;
}

// ==================== ФИЛЬТРЫ ====================

function ensureFilterExists() {
  try {
    const sh = getTasksSheet();
    if (!sh.getFilter()) sh.getDataRange().createFilter();
  } catch (e) {}
}

function resetFilter_(sh) {
  const existing = sh.getFilter();
  if (existing) existing.remove();
  return sh.getDataRange().createFilter();
}

function hideClosedTasks() {
  const filter = resetFilter_(getTasksSheet());
  filter.setColumnFilterCriteria(COL.STATUS, SpreadsheetApp.newFilterCriteria().setHiddenValues(CLOSED_STATUSES).build());
  SpreadsheetApp.getActive().toast('Закрытые задачи скрыты. Вернуть: «👁 Показать все задачи».', 'Фильтр', 6);
}

function showAllTasks() {
  resetFilter_(getTasksSheet());
  SpreadsheetApp.getActive().toast('Показаны все задачи');
}

function filterTasksByManager_(name) {
  const sh = getTasksSheet();
  const values = readRows_(sh, TASK_COLS).map(r => str_(r[COL.MANAGER - 1])).filter(Boolean);
  const hidden = values.filter((v, i) => v !== name && values.indexOf(v) === i);
  const filter = resetFilter_(sh);
  if (hidden.length) filter.setColumnFilterCriteria(COL.MANAGER, SpreadsheetApp.newFilterCriteria().setHiddenValues(hidden).build());
  SpreadsheetApp.getActive().toast('Показаны задачи: ' + name + '. Вернуть: «👁 Показать все задачи».', 'Фильтр', 6);
}

function filterByManager() {
  const managers = getListsData().managers;
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Задачи одного менеджера', 'Впишите имя как в списке:\n\n' + managers.join(', '), ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const name = resp.getResponseText().trim();
  if (managers.indexOf(name) === -1) { ui.alert('Не нашла такого менеджера в «Списках». Проверьте написание.'); return; }
  filterTasksByManager_(name);
}

function showMyTasks() {
  const me = currentManager_();
  if (!me) {
    SpreadsheetApp.getUi().alert(
      'Не нашла вашу почту в списке менеджеров.\n' +
      'Попросите вписать её на листе «Списки», колонка F — тогда «Мои задачи» заработают.\n' +
      'Пока можно выбрать себя через «Задачи одного менеджера».');
    return;
  }
  filterTasksByManager_(me);
}

// ==================== НОМЕР ЗАДАЧИ ====================

// Нумерация продолжается со старой таблицы: последние номера оттуда (на 30.09.2026).
// Берётся больший из двух: это число или последний номер на листе.
const TASK_ID_START = { 'LIT-26': 2231, 'BP-26': 366, 'LoG-26': 75, 'JWW-26': 29, 'AWT-26': 1 };

// Старая таблица, в которой тоже заводят задачи. Её запоминает перенос истории,
// и новые номера сверяются с ней — чтобы две таблицы не выдали один и тот же номер.
const OLD_SHEET_NAME = 'Localization Misc';
const OLD_TABLE_PROP = 'OLD_TABLE_ID';

/** Последний номер вида base-N в старой таблице (запоминается на минуту, чтобы не открывать её каждый раз). */
function oldTableMaxSeq_(base) {
  let id = '';
  try { id = PropertiesService.getDocumentProperties().getProperty(OLD_TABLE_PROP) || ''; } catch (e) {}
  if (!id) return 0;
  let cache = null;
  try { cache = CacheService.getDocumentCache(); } catch (e) {}
  const key = 'oldmax:' + base;
  const cached = cache && cache.get(key);
  if (cached) return Number(cached);
  let max = 0;
  try {
    const old = SpreadsheetApp.openById(id).getSheetByName(OLD_SHEET_NAME);
    const re = new RegExp('^' + base + '-(\\d+)$');
    old.getRange(1, 1, Math.max(old.getLastRow(), 1), 1).getValues().forEach(r => {
      const m = String(r[0]).trim().match(re);
      if (m) max = Math.max(max, Number(m[1]));
    });
  } catch (e) {
    // старая таблица недоступна (или правка прямо в листе, где её открывать нельзя) — считаем только по новой
  }
  // Открывать старую таблицу долго — раз в 6 часов достаточно: новые номера всё равно сверяются с новой таблицей
  if (cache) cache.put(key, String(max), 21600);
  return max;
}

/** Номер вида ПОДРЯДЧИК-ГГ-N (LIT-26-2232), как в старой таблице и в именах смет. */
function generateNextTaskId(sh, contractor) {
  const base = contractorPrefix_(contractor) + '-' + Utilities.formatDate(new Date(), tz_(), 'yy');
  const n = dataRowCount_(sh);
  const values = n ? sh.getRange(2, COL.ID, n, 1).getValues() : [];
  let maxSeq = Math.max(TASK_ID_START[base] || 0, oldTableMaxSeq_(base));
  const re = new RegExp('^' + base + '-(\\d+)$');
  values.forEach(r => {
    const m = String(r[0]).trim().match(re);
    if (m) maxSeq = Math.max(maxSeq, Number(m[1]));
  });
  return base + '-' + (maxSeq + 1);
}

function previewNextTaskId(contractor) {
  return contractor ? generateNextTaskId(getTasksSheet(), contractor) : '';
}

/** Следующие номера сразу для всех подрядчиков — расширение берёт их одним запросом при открытии формы. */
function previewNextTaskIds() {
  const sh = getTasksSheet();
  const out = {};
  getListsData().contractors.forEach(c => { out[c] = generateNextTaskId(sh, c); });
  return out;
}

// ==================== ДОБАВИТЬ ЗАДАЧУ ====================

function showAddTaskDialog() {
  showDialog_('AddTaskDialog', 'Новая задача', 540, 720);
}

function getAddTaskFormLists() {
  const lists = getListsData();
  const shtat = lists.langs.filter(l => /^ШТАТ/i.test(l));
  const rare = lists.langs.filter(l => RARE_LANGS.indexOf(l) !== -1);
  const regular = lists.langs.filter(l => shtat.indexOf(l) === -1 && rare.indexOf(l) === -1);
  return {
    contractors: lists.contractors, products: lists.products, deadlines: lists.deadlines,
    statuses: lists.statuses, deliveryStatuses: lists.deliveryStatuses, managers: lists.managers,
    languages: { regular: regular, shtat: shtat, rare: rare },
    langCodes: lists.langCodes,
    currentManager: currentManager_(),
    templates: getTaskTemplates()
  };
}

// ==================== ШАБЛОНЫ ЗАДАЧ ====================
// Свои у каждого: хранятся в личных настройках скрипта (видит только тот, кто сохранил).
// Тот же формат файла, что у расширения: { type: 'wb-task-templates', templates: [...] }.
const TPL_PROP = 'TASK_TEMPLATES';
const TPL_CHUNK = 8000; // одно значение в Properties — до 9 КБ

function cleanTemplate_(t) {
  t = t || {};
  const out = { name: str_(t.name) };
  ['contractor', 'ticket', 'subject', 'product', 'customer', 'deadline', 'comment'].forEach(k => {
    out[k] = typeof t[k] === 'string' ? t[k].trim() : '';
  });
  out.languages = Array.isArray(t.languages) ? t.languages.filter(l => typeof l === 'string') : [];
  return out;
}

function getTaskTemplates() {
  const props = PropertiesService.getUserProperties();
  const n = Number(props.getProperty(TPL_PROP + '_N') || 0);
  let json = '';
  for (let i = 0; i < n; i++) json += props.getProperty(TPL_PROP + '_' + i) || '';
  try {
    return json ? JSON.parse(json).map(cleanTemplate_).filter(t => t.name) : [];
  } catch (e) {
    return [];
  }
}

/** Сохраняет весь список шаблонов (окно присылает его целиком после правки). */
function saveTaskTemplates(list) {
  const clean = (list || []).map(cleanTemplate_).filter(t => t.name);
  const json = JSON.stringify(clean);
  const props = PropertiesService.getUserProperties();
  const old = Number(props.getProperty(TPL_PROP + '_N') || 0);
  const n = Math.ceil(json.length / TPL_CHUNK);
  const values = {};
  for (let i = 0; i < n; i++) values[TPL_PROP + '_' + i] = json.slice(i * TPL_CHUNK, (i + 1) * TPL_CHUNK);
  values[TPL_PROP + '_N'] = String(n);
  props.setProperties(values);
  for (let i = n; i < old; i++) props.deleteProperty(TPL_PROP + '_' + i);
  return clean;
}

/** Тема письма: [номер][подрядчик][коды языков][продукт] суть. */
function buildTaskSubject_(id, task, lists) {
  const langCodes = (lists || getListsData()).langCodes;
  const codes = (task.languages || []).map(l => langCodes[l]).filter(Boolean);
  const prefix = '[' + id + ']' + (task.contractor ? '[' + task.contractor + ']' : '') +
    codes.map(c => '[' + c + ']').join('') + (task.product ? '[' + task.product + ']' : '');
  const raw = str_(task.subject);
  return raw ? prefix + ' ' + raw : prefix;
}

function submitNewTaskFromDialog(task) {
  ['contractor', 'product', 'deadline', 'status', 'deliveryStatus', 'manager', 'subject', 'customer', 'comment', 'link', 'link2', 'estimateLink']
    .forEach(k => { task[k] = str_(task[k]); });
  if (!task.contractor) throw new Error('Выберите подрядчика');

  return withScriptLock_(() => {
    const sh = getTasksSheet();
    const id = generateNextTaskId(sh, task.contractor);
    const subject = buildTaskSubject_(id, task);
    const row = insertTopRow_(sh, TASK_COLS, () => applyTaskValidations_(sh, 2, 1));

    const values = [
      id, normTicket_(task.ticket), subject,
      parseIsoDate_(task.date), task.product, task.customer, (task.languages || []).join(', '),
      task.deadline, parseIsoDate_(task.exactDeadline), task.status,
      task.estimateLink, task.total ? Number(task.total) : '', task.contractor, task.manager,
      task.deliveryStatus, task.sp ? Number(task.sp) : '', task.comment
    ];
    sh.getRange(row, 1, 1, TASK_COLS).setValues([values]);
    if (task.link || task.link2) sh.getRange(row, COL.SUBJECT).setRichTextValue(buildSubjectRich_(subject, task.link, task.link2));
    if (task.estimateLink) sh.getRange(row, COL.ESTIMATE).setRichTextValue(estimateRich_(task.estimateLink, id));
    if (values[COL.LANGS - 1]) colorizeLanguagesCell(sh, row);
    colorizeRowDirectly(sh, row);

    logChange('Создание задачи', id, 'Тема письма', '', subject);
    return id;
  });
}

// ==================== ПОИСК И ПРАВКА (менеджеры) ====================

function showSearchEditSidebar() {
  showSidebar_('SearchEditSidebar', 'Поиск и правка');
}

/**
 * Поиск по номеру, тикету, теме, нику и комментарию.
 * Пустой запрос — ваши задачи (или все последние, если почты нет в «Списках»).
 */
function searchTasks(query) {
  const sh = getTasksSheet();
  const rows = readRows_(sh, TASK_COLS);
  const q = str_(query).toLowerCase();
  const me = q ? '' : currentManager_();
  const out = [];
  rows.forEach((r, i) => {
    if (!r[COL.SUBJECT - 1]) return;
    if (q) {
      const hay = [r[COL.ID - 1], r[COL.TICKET - 1], r[COL.SUBJECT - 1], r[COL.CUSTOMER - 1], r[COL.COMMENT - 1]]
        .map(v => String(v || '').toLowerCase()).join(' ');
      if (hay.indexOf(q) === -1) return;
    } else if (me && str_(r[COL.MANAGER - 1]) !== me) {
      return;
    }
    out.push({
      row: i + 2, id: str_(r[COL.ID - 1]), ticket: str_(r[COL.TICKET - 1]),
      title: shortSubject_(r[COL.SUBJECT - 1]), date: fmtDate_(r[COL.DATE - 1], 'dd.MM'),
      status: str_(r[COL.STATUS - 1])
    });
  });
  return out.slice(0, 30);
}

function getTaskForEdit(row) {
  const sh = getTasksSheet();
  const r = sh.getRange(row, 1, 1, TASK_COLS).getValues()[0];
  const links = linksFromRich_(sh.getRange(row, COL.SUBJECT).getRichTextValue());
  const subject = String(r[COL.SUBJECT - 1] || '');
  return {
    row: row, id: str_(r[COL.ID - 1]), ticket: str_(r[COL.TICKET - 1]),
    subject: stripLinkMarkers_(subject), origSubject: subject,
    link: links.link, link2: links.link2,
    date: fmtDate_(r[COL.DATE - 1], 'yyyy-MM-dd'), product: str_(r[COL.PRODUCT - 1]),
    customer: str_(r[COL.CUSTOMER - 1]), languages: str_(r[COL.LANGS - 1]),
    deadline: str_(r[COL.DEADLINE - 1]), exactDeadline: fmtDate_(r[COL.DUE - 1], 'yyyy-MM-dd'),
    status: str_(r[COL.STATUS - 1]), estimateLink: estimateUrl_(sh.getRange(row, COL.ESTIMATE).getRichTextValue(), r[COL.ESTIMATE - 1]),
    total: r[COL.TOTAL - 1] === '' ? '' : r[COL.TOTAL - 1], contractor: str_(r[COL.CONTRACTOR - 1]),
    manager: str_(r[COL.MANAGER - 1]), deliveryStatus: str_(r[COL.DELIVERY - 1]),
    sp: r[COL.SP - 1] === '' ? '' : r[COL.SP - 1], comment: str_(r[COL.COMMENT - 1])
  };
}

/**
 * Где сейчас задача. Пока окно было открыто, кто-то мог добавить или удалить задачи,
 * и строки сдвинулись — поэтому ищем по номеру и теме. В старых задачах у одного номера
 * бывает несколько строк (несколько запросов в одном письме) — тогда тема решает, какая строка.
 */
function locateTaskRow_(sh, row, id, origSubject) {
  if (!id && origSubject == null) throw new Error('Не знаю, какую задачу искать — откройте её заново.');
  const hasSubject = origSubject != null;
  const exact = r => (!id || str_(r[COL.ID - 1]) === id) && (!hasSubject || String(r[COL.SUBJECT - 1]) === origSubject);
  const n = dataRowCount_(sh);
  if (row >= 2 && row <= n + 1 && exact(sh.getRange(row, 1, 1, TASK_COLS).getValues()[0])) return row;
  const rows = readRows_(sh, TASK_COLS);
  const i = rows.findIndex(exact);
  if (i !== -1) return i + 2;
  if (id) {
    // Тему могли поменять, пока окно было открыто: если строка с таким номером одна — это она
    const byId = rows.map((r, k) => str_(r[COL.ID - 1]) === id ? k : -1).filter(k => k !== -1);
    if (byId.length === 1) return byId[0] + 2;
  }
  throw new Error('Задача ' + (id || '«' + shortSubject_(origSubject) + '»') + ' не найдена — возможно, её удалили или изменили. Найдите её заново.');
}

const TASK_FIELD_NAMES = {};
[['ID', '№ задачи'], ['TICKET', 'Тикет'], ['SUBJECT', 'Тема письма'], ['DATE', 'Дата'], ['PRODUCT', 'Продукт'],
 ['CUSTOMER', 'Ник заказчика'], ['LANGS', 'Языки'], ['DEADLINE', 'Дедлайн'], ['DUE', 'Срок сдачи'], ['STATUS', 'Статус'],
 ['ESTIMATE', 'Смета'], ['TOTAL', 'Итого с НДС'], ['CONTRACTOR', 'Подрядчик'], ['MANAGER', 'Менеджер'],
 ['DELIVERY', 'Статус отдачи'], ['SP', 'SP'], ['COMMENT', 'Комментарий']]
  .forEach(([k, name]) => { TASK_FIELD_NAMES[COL[k]] = name; });

function cellText_(v) {
  return v instanceof Date ? fmtDate_(v, 'dd.MM.yyyy') : String(v == null ? '' : v);
}

function saveTaskEdits(task) {
  return withScriptLock_(() => {
    const sh = getTasksSheet();
    const row = locateTaskRow_(sh, Number(task.row), str_(task.id), task.origSubject);
    const old = sh.getRange(row, 1, 1, TASK_COLS).getValues()[0];
    old[COL.ESTIMATE - 1] = estimateUrl_(sh.getRange(row, COL.ESTIMATE).getRichTextValue(), old[COL.ESTIMATE - 1]);
    const next = old.slice();
    const set = (col, v) => { next[col - 1] = v; };
    set(COL.TICKET, normTicket_(task.ticket));
    set(COL.DATE, parseIsoDate_(task.date));
    set(COL.PRODUCT, str_(task.product));
    set(COL.CUSTOMER, str_(task.customer));
    set(COL.LANGS, str_(task.languages));
    set(COL.DEADLINE, str_(task.deadline));
    set(COL.DUE, parseIsoDate_(task.exactDeadline));
    set(COL.STATUS, str_(task.status));
    set(COL.ESTIMATE, str_(task.estimateLink));
    set(COL.TOTAL, task.total === '' || task.total == null ? '' : Number(task.total));
    set(COL.CONTRACTOR, str_(task.contractor));
    set(COL.MANAGER, str_(task.manager));
    set(COL.DELIVERY, str_(task.deliveryStatus));
    set(COL.SP, task.sp === '' || task.sp == null ? '' : Number(task.sp));
    set(COL.COMMENT, str_(task.comment));

    const id = str_(old[COL.ID - 1]);
    for (let c = COL.TICKET; c <= TASK_COLS; c++) {
      if (c === COL.SUBJECT) continue;
      if (cellText_(old[c - 1]) !== cellText_(next[c - 1])) {
        logChange('Правка', id, TASK_FIELD_NAMES[c], cellText_(old[c - 1]), cellText_(next[c - 1]));
      }
    }
    sh.getRange(row, COL.TICKET).setValue(next[COL.TICKET - 1]);
    sh.getRange(row, COL.DATE, 1, TASK_COLS - COL.DATE + 1).setValues([next.slice(COL.DATE - 1)]);
    if (next[COL.ESTIMATE - 1]) sh.getRange(row, COL.ESTIMATE).setRichTextValue(estimateRich_(next[COL.ESTIMATE - 1], id));

    // Тема со ссылками: пишем, только если что-то поменялось, чтобы не трогать лишнее
    const subject = str_(task.subject);
    const oldLinks = linksFromRich_(sh.getRange(row, COL.SUBJECT).getRichTextValue());
    const oldSubject = stripLinkMarkers_(old[COL.SUBJECT - 1]);
    const newExtra = extraLinks_(task.link2).join('\n');
    if (subject !== oldSubject || str_(task.link) !== oldLinks.link || newExtra !== oldLinks.link2) {
      sh.getRange(row, COL.SUBJECT).setRichTextValue(buildSubjectRich_(subject, str_(task.link), str_(task.link2)));
      logChange('Правка', id, 'Тема письма', oldSubject, subject);
    }
    colorizeLanguagesCell(sh, row);
    return { ok: true, row: row };
  });
}

function deleteTask(row, id, origSubject) {
  return withScriptLock_(() => {
    const sh = getTasksSheet();
    const r = locateTaskRow_(sh, Number(row), str_(id), origSubject);
    const subject = sh.getRange(r, COL.SUBJECT).getValue();
    logChange('Удаление задачи', str_(id), 'Тема письма', subject, '');
    sh.deleteRow(r);
    return true;
  });
}

// ==================== ПРОВЕРКА НА ДУБЛИ ====================

/**
 * Похожие задачи: та же ссылка на Band или та же тема с тем же тикетом.
 * Окно «Новая задача» и расширение спрашивают перед добавлением.
 */
// Дубли ищем среди последних задач (новые — сверху): читать ссылки у всех тысяч строк долго
const DUP_SCAN_ROWS = 800;
function findDuplicateTasks(task) {
  const sh = getTasksSheet();
  const n = Math.min(dataRowCount_(sh), DUP_SCAN_ROWS);
  const rows = n ? sh.getRange(2, 1, n, TASK_COLS).getValues() : [];
  const links = n ? sh.getRange(2, COL.SUBJECT, n, 1).getRichTextValues().map(r => linksFromRich_(r[0])) : [];
  const subject = shortSubject_(task.subject).toLowerCase();
  const ticket = normTicket_(task.ticket);
  const link = str_(task.link);
  const out = [];
  rows.forEach((r, i) => {
    if (!r[COL.SUBJECT - 1] || out.length >= 5) return;
    const sameLink = link && links[i] && (links[i].link === link || extraLinks_(links[i].link2).indexOf(link) !== -1);
    const sameSubject = subject && shortSubject_(r[COL.SUBJECT - 1]).toLowerCase() === subject &&
      normTicket_(r[COL.TICKET - 1]) === ticket;
    if (sameLink || sameSubject) {
      out.push({ id: str_(r[COL.ID - 1]), title: shortSubject_(r[COL.SUBJECT - 1]), date: fmtDate_(r[COL.DATE - 1], 'dd.MM.yyyy'),
        manager: str_(r[COL.MANAGER - 1]), why: sameLink ? 'та же ссылка на Band' : 'та же тема и тикет' });
    }
  });
  return out;
}

// ==================== МОИ ЗАДАЧИ (для расширения) ====================

/** Открытые задачи менеджера: сначала просроченные, потом по сроку. */
/**
 * Задачи менеджера для расширения. filter: 'open' (по умолчанию) — открытые, по сроку;
 * 'done' — отданные, 'cancelled' — отменённые, 'all' — все; закрытые и «все» — свежие сверху, не больше 100.
 * counts — сколько задач в каждом фильтре; overdue — просрочки среди открытых (для бейджа).
 */
const MY_TASKS_LIMIT = 100;
function myTaskFilter_(filter, status) {
  if (filter === 'done') return status === 'Отдано';
  if (filter === 'cancelled') return status === 'Отменено';
  if (filter === 'all') return true;
  return CLOSED_STATUSES.indexOf(status) === -1;
}

function getMyOpenTasks(manager, filter) {
  filter = ['open', 'done', 'cancelled', 'all'].indexOf(filter) !== -1 ? filter : 'open';
  manager = manager === '*' ? '' : (str_(manager) || currentManager_()); // «*» — пока не выбран
  if (!manager) return { manager: '', filter: filter, tasks: [], overdue: 0, counts: {}, statuses: [], managers: getListsData().managers };
  const sh = getTasksSheet();
  const rows = readRows_(sh, TASK_COLS);
  const today = today_();
  const tasks = [];
  const counts = { open: 0, done: 0, cancelled: 0, all: 0 };
  let overdue = 0;
  rows.forEach((r, i) => {
    if (!r[COL.SUBJECT - 1] || str_(r[COL.MANAGER - 1]) !== manager) return;
    const status = str_(r[COL.STATUS - 1]);
    const isOpen = CLOSED_STATUSES.indexOf(status) === -1;
    const due = r[COL.DUE - 1];
    const dueTime = due instanceof Date ? due.getTime() : null;
    const late = isOpen && dueTime !== null && dueTime < today.getTime();
    ['open', 'done', 'cancelled', 'all'].forEach(f => { if (myTaskFilter_(f, status)) counts[f]++; });
    if (late) overdue++;
    if (!myTaskFilter_(filter, status)) return;
    const date = r[COL.DATE - 1];
    tasks.push({
      row: i + 2, id: str_(r[COL.ID - 1]), subject: String(r[COL.SUBJECT - 1]), title: shortSubject_(r[COL.SUBJECT - 1]),
      ticket: str_(r[COL.TICKET - 1]), status: status, deadline: str_(r[COL.DEADLINE - 1]),
      date: fmtDate_(date, 'dd.MM.yy'), due: fmtDate_(due, 'dd.MM'), overdue: late,
      dueToday: isOpen && dueTime !== null && dueTime >= today.getTime() && dueTime < today.getTime() + 86400000,
      sort: filter === 'open' ? (dueTime === null ? Infinity : dueTime) : -(date instanceof Date ? date.getTime() : 0)
    });
  });
  tasks.sort((a, b) => a.sort - b.sort);
  tasks.forEach(t => { delete t.sort; });
  const shown = filter === 'open' ? tasks : tasks.slice(0, MY_TASKS_LIMIT);
  // Ссылки на Band — только для показанных строк: читать оформление темы у всех тысяч строк долго
  if (shown.length) {
    const from = Math.min.apply(null, shown.map(t => t.row)), to = Math.max.apply(null, shown.map(t => t.row));
    const rich = sh.getRange(from, COL.SUBJECT, to - from + 1, 1).getRichTextValues();
    shown.forEach(t => { t.link = linksFromRich_(rich[t.row - from][0]).link; });
  }
  const lists = getListsData();
  return { manager: manager, filter: filter, tasks: shown,
    more: filter === 'open' ? 0 : Math.max(0, tasks.length - MY_TASKS_LIMIT),
    counts: counts, overdue: overdue, statuses: lists.statuses, managers: lists.managers };
}

/** Поменять только статус (быстрая кнопка в расширении). */
function setTaskStatus(row, id, origSubject, status) {
  return withScriptLock_(() => {
    const sh = getTasksSheet();
    const r = locateTaskRow_(sh, Number(row), str_(id), origSubject);
    const cell = sh.getRange(r, COL.STATUS);
    const before = str_(cell.getValue());
    cell.setValue(str_(status));
    if (before !== str_(status)) logChange('Правка', str_(id), 'Статус', before, str_(status));
    return { ok: true, row: r };
  });
}

// ==================== ЖУРНАЛ ====================

// Кто действует, если это не человек в таблице (например, расширение). Заполняет Smeta.gs.
var LOG_ACTOR = '';

function logChange(action, refId, field, oldVal, newVal) {
  try {
    const log = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET);
    let who = LOG_ACTOR;
    if (!who) { try { who = Session.getActiveUser().getEmail(); } catch (e) {} }
    log.appendRow([new Date(), who || '(неизвестно)', action, refId, field || '',
      oldVal == null ? '' : oldVal, newVal == null ? '' : newVal]);
  } catch (e) {}
}

function getTaskHistory(refId) {
  if (!refId) return [];
  try {
    const log = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET);
    const data = log.getDataRange().getValues();
    return data.filter((r, i) => i > 0 && str_(r[3]) === refId && r[0] instanceof Date)
      .sort((a, b) => b[0] - a[0]).slice(0, 10)
      .map(r => ({ date: fmtDate_(r[0], 'dd.MM.yy HH:mm'), who: str_(r[1]), action: str_(r[2]), field: str_(r[4]) }));
  } catch (e) {
    return [];
  }
}

// ==================== ЗАЩИТА ====================

function protectImportantRanges() {
  getTasksSheet().getRange(1, 1, 1, TASK_COLS).protect().setDescription('Шапка листа задач').setWarningOnly(true);
  listsSheet_().getDataRange().protect().setDescription('Справочники (Списки)').setWarningOnly(true);
  SpreadsheetApp.getUi().alert('Готово: шапка и лист «Списки» защищены — при правке будет предупреждение.');
}

// ==================== ДАШБОРД ====================

function showDashboard() {
  showDialog_('DashboardSidebar', 'Дашборд', 780, 700);
}

function getDashboardYears() {
  const years = [];
  readRows_(getTasksSheet(), COL.DATE).forEach(r => {
    const d = r[COL.DATE - 1];
    if (d instanceof Date && years.indexOf(d.getFullYear()) === -1) years.push(d.getFullYear());
  });
  return years.sort((a, b) => b - a);
}

function inPeriod_(date, year, month) {
  if (!year) return true;
  if (!(date instanceof Date) || date.getFullYear() !== year) return false;
  return !month || date.getMonth() + 1 === month;
}

function isOverdue_(r, today) {
  const due = r[COL.DUE - 1];
  return due instanceof Date && due < today && CLOSED_STATUSES.indexOf(str_(r[COL.STATUS - 1])) === -1 && r[COL.SUBJECT - 1];
}

function today_() {
  const t = new Date();
  t.setHours(0, 0, 0, 0);
  return t;
}

function getOverdueCount() {
  const today = today_();
  return readRows_(getTasksSheet(), TASK_COLS).filter(r => isOverdue_(r, today)).length;
}

function getDashboardData(yearFilter, monthFilter) {
  const sh = getTasksSheet();
  const data = readRows_(sh, TASK_COLS);
  const year = yearFilter ? Number(yearFilter) : null;
  const month = monthFilter ? Number(monthFilter) : null;
  const lists = getListsData();
  const today = today_();

  let totalTasks = 0, totalSp = 0, totalMoney = 0;
  const bySP = {}, byMoney = {}, byCount = {}, byContractorMoney = {}, langCounts = {};
  lists.langs.forEach(l => { langCounts[l] = 0; });
  const overdue = [];

  data.forEach(r => {
    if (!r[COL.SUBJECT - 1]) return;
    if (isOverdue_(r, today)) {
      overdue.push({ id: str_(r[COL.ID - 1]), title: shortSubject_(r[COL.SUBJECT - 1]), due: fmtDate_(r[COL.DUE - 1], 'dd.MM'), manager: str_(r[COL.MANAGER - 1]) });
    }
    if (!inPeriod_(r[COL.DATE - 1], year, month)) return;
    totalTasks++;
    const manager = str_(r[COL.MANAGER - 1]), product = str_(r[COL.PRODUCT - 1]), contractor = str_(r[COL.CONTRACTOR - 1]);
    const sp = Number(r[COL.SP - 1]) || 0, money = Number(r[COL.TOTAL - 1]) || 0;
    totalSp += sp; totalMoney += money;
    if (manager) { bySP[manager] = (bySP[manager] || 0) + sp; byMoney[manager] = (byMoney[manager] || 0) + money; }
    if (product) byCount[product] = (byCount[product] || 0) + 1;
    if (contractor) byContractorMoney[contractor] = (byContractorMoney[contractor] || 0) + money;
    const langText = String(r[COL.LANGS - 1] || '');
    lists.langs.forEach(l => { if (langText.indexOf(l) !== -1) langCounts[l]++; });
  });

  const monthNames = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
  const now = new Date();
  const monthKeys = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    monthKeys.push({ label: monthNames[d.getMonth()] + ' ' + d.getFullYear(), y: d.getFullYear(), m: d.getMonth(), tasks: 0, sp: 0 });
  }
  const byYear = {};
  data.forEach(r => {
    const date = r[COL.DATE - 1];
    if (!r[COL.SUBJECT - 1] || !(date instanceof Date)) return;
    const sp = Number(r[COL.SP - 1]) || 0; // раньше здесь по ошибке бралась колонка «Статус отдачи»
    const y = date.getFullYear();
    byYear[y] = byYear[y] || [String(y), 0, 0];
    byYear[y][1]++; byYear[y][2] += sp;
    const mk = monthKeys.filter(k => k.y === y && k.m === date.getMonth())[0];
    if (mk) { mk.tasks++; mk.sp += sp; }
  });

  return {
    totalTasks: totalTasks, totalSp: totalSp, totalMoney: Math.round(totalMoney * 100) / 100,
    overdueCount: overdue.length, overdue: overdue.slice(0, 10),
    managers: lists.managers.map(m => [m, bySP[m] || 0]),
    managersMoney: lists.managers.map(m => [m, byMoney[m] || 0]),
    products: lists.products.map(p => [p, byCount[p] || 0]),
    contractorsMoney: lists.contractors.map(c => [c, byContractorMoney[c] || 0]),
    languages: lists.langs.map(l => [l, langCounts[l]]),
    byMonth: monthKeys.map(k => [k.label, k.tasks, k.sp]),
    byYear: Object.keys(byYear).sort().map(y => byYear[y])
  };
}

// ==================== ОТЧЁТ ПО МЕНЕДЖЕРУ ====================

function showManagerReportSidebar() {
  showDialog_('ManagerReportSidebar', 'Отчёт по менеджеру', 560, 760);
}

function getManagersList() {
  return getListsData().managers;
}

function getMonthsList() {
  const names = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  const now = new Date();
  const months = [];
  for (let i = 0; i < 18; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push({ value: d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'), label: names[d.getMonth()] + ' ' + d.getFullYear() });
  }
  return months;
}

function parseMonth_(monthStr) {
  const m = String(monthStr || '').match(/^(\d{4})-(\d{2})$/);
  return m ? { year: Number(m[1]), month: Number(m[2]) } : { year: null, month: null };
}

function getManagerReport(manager, monthStr) {
  const sh = getTasksSheet();
  const data = readRows_(sh, TASK_COLS);
  const links = subjectLinks_(sh);
  const p = parseMonth_(monthStr);

  const groups = {}, order = [];
  let taskCount = 0, lineCount = 0, grandTotal = 0, grandMoney = 0;
  data.forEach((r, i) => {
    const subject = r[COL.SUBJECT - 1];
    if (str_(r[COL.MANAGER - 1]) !== manager || !subject) return;
    if (p.year && !inPeriod_(r[COL.DATE - 1], p.year, p.month)) return;
    const ticket = str_(r[COL.TICKET - 1]);
    const key = ticket || '(без тикета)';
    const sp = Number(r[COL.SP - 1]) || 0, money = Number(r[COL.TOTAL - 1]) || 0;
    if (!groups[key]) { groups[key] = { ticket: ticket, lines: [], ticketTotal: 0, ticketMoneyTotal: 0 }; order.push(key); taskCount++; }
    groups[key].lines.push({ subject: stripLinkMarkers_(subject), link: links[i].link || '', link2: links[i].link2 || '', sp: sp, money: money });
    groups[key].ticketTotal += sp;
    groups[key].ticketMoneyTotal += money;
    grandTotal += sp; grandMoney += money; lineCount++;
  });

  const label = (getMonthsList().filter(m => m.value === monthStr)[0] || {}).label || monthStr || '';
  const list = order.map(k => groups[k]);
  let copyText = '';
  list.forEach(g => {
    copyText += (g.ticket || '(без тикета)') + '\n';
    g.lines.forEach(l => { copyText += l.subject + ' — ' + l.sp + ' SP' + (l.link2 ? ' (доп. ссылки: ' + extraLinks_(l.link2).join(', ') + ')' : '') + '\n'; });
    copyText += 'Итого по тикету: ' + g.ticketTotal + ' SP\n\n';
  });
  copyText += 'Общий итог по ' + manager + ': ' + grandTotal + ' SP';

  return { manager: manager, monthLabel: label, taskCount: taskCount, lineCount: lineCount,
    grandTotal: grandTotal, grandMoney: grandMoney, groups: list, copyText: copyText };
}

/** Отчёт в новую Google Таблицу; возвращает ссылку на скачивание .xlsx. */
function exportReportToDoc(manager, monthStr) {
  const report = getManagerReport(manager, monthStr);
  const title = 'Отчёт по SP — ' + manager + (report.monthLabel ? ' — ' + report.monthLabel : '');
  const ss = SpreadsheetApp.create(title);
  const sheet = ss.getSheets()[0].setName('Отчёт');
  const rows = [['Тикет', 'Тема', 'SP', 'Доп. ссылки']];
  const rich = [];
  report.groups.forEach(g => {
    g.lines.forEach(l => { rows.push([g.ticket || '(без тикета)', l.subject, l.sp, extraLinks_(l.link2).join('\n')]); rich.push([rows.length, l.subject, l.link]); });
    rows.push(['', 'Итого по тикету', g.ticketTotal, '']);
    rows.push(['', '', '', '']);
  });
  rows.push(['ОБЩИЙ ИТОГ по ' + manager, '', report.grandTotal, '']);
  sheet.getRange(1, 1, rows.length, 4).setValues(rows);
  rich.forEach(([row, text, link]) => {
    if (link && text) sheet.getRange(row, 2).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(link).build());
  });
  styleReportSheet_(sheet, 4, 400);
  return exportUrl_(ss);
}

function styleReportSheet_(sheet, ncols, wideColWidth) {
  sheet.getRange(1, 1, 1, ncols).setFontWeight('bold').setBackground(HEAD_BG).setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, ncols);
  if (wideColWidth) sheet.setColumnWidth(2, wideColWidth);
}

function exportUrl_(ss) {
  SpreadsheetApp.flush();
  return 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
}

// ==================== ОТЧЁТ ДЛЯ ТРЕКЕРА (менеджеры) ====================

function getTicketsForReport() {
  const out = [];
  readRows_(getTasksSheet(), COL.TICKET).forEach(r => {
    const t = str_(r[COL.TICKET - 1]);
    if (t && out.indexOf(t) === -1) out.push(t);
  });
  return out;
}

function getTrackerReportText(ticket) {
  const sh = getTasksSheet();
  const data = readRows_(sh, TASK_COLS);
  const links = subjectLinks_(sh);
  const rows = [];
  data.forEach((r, i) => {
    if (str_(r[COL.TICKET - 1]) !== ticket || !r[COL.SUBJECT - 1]) return;
    const label = '[' + str_(r[COL.ID - 1]) + '] ' + shortSubject_(r[COL.SUBJECT - 1]);
    rows.push({ label: label, link: links[i].link, sp: r[COL.SP - 1] || 0 });
  });
  if (!rows.length) return { text: '', count: 0 };
  let text = 'Задачи в Band:\n\n';
  rows.forEach((r, i) => { text += (i + 1) + '. ' + (r.link ? '[' + r.label + '](' + r.link + ')' : r.label) + ' — ' + r.sp + ' SP\n'; });
  return { text: text.trim(), count: rows.length };
}

// ==================== КАСТОМНЫЙ ОТЧЁТ ====================

function showCustomReportSidebar() {
  showDialog_('CustomReportSidebar', 'Кастомный отчёт', 600, 760);
}

function getCustomReport(yearFilter, monthFilter) {
  const sh = getTasksSheet();
  const data = readRows_(sh, TASK_COLS);
  const links = subjectLinks_(sh);
  const year = yearFilter ? Number(yearFilter) : null;
  const month = monthFilter ? Number(monthFilter) : null;

  let totalTasks = 0, totalSp = 0, totalMoney = 0;
  const byProduct = {}, byLang = {}, byDelivery = {}, byManager = {}, byContractor = {};
  const taskRows = [];
  data.forEach((r, i) => {
    if (!r[COL.SUBJECT - 1] || !inPeriod_(r[COL.DATE - 1], year, month)) return;
    totalTasks++;
    const product = str_(r[COL.PRODUCT - 1]), delivery = str_(r[COL.DELIVERY - 1]) || '(не указано)';
    const manager = str_(r[COL.MANAGER - 1]), contractor = str_(r[COL.CONTRACTOR - 1]);
    const sp = Number(r[COL.SP - 1]) || 0, money = Number(r[COL.TOTAL - 1]) || 0;
    totalSp += sp; totalMoney += money;
    if (product) byProduct[product] = (byProduct[product] || 0) + 1;
    String(r[COL.LANGS - 1] || '').split(',').map(s => s.trim()).filter(Boolean).forEach(l => { byLang[l] = (byLang[l] || 0) + 1; });
    byDelivery[delivery] = (byDelivery[delivery] || 0) + 1;
    if (manager) {
      byManager[manager] = byManager[manager] || { count: 0, sp: 0 };
      byManager[manager].count++; byManager[manager].sp += sp;
    }
    if (contractor) byContractor[contractor] = (byContractor[contractor] || 0) + 1;
    taskRows.push({ subject: stripLinkMarkers_(r[COL.SUBJECT - 1]), link: links[i].link, sp: sp, manager: manager, product: product });
  });

  const sortDesc = obj => Object.keys(obj).map(k => [k, obj[k]]).sort((a, b) => b[1] - a[1]);
  const monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  return {
    periodLabel: year && month ? monthNames[month - 1] + ' ' + year : year ? String(year) : 'Всё время',
    totalTasks: totalTasks, totalSp: totalSp, totalMoney: Math.round(totalMoney * 100) / 100,
    avgSp: totalTasks ? Math.round(totalSp / totalTasks * 10) / 10 : 0,
    topProducts: sortDesc(byProduct), languages: sortDesc(byLang),
    deliveryStatus: Object.keys(byDelivery).map(k => [k, byDelivery[k], totalTasks ? Math.round(byDelivery[k] / totalTasks * 100) : 0]),
    topManagers: Object.keys(byManager).map(k => [k, byManager[k].count, byManager[k].sp]).sort((a, b) => b[2] - a[2]).slice(0, 5),
    topContractors: sortDesc(byContractor).slice(0, 5),
    topTasksBySP: taskRows.filter(t => t.sp > 0).sort((a, b) => b.sp - a.sp).slice(0, 10)
  };
}

function exportCustomReportToExcel(yearFilter, monthFilter) {
  const report = getCustomReport(yearFilter, monthFilter);
  const title = 'Отчёт для руководства — ' + report.periodLabel;
  const ss = SpreadsheetApp.create(title);
  const sheet = ss.getSheets()[0].setName('Отчёт');
  const rows = [];
  const heads = [];
  const add = (a, b, c) => rows.push([a, b === undefined ? '' : b, c === undefined ? '' : c]);
  const header = text => { heads.push(rows.length + 1); add(text); };

  header(title);
  add('Всего задач', report.totalTasks); add('Итого SP', report.totalSp);
  add('Итого ₽', report.totalMoney); add('Средний SP на задачу', report.avgSp); add('');
  header('Продукты по числу задач'); report.topProducts.forEach(r => add(r[0], r[1])); add('');
  header('Языки по частоте'); report.languages.forEach(r => add(r[0], r[1])); add('');
  header('Статус отдачи заказчику'); add('Статус', 'Кол-во', '%'); report.deliveryStatus.forEach(r => add(r[0], r[1], r[2] + '%')); add('');
  header('Топ-5 менеджеров по SP'); add('Менеджер', 'Задач', 'SP'); report.topManagers.forEach(r => add(r[0], r[1], r[2])); add('');
  header('Топ-5 подрядчиков по числу задач'); report.topContractors.forEach(r => add(r[0], r[1])); add('');
  header('Топ-10 самых трудоёмких задач (по SP)'); add('Тема', 'Менеджер', 'SP');
  const linkRows = [];
  report.topTasksBySP.forEach(t => { add(t.subject, t.manager, t.sp); linkRows.push([rows.length, t.subject, t.link]); });

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  heads.forEach(r => sheet.getRange(r, 1, 1, 3).setFontWeight('bold').setBackground(HEAD_BG).setFontColor('#ffffff'));
  linkRows.forEach(([r, text, link]) => {
    if (link && text) sheet.getRange(r, 1).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(link).build());
  });
  sheet.autoResizeColumns(1, 3);
  sheet.setColumnWidth(1, 380);
  return exportUrl_(ss);
}

function exportCustomReportToDoc(yearFilter, monthFilter) {
  const report = getCustomReport(yearFilter, monthFilter);
  const title = 'Отчёт для руководства — ' + report.periodLabel;
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(50).setMarginRight(50);
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('Сформировано: ' + fmtDate_(new Date(), 'dd.MM.yyyy')).setForegroundColor('#666666');

  function section(heading, rows, headerRow) {
    body.appendParagraph(heading).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const table = body.appendTable(headerRow ? [headerRow].concat(rows) : rows);
    styleDocTable_(table, !!headerRow);
    body.appendParagraph('');
    return table;
  }
  section('Общая картина', [['Всего задач', String(report.totalTasks)], ['Итого SP', String(report.totalSp)],
    ['Итого ₽', String(report.totalMoney)], ['Средний SP на задачу', String(report.avgSp)]]);
  section('Продукты по числу задач', report.topProducts.map(r => [r[0], String(r[1])]));
  section('Языки по частоте', report.languages.map(r => [r[0], String(r[1])]));
  section('Статус отдачи заказчику', report.deliveryStatus.map(r => [r[0], String(r[1]), r[2] + '%']), ['Статус', 'Кол-во', '%']);
  section('Топ-5 менеджеров по SP', report.topManagers.map(r => [r[0], String(r[1]), String(r[2])]), ['Менеджер', 'Задач', 'SP']);
  section('Топ-5 подрядчиков по числу задач', report.topContractors.map(r => [r[0], String(r[1])]));
  const taskTable = section('Топ-10 самых трудоёмких задач (по SP)',
    report.topTasksBySP.map(t => [t.subject, t.manager || '', String(t.sp)]), ['Тема', 'Менеджер', 'SP']);
  report.topTasksBySP.forEach((t, i) => {
    if (!t.link) return;
    const text = taskTable.getCell(i + 1, 0).editAsText();
    if (text.getText().length) text.setLinkUrl(0, text.getText().length - 1, t.link);
  });
  doc.saveAndClose();
  return doc.getUrl();
}

function styleDocTable_(table, hasHeader) {
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(8).setPaddingRight(8);
      if (hasHeader && r === 0) {
        cell.setBackgroundColor(HEAD_BG);
        cell.editAsText().setForegroundColor('#ffffff').setBold(true);
      }
    }
  }
}

// ==================== ПИСЬМА ====================

const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1ntXZIzlVB0zaoExpWNbB0hlSPD_KI-vHoKs3TEiYz3Q/edit';

/** Менеджеры и почты для рассылки («Списки»: E — имя, M — почта для рассылки). */
function managerMailList_() {
  const lst = listsSheet_();
  const names = lst.getRange(3, 5, 200, 1).getValues();
  const mails = lst.getRange(3, 13, 200, 1).getValues();
  const out = [];
  for (let i = 0; i < names.length; i++) {
    if (!names[i][0]) break;
    out.push({ manager: str_(names[i][0]), email: str_(mails[i][0]) });
  }
  return out;
}

function sheetLinkButtonHtml() {
  return '<div style="margin-top:16px;"><a href="' + SHEET_URL + '" style="display:inline-block;padding:9px 16px;background:' + HEAD_BG +
    ';color:#fff;text-decoration:none;border-radius:6px;font-size:13px;font-weight:bold;">Открыть таблицу</a></div>';
}

function mailShell_(manager, inner) {
  return '<div style="font-family:Arial,sans-serif;color:#1F2328;max-width:560px;">' +
    '<h2 style="color:' + HEAD_BG + ';margin:0 0 12px;">Привет, ' + esc_(manager) + '!</h2>' + inner + sheetLinkButtonHtml() + '</div>';
}

function taskCardHtml(item, bg, tagHtml) {
  const subj = item.link
    ? '<a href="' + esc_(item.link) + '" style="color:#1155CC;text-decoration:none;font-weight:bold;">' + esc_(item.subject) + '</a>'
    : '<span style="font-weight:bold;">' + esc_(item.subject) + '</span>';
  return '<div style="background:' + bg + ';border-radius:8px;padding:10px 14px;margin-bottom:10px;">' + (tagHtml || '') + '<div>' + subj + '</div></div>';
}

function deadlineLabel(row) {
  if (row.exact instanceof Date) return 'Срок сдачи: ' + fmtDate_(row.exact, 'dd.MM.yyyy');
  if (row.approx) return 'Срочность: ' + row.approx;
  return 'Срок не указан';
}

function sendWeeklyDigests() {
  const sh = getTasksSheet();
  const data = readRows_(sh, TASK_COLS);
  const links = subjectLinks_(sh);
  managerMailList_().forEach(m => {
    if (!m.email) return; // рассылка строго по колонке M «Списков»
    const rows = [];
    data.forEach((r, i) => {
      if (str_(r[COL.MANAGER - 1]) !== m.manager || !r[COL.SUBJECT - 1]) return;
      if (CLOSED_STATUSES.indexOf(str_(r[COL.STATUS - 1])) !== -1) return;
      rows.push({ subject: stripLinkMarkers_(r[COL.SUBJECT - 1]), link: links[i].link, exact: r[COL.DUE - 1], approx: str_(r[COL.DEADLINE - 1]) });
    });
    if (!rows.length) return;
    const cards = rows.map(r => taskCardHtml(r, '#F1F3F5', '<div style="font-size:12px;color:#990000;margin-bottom:4px;">' + esc_(deadlineLabel(r)) + '</div>')).join('');
    MailApp.sendEmail({
      to: m.email,
      subject: 'Открытые задачи — еженедельная сводка (' + rows.length + ')',
      body: 'Привет, ' + m.manager + '!\n\nТвои открытые задачи (' + rows.length + '):\n\n' + rows.map(r => '- ' + r.subject + ' — ' + deadlineLabel(r)).join('\n'),
      htmlBody: mailShell_(m.manager, '<p>Твои открытые задачи (' + rows.length + '):</p>' + cards)
    });
  });
}

function createWeeklyDigestTrigger() {
  replaceTrigger_('sendWeeklyDigests', b => b.onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9));
  SpreadsheetApp.getUi().alert('Готово: сводка будет приходить каждый понедельник в 9:00.');
}

function sendDueSoonAlerts() {
  const sh = getTasksSheet();
  const data = readRows_(sh, TASK_COLS);
  const links = subjectLinks_(sh);
  const today = today_();
  const tomorrow = new Date(today); tomorrow.setDate(tomorrow.getDate() + 1);
  const sameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();

  managerMailList_().forEach(m => {
    if (!m.email) return;
    const overdue = [], dueToday = [], dueTomorrow = [];
    data.forEach((r, i) => {
      if (str_(r[COL.MANAGER - 1]) !== m.manager || !r[COL.SUBJECT - 1]) return;
      if (CLOSED_STATUSES.indexOf(str_(r[COL.STATUS - 1])) !== -1) return;
      const due = r[COL.DUE - 1];
      if (!(due instanceof Date)) return;
      const item = { subject: stripLinkMarkers_(r[COL.SUBJECT - 1]), link: links[i].link, exact: due };
      if (sameDay(due, today)) dueToday.push(item);
      else if (due < today) overdue.push(item);
      else if (sameDay(due, tomorrow)) dueTomorrow.push(item);
    });
    const total = overdue.length + dueToday.length + dueTomorrow.length;
    if (!total) return;

    let html = '', plain = 'Привет, ' + m.manager + '!\n\n';
    if (overdue.length) {
      html += '<p>Тут накопились просроченные задачки. Обрати внимание и не забудь проставить актуальный статус!</p>';
      overdue.forEach(it => {
        const days = Math.round((today - it.exact) / 86400000);
        html += taskCardHtml(it, '#FDEDEA', '<div style="font-size:12px;font-weight:bold;color:#990000;">🔴 Просрочено на ' + days + ' дн.</div>');
        plain += '- [ПРОСРОЧЕНО] ' + it.subject + '\n';
      });
    }
    if (dueToday.length) {
      html += '<p>Срок сегодня:</p>';
      dueToday.forEach(it => { html += taskCardHtml(it, '#FCE5CD', '<div style="font-size:12px;font-weight:bold;color:#990000;">🔥 Сегодня!</div>'); plain += '- [СЕГОДНЯ] ' + it.subject + '\n'; });
    }
    if (dueTomorrow.length) {
      html += '<p>Срок завтра:</p>';
      dueTomorrow.forEach(it => { html += taskCardHtml(it, '#FFF2CC', '<div style="font-size:12px;font-weight:bold;color:#7F6000;">⏰ Завтра</div>'); plain += '- [ЗАВТРА] ' + it.subject + '\n'; });
    }
    MailApp.sendEmail({
      to: m.email,
      subject: (overdue.length ? '🔴 Есть просроченные задачи — ' : '🔥 Горящие сроки — ') + total,
      body: plain, htmlBody: mailShell_(m.manager, html)
    });
  });
}

function createDailyDueSoonTrigger() {
  replaceTrigger_('sendDueSoonAlerts', b => b.everyDays(1).atHour(9));
  SpreadsheetApp.getUi().alert('Готово: напоминания о сроках будут приходить каждый день в 9:00.');
}

function sendMonthEndReminders() {
  const sh = getTasksSheet();
  const data = readRows_(sh, TASK_COLS);
  const links = subjectLinks_(sh);
  const checks = [
    [COL.TICKET, 'Тикет'], [COL.DATE, 'Дата'], [COL.PRODUCT, 'Продукт'], [COL.CUSTOMER, 'Ник заказчика'],
    [COL.LANGS, 'Языки'], [COL.DEADLINE, 'Дедлайн'], [COL.DUE, 'Срок сдачи'], [COL.STATUS, 'Статус'],
    [COL.ESTIMATE, 'Смета'], [COL.TOTAL, 'Итого с НДС'], [COL.CONTRACTOR, 'Подрядчик'], [COL.MANAGER, 'Менеджер'],
    [COL.DELIVERY, 'Статус отдачи'], [COL.SP, 'SP']
  ];
  const now = new Date();
  let sent = 0, skippedNoEmail = 0, withGaps = 0;

  managerMailList_().forEach(m => {
    const gaps = [];
    data.forEach((r, i) => {
      if (str_(r[COL.MANAGER - 1]) !== m.manager || !r[COL.SUBJECT - 1]) return;
      const d = r[COL.DATE - 1];
      if (!(d instanceof Date) || d.getFullYear() !== now.getFullYear() || d.getMonth() !== now.getMonth()) return;
      const missing = checks.filter(c => r[c[0] - 1] === '' || r[c[0] - 1] === null).map(c => c[1]);
      if (!links[i].link) missing.push('Ссылка на Band');
      if (missing.length) gaps.push({ subject: stripLinkMarkers_(r[COL.SUBJECT - 1]), link: links[i].link, missing: missing });
    });
    if (!gaps.length) return;
    withGaps++;
    if (!m.email) { skippedNoEmail++; return; }
    const cards = gaps.map(g => taskCardHtml(g, '#FFF8E6', '')
      .replace('</div></div>', '</div><div style="font-size:12px;color:#7F6000;margin-top:4px;">Не заполнено: ' + esc_(g.missing.join(', ')) + '</div></div>')).join('');
    MailApp.sendEmail({
      to: m.email,
      subject: '📋 Незаполненные поля перед концом месяца — ' + gaps.length,
      body: 'Привет, ' + m.manager + '!\n\nНе хватает данных для отчёта в этих задачах:\n\n' +
        gaps.map(g => '- ' + g.subject + ' (не заполнено: ' + g.missing.join(', ') + ')').join('\n'),
      htmlBody: mailShell_(m.manager, '<p>Месяц скоро закроется, а у этих задач не хватает данных — пожалуйста, заполни:</p>' + cards +
        '<p style="font-size:11px;color:#999;">Некоторые поля могут остаться пустыми, если по ним просто нет данных.</p>')
    });
    sent++;
  });

  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {} // по расписанию окна нет
  if (!ui) return;
  if (!withGaps) ui.alert('У всех задач за этот месяц всё заполнено — писать некому.');
  else ui.alert('Отправлено писем: ' + sent + ' из ' + withGaps + '.' +
    (skippedNoEmail ? '\n\n⚠️ Пропущено ' + skippedNoEmail + ' — не заполнена колонка M «Email для рассылки» в «Списках».' : ''));
}

function createMonthEndReminderTrigger() {
  replaceTrigger_('sendMonthEndReminders', b => b.onMonthDay(26).atHour(10));
  SpreadsheetApp.getUi().alert('Готово: напоминание будет приходить 26 числа в 10:00.');
}

function replaceTrigger_(handler, configure) {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t); });
  configure(ScriptApp.newTrigger(handler).timeBased()).create();
}

// ==================== ПЕРЕВОДЧИКИ ====================

// Короткие имена, которые можно писать в таблице, — связаны с полными
const TRANSLATOR_NICKNAMES = {
  'Валерия Гасанова': ['Лера'], 'Виктория Гусева': ['Вика'], 'Влада Тимощенко': ['Влада'],
  'Дмитрий Меделяновский': ['Дима'], 'Максим Селищев': ['Максим'], 'Мария Трофимова': ['Маша Т', 'Маша'],
  'София Горская': ['Соня'], 'Татьяна Козлова': ['Таня']
};

function resolveTranslatorName(text) {
  const t = str_(text);
  if (!t || TRANSLATOR_NICKNAMES[t]) return t;
  for (const full in TRANSLATOR_NICKNAMES) if (TRANSLATOR_NICKNAMES[full].indexOf(t) !== -1) return full;
  return t;
}

function translatorMatches(cellText, fullName) {
  const text = String(cellText || '');
  if (!text) return false;
  if (text.indexOf(fullName) !== -1) return true;
  return (TRANSLATOR_NICKNAMES[fullName] || []).some(n => text.indexOf(n) !== -1);
}

function getTranslatorNames() {
  return getListsData().translators;
}

function showAddTranslatorTaskDialog() {
  showDialog_('AddTranslatorTaskDialog', 'Новая задача переводчика', 520, 680);
}

function submitNewTranslatorTask(task) {
  return withScriptLock_(() => {
    const sh = getTranslatorsSheet();
    const row = insertTopRow_(sh, TR_COLS, () => applyTranslatorValidations_(sh, 2, 1));
    const values = [
      parseIsoDate_(task.date) || new Date(), str_(task.side), str_(task.razdel), str_(task.task), str_(task.customer),
      str_(task.translator), str_(task.editor), str_(task.readiness) || 'Не начато', str_(task.comment)
    ];
    sh.getRange(row, 1, 1, TR_COLS).setValues([values]);
    sh.getRange(row, TCOL.DATE).setNumberFormat('dd.mm.yy');
    if (values[TCOL.TASK - 1] && str_(task.link)) {
      sh.getRange(row, TCOL.TASK).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(values[TCOL.TASK - 1]).setLinkUrl(str_(task.link)).build());
    }
    logChange('Создание задачи (переводчик)', '', 'Задача', '', values[TCOL.TASK - 1]);
    return true;
  });
}

function showSearchEditTranslatorSidebar() {
  showSidebar_('SearchEditTranslatorSidebar', 'Поиск и правка (переводчики)');
}

function searchTranslatorTasks(query) {
  const sh = getTranslatorsSheet();
  const rows = readRows_(sh, TR_COLS);
  const q = str_(query).toLowerCase();
  const me = q ? '' : currentTranslator_();
  const out = [];
  rows.forEach((r, i) => {
    if (!r[TCOL.TASK - 1] && !r[TCOL.RAZDEL - 1]) return;
    if (q) {
      const hay = [r[TCOL.TASK - 1], r[TCOL.CUSTOMER - 1], r[TCOL.COMMENT - 1], r[TCOL.RAZDEL - 1], r[TCOL.SIDE - 1],
        r[TCOL.TRANSLATOR - 1], r[TCOL.EDITOR - 1]].map(v => String(v || '').toLowerCase()).join(' ');
      if (hay.indexOf(q) === -1) return;
    } else if (me && !translatorMatches(r[TCOL.TRANSLATOR - 1], me) && !translatorMatches(r[TCOL.EDITOR - 1], me)) {
      return;
    }
    out.push({
      row: i + 2, title: str_(r[TCOL.TASK - 1]) || '(задача не заполнена) ' + str_(r[TCOL.RAZDEL - 1]),
      side: str_(r[TCOL.SIDE - 1]), date: fmtDate_(r[TCOL.DATE - 1], 'dd.MM'), readiness: str_(r[TCOL.READY - 1])
    });
  });
  return out.slice(0, 30);
}

function getTranslatorTaskForEdit(row) {
  const sh = getTranslatorsSheet();
  const r = sh.getRange(row, 1, 1, TR_COLS).getValues()[0];
  const rt = sh.getRange(row, TCOL.TASK).getRichTextValue();
  return {
    row: row, date: fmtDate_(r[TCOL.DATE - 1], 'yyyy-MM-dd'), side: str_(r[TCOL.SIDE - 1]), razdel: str_(r[TCOL.RAZDEL - 1]),
    task: str_(r[TCOL.TASK - 1]), customer: str_(r[TCOL.CUSTOMER - 1]), link: linksFromRich_(rt).link,
    translator: str_(r[TCOL.TRANSLATOR - 1]), editor: str_(r[TCOL.EDITOR - 1]),
    translatorNames: splitPeople_(r[TCOL.TRANSLATOR - 1]), editorNames: splitPeople_(r[TCOL.EDITOR - 1]),
    readiness: str_(r[TCOL.READY - 1]), comment: str_(r[TCOL.COMMENT - 1]),
    origKey: translatorRowKey_(r)
  };
}

/** «Таня, Максим Селищев» → полные имена по списку коротких имён. */
function splitPeople_(text) {
  return String(text || '').split(/[,;]/).map(resolveTranslatorName).filter(Boolean);
}

/** У задач переводчиков нет номера — узнаём строку по дате, стороне, разделу и задаче. */
function translatorRowKey_(r) {
  return [fmtDate_(r[TCOL.DATE - 1], 'yyyy-MM-dd'), str_(r[TCOL.SIDE - 1]), str_(r[TCOL.RAZDEL - 1]), str_(r[TCOL.TASK - 1])].join('|');
}

function locateTranslatorRow_(sh, row, key) {
  const n = dataRowCount_(sh);
  if (row >= 2 && row <= n + 1 && translatorRowKey_(sh.getRange(row, 1, 1, TR_COLS).getValues()[0]) === key) return row;
  const rows = readRows_(sh, TR_COLS);
  for (let i = 0; i < rows.length; i++) if (translatorRowKey_(rows[i]) === key) return i + 2;
  throw new Error('Задача не найдена — возможно, её удалили или изменили. Найдите её заново.');
}

function saveTranslatorTaskEdits(task) {
  return withScriptLock_(() => {
    const sh = getTranslatorsSheet();
    const row = locateTranslatorRow_(sh, Number(task.row), task.origKey);
    const values = [
      parseIsoDate_(task.date), str_(task.side), str_(task.razdel), str_(task.task), str_(task.customer),
      str_(task.translator), str_(task.editor), str_(task.readiness), str_(task.comment)
    ];
    sh.getRange(row, 1, 1, TR_COLS).setValues([values]);
    if (values[TCOL.TASK - 1] && str_(task.link)) {
      sh.getRange(row, TCOL.TASK).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(values[TCOL.TASK - 1]).setLinkUrl(str_(task.link)).build());
    }
    logChange('Правка (переводчик)', '', 'Задача', '', values[TCOL.TASK - 1]);
    return { ok: true, row: row, origKey: translatorRowKey_(values) };
  });
}

function deleteTranslatorTask(row, key) {
  return withScriptLock_(() => {
    const sh = getTranslatorsSheet();
    const r = locateTranslatorRow_(sh, Number(row), key);
    logChange('Удаление задачи (переводчик)', '', 'Задача', sh.getRange(r, TCOL.TASK).getValue(), '');
    sh.deleteRow(r);
    return true;
  });
}

// --- фильтры переводчиков ---

function setTranslatorFilter_(col, hiddenValues, message) {
  const sh = getTranslatorsSheet();
  const filter = sh.getFilter() || sh.getDataRange().createFilter();
  filter.setColumnFilterCriteria(col, SpreadsheetApp.newFilterCriteria().setHiddenValues(hiddenValues).build());
  SpreadsheetApp.getActive().toast(message + '. Вернуть: «👁 Показать все».', 'Фильтр', 6);
}

function filterPerson_(col, title) {
  const names = getTranslatorNames();
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(title, 'Впишите имя как в списке:\n\n' + names.join(', '), ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const name = resp.getResponseText().trim();
  if (!name) return;
  const values = readRows_(getTranslatorsSheet(), col).map(r => r[col - 1]).filter(Boolean);
  const hidden = values.filter((v, i) => values.indexOf(v) === i && !translatorMatches(v, name));
  setTranslatorFilter_(col, hidden, 'Показаны задачи: ' + name);
}

function filterTranslator() { filterPerson_(TCOL.TRANSLATOR, 'Показать одного переводчика'); }
function filterEditor() { filterPerson_(TCOL.EDITOR, 'Показать одного редактора'); }

function filterTranslatorByReadiness(status) {
  setTranslatorFilter_(TCOL.READY, READINESS.filter(s => s !== status), 'Показаны задачи: ' + status);
}
function showInProgressTranslatorTasks() { filterTranslatorByReadiness('В работе'); }
function showDoneTranslatorTasks() { filterTranslatorByReadiness('Готово'); }

function showAllTranslatorTasks() {
  resetFilter_(getTranslatorsSheet());
  SpreadsheetApp.getActive().toast('Показаны все задачи переводчиков');
}

function showMyTranslatorTasks() {
  const me = currentTranslator_();
  if (!me) {
    SpreadsheetApp.getUi().alert(
      'Не нашла вашу почту в списке переводчиков.\n' +
      'Попросите вписать её на листе «Списки», колонка K — тогда «Мои задачи» заработают.\n' +
      'Пока можно выбрать себя через «Показать одного переводчика».');
    return;
  }
  const filter = resetFilter_(getTranslatorsSheet());
  filter.setColumnFilterCriteria(TCOL.TRANSLATOR, SpreadsheetApp.newFilterCriteria().whenTextContains(me).build());
  SpreadsheetApp.getActive().toast('Показаны задачи: ' + me + '. Вернуть: «👁 Показать все».', 'Фильтр', 6);
}

// --- отчёт по переводчику ---

function showTranslatorReportSidebar() {
  showDialog_('TranslatorReportSidebar', 'Отчёт по переводчику', 560, 720);
}

function getTranslatorReport(translator, monthStr) {
  const sh = getTranslatorsSheet();
  const data = readRows_(sh, TR_COLS);
  const links = subjectLinks_(sh, TCOL.TASK);
  const p = parseMonth_(monthStr);
  const bySide = {};
  let total = 0, done = 0, inProgress = 0;
  data.forEach((r, i) => {
    const task = str_(r[TCOL.TASK - 1]), readiness = str_(r[TCOL.READY - 1]) || 'Не начато';
    if (!task || !translatorMatches(r[TCOL.TRANSLATOR - 1], translator)) return;
    if (p.year && !inPeriod_(r[TCOL.DATE - 1], p.year, p.month)) return;
    total++;
    if (readiness === 'Готово') done++;
    if (readiness === 'В работе') inProgress++;
    const key = str_(r[TCOL.SIDE - 1]) || '(без стороны)';
    (bySide[key] = bySide[key] || []).push({ task: task, link: links[i].link, readiness: readiness });
  });
  return { total: total, done: done, inProgress: inProgress, bySide: bySide };
}

function exportTranslatorReportToExcel(translator, monthStr) {
  const report = getTranslatorReport(translator, monthStr);
  const ss = SpreadsheetApp.create('Отчёт по переводчику — ' + translator);
  const sheet = ss.getSheets()[0].setName('Отчёт');
  const rows = [['Сторона', 'Задача', 'Готовность']];
  const linkRows = [];
  Object.keys(report.bySide).forEach(side => report.bySide[side].forEach(t => {
    rows.push([side, t.task, t.readiness]);
    if (t.link) linkRows.push([rows.length, t.task, t.link]);
  }));
  rows.push(['', '', '']);
  rows.push(['Всего: ' + report.total + ' | Готово: ' + report.done + ' | В работе: ' + report.inProgress, '', '']);
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  linkRows.forEach(([r, text, link]) => sheet.getRange(r, 2).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(link).build()));
  styleReportSheet_(sheet, 3, 400);
  return exportUrl_(ss);
}

function getTranslatorTrackerReportText(translator) {
  const sh = getTranslatorsSheet();
  const data = readRows_(sh, TR_COLS);
  const links = subjectLinks_(sh, TCOL.TASK);
  const rows = [];
  data.forEach((r, i) => {
    if (!r[TCOL.TASK - 1] || !translatorMatches(r[TCOL.TRANSLATOR - 1], translator)) return;
    rows.push({ task: str_(r[TCOL.TASK - 1]), link: links[i].link });
  });
  if (!rows.length) return { text: '', count: 0 };
  let text = 'Задачи в Band:\n\n';
  rows.forEach((r, i) => { text += (i + 1) + '. ' + (r.link ? '[' + r.task + '](' + r.link + ')' : r.task) + '\n'; });
  return { text: text.trim(), count: rows.length };
}

// ==================== ОФОРМЛЕНИЕ ====================

/**
 * «⚙️ Настройки → 🎨 Оформить таблицу». Можно запускать сколько угодно раз.
 * Колонки, их порядок и данные не меняются — только вид, выпадающие списки и инструкции.
 */
function setupDesign() {
  const ss = SpreadsheetApp.getActive();
  ensureListsExtras_();
  designTasksSheet_(getTasksSheet());
  designTranslatorsSheet_(getTranslatorsSheet());
  writeManagerGuide_();
  writeTranslatorGuide_();
  ss.toast('Готово: листы оформлены, инструкции обновлены.', 'Оформление', 6);
}

/** «Списки», колонка O — префиксы номеров задач для каждого подрядчика. */
function ensureListsExtras_() {
  const lst = listsSheet_();
  if (lst.getMaxColumns() < 15) lst.insertColumnsAfter(lst.getMaxColumns(), 15 - lst.getMaxColumns());
  if (!str_(lst.getRange('O2').getValue())) lst.getRange('O2').setValue('Префикс подрядчика').setFontWeight('bold');
  const contractors = getListsData().contractors;
  const cur = lst.getRange(3, 15, Math.max(contractors.length, 1), 1).getValues();
  const next = contractors.map((c, i) => [str_(cur[i] && cur[i][0]) || CONTRACTOR_PREFIX[c] || '']);
  if (next.length) lst.getRange(3, 15, next.length, 1).setValues(next);
}

/** Правило «значение равно X» с цветом. */
function textRule_(range, value, pair) {
  return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(value)
    .setBackground(pair[0]).setFontColor(pair[1]).setBold(true).setRanges([range]).build();
}

/** Правила цветов для листа менеджеров. Диапазоны с 1-й строки — чтобы новые строки сверху попадали внутрь. */
function taskColorRules_(sh, lists) {
  const col = c => sh.getRange(1, c, sh.getMaxRows(), 1);
  const rules = [];
  lists.products.forEach((v, i) => rules.push(textRule_(col(COL.PRODUCT), v, BADGE_PALETTE[i % BADGE_PALETTE.length])));
  Object.keys(DEADLINE_COLORS).forEach(v => rules.push(textRule_(col(COL.DEADLINE), v, DEADLINE_COLORS[v])));
  Object.keys(STATUS_COLORS).forEach(v => rules.push(textRule_(col(COL.STATUS), v, STATUS_COLORS[v])));
  lists.contractors.forEach((v, i) => rules.push(textRule_(col(COL.CONTRACTOR), v, CONTRACTOR_PALETTE[i % CONTRACTOR_PALETTE.length])));
  lists.managers.forEach((v, i) => rules.push(textRule_(col(COL.MANAGER), v, MANAGER_PALETTE[i % MANAGER_PALETTE.length])));
  Object.keys(DELIVERY_STATUS_COLORS).forEach(v => rules.push(textRule_(col(COL.DELIVERY), v, DELIVERY_STATUS_COLORS[v])));

  // Просрочка: срок прошёл, а задача не «Отдано»/«Отменено». Срок — красным, строка — розовым.
  const overdue = '=AND(ROW()>1,$I1<>"",$I1<TODAY(),$J1<>"Отдано",$J1<>"Отменено")';
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(overdue)
    .setBackground(OVERDUE_BG).setFontColor(OVERDUE_INK).setBold(true).setRanges([col(COL.DUE)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(overdue)
    .setBackground(OVERDUE_BG).setRanges([sh.getRange(1, 1, sh.getMaxRows(), TASK_COLS)]).build());
  return rules;
}

function translatorColorRules_(sh, lists) {
  const col = c => sh.getRange(1, c, sh.getMaxRows(), 1);
  const rules = [];
  Object.keys(getContentCategories()).forEach((v, i) => rules.push(textRule_(col(TCOL.SIDE), v, BADGE_PALETTE[i % BADGE_PALETTE.length])));
  // Один человек — один цвет и в «Переводчике», и в «Редакторе». Узнаём и по короткому имени.
  lists.translators.forEach((name, i) => {
    const pair = MANAGER_PALETTE[i % MANAGER_PALETTE.length];
    [name].concat(TRANSLATOR_NICKNAMES[name] || []).forEach(n => {
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith(n)
        .setBackground(pair[0]).setFontColor(pair[1]).setBold(true).setRanges([col(TCOL.TRANSLATOR), col(TCOL.EDITOR)]).build());
    });
  });
  Object.keys(READINESS_COLORS).forEach(v => rules.push(textRule_(col(TCOL.READY), v, READINESS_COLORS[v])));
  return rules;
}

/** Перестраивает только цвета (после правки «Списков»). */
function refreshColorRules_() {
  const lists = getListsData();
  const tasks = getTasksSheet();
  tasks.setConditionalFormatRules(taskColorRules_(tasks, lists));
  const tr = getTranslatorsSheet();
  tr.setConditionalFormatRules(translatorColorRules_(tr, lists));
}

function styleHeader_(sh, ncols) {
  sh.getRange(1, 1, 1, ncols).setBackground(HEAD_BG).setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontFamily(FONT).setFontSize(10).setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 36);
}

/**
 * Группа колонок: над ними появляется «–», которым их можно свернуть, и «+», чтобы развернуть.
 * Сами не сворачиваем — иначе колонки «пропадают» и их не найти.
 */
function groupColumns_(sh, from, count) {
  if (sh.getColumnGroupDepth(from) === 0) sh.getRange(1, from, 1, count).shiftColumnGroupDepth(1);
  try { sh.getColumnGroup(from, 1).expand(); } catch (e) {}
}

function designTasksSheet_(sh) {
  const lists = getListsData();
  const n = Math.max(dataRowCount_(sh), 1);
  const maxRows = sh.getMaxRows();

  styleHeader_(sh, TASK_COLS);
  sh.getRange(1, COL.DUE).setValue('Срок сдачи'); // было «Точный дедлайн (дата сдачи заказчику)» — вводило в заблуждение
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);
  // Ширина — чтобы продукт, языки и дедлайн помещались, а шапка не рвала слова («Примерны / й дедлайн»)
  [115, 105, 380, 92, 215, 175, 330, 150, 100, 115, 205, 110, 140, 170, 170, 60, 260]
    .forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // Данные: один шрифт, без старой ручной раскраски — цвета ставят правила ниже
  const body = sh.getRange(2, 1, maxRows - 1, TASK_COLS);
  // По центру строки и с переносом: в коротких строках текст не «липнет» к верху, длинный не обрезается
  body.setBackground(null).setFontFamily(FONT).setFontSize(10).setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  [COL.ID, COL.TICKET, COL.DATE, COL.PRODUCT, COL.CUSTOMER, COL.DEADLINE, COL.DUE, COL.STATUS, COL.ESTIMATE,
   COL.TOTAL, COL.CONTRACTOR, COL.MANAGER, COL.DELIVERY, COL.SP, COL.COMMENT].forEach(c => {
    sh.getRange(2, c, maxRows - 1, 1).setFontColor(INK).setFontWeight('normal');
  });
  sh.getRange(2, COL.ID, maxRows - 1, 1).setFontWeight('bold');
  sh.getRange(2, COL.SUBJECT, maxRows - 1, 1).setFontLine('none').setFontWeight('bold'); // тема письма — жирным
  sh.getRange(2, COL.ESTIMATE, maxRows - 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sh.getRange(2, COL.DATE, maxRows - 1, 1).setNumberFormat('dd.mm.yyyy');
  sh.getRange(2, COL.DUE, maxRows - 1, 1).setNumberFormat('dd.mm.yyyy');
  sh.getRange(2, COL.TOTAL, maxRows - 1, 1).setNumberFormat('#,##0.00'); // просто число, без «₽»
  sh.getRange(2, COL.ESTIMATE, maxRows - 1, 1).setHorizontalAlignment('center');
  estimateLinksToLabels_(sh);
  sh.getRange(2, COL.SP, maxRows - 1, 1).setNumberFormat('0');

  // Языки: если «чипы» не включены — каждый язык своим цветом текста (одним запросом на весь лист)
  if (!sh.getRange(2, COL.LANGS).getDataValidation()) {
    const langs = sh.getRange(2, COL.LANGS, n, 1).getValues();
    sh.getRange(2, COL.LANGS, n, 1).setRichTextValues(langs.map(r => [langsRich_(r[0])]));
  }

  // Выпадающие списки — только в строках с задачами (в пустых нет лишних стрелочек).
  // Колонку «Языки» не трогаем: там можно вручную включить «чипы».
  applyTaskValidations_(sh, 2, n);
  if (maxRows > n + 1) {
    TASK_LIST_SOURCES.map(s => s[0]).concat([COL.DATE, COL.DUE])
      .forEach(c => sh.getRange(n + 2, c, maxRows - n - 1, 1).clearDataValidations());
  }

  sh.setConditionalFormatRules(taskColorRules_(sh, lists));
  groupColumns_(sh, COL.ESTIMATE, 2);   // Смета, Итого
  groupColumns_(sh, COL.DELIVERY, 3);   // Статус отдачи, SP, Комментарий
  sh.autoResizeRows(2, n);
  if (!sh.getFilter()) sh.getDataRange().createFilter();
}

function designTranslatorsSheet_(sh) {
  const lists = getListsData();
  const n = Math.max(dataRowCount_(sh), 1);
  const maxRows = sh.getMaxRows();

  styleHeader_(sh, TR_COLS);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);
  [90, 150, 170, 320, 180, 190, 190, 120, 240].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const body = sh.getRange(2, 1, maxRows - 1, TR_COLS);
  body.setBackground(null).setFontFamily(FONT).setFontSize(10).setVerticalAlignment('middle').setWrap(true)
    .setHorizontalAlignment('left');
  [TCOL.DATE, TCOL.SIDE, TCOL.RAZDEL, TCOL.CUSTOMER, TCOL.TRANSLATOR, TCOL.EDITOR, TCOL.READY, TCOL.COMMENT]
    .forEach(c => sh.getRange(2, c, maxRows - 1, 1).setFontColor(INK).setFontWeight('normal'));
  sh.getRange(2, TCOL.TASK, maxRows - 1, 1).setFontLine('none');
  sh.getRange(2, TCOL.DATE, maxRows - 1, 1).setNumberFormat('dd.mm.yy');

  applyTranslatorValidations_(sh, 2, n);
  if (maxRows > n + 1) {
    [TCOL.DATE, TCOL.SIDE, TCOL.TRANSLATOR, TCOL.EDITOR, TCOL.READY]
      .forEach(c => sh.getRange(n + 2, c, maxRows - n - 1, 1).clearDataValidations());
  }
  sh.setConditionalFormatRules(translatorColorRules_(sh, lists));
  sh.autoResizeRows(2, n);
  if (!sh.getFilter()) sh.getDataRange().createFilter();
}

// ==================== ИНСТРУКЦИИ ====================

/**
 * Инструкция — шпаргалка карточками в две колонки на сером фоне.
 * blocks: [{ icon, title, steps: [...], rows: [[термин, пояснение]], legend: [...], hint, wide }]
 * Колонки: A — поле, B — номер шага, C — текст, D — промежуток, E/F — вторая колонка карточек, G — поле.
 */
const GUIDE_BG = '#F3F4F6';
const GUIDE_TEXT_W = 380;       // ширина текстовой колонки карточки, px
const GUIDE_CHARS_PER_LINE = 58; // сколько символов влезает в строку — чтобы задать высоту строк

function writeGuide_(sheetName, title, lead, blocks) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName, 0);
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(0);
  if (sh.getMaxColumns() < 7) sh.insertColumnsAfter(sh.getMaxColumns(), 7 - sh.getMaxColumns());
  [20, 30, GUIDE_TEXT_W, 22, 30, GUIDE_TEXT_W, 20].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // Каждая карточка — список строк: { kind, num, text, rich }
  function cardLines(b) {
    const lines = [{ kind: 'head', text: b.icon + '  ' + b.title }];
    (b.steps || []).forEach((t, i) => lines.push({ kind: 'step', num: String(i + 1), text: t }));
    (b.rows || []).forEach(r => lines.push({ kind: 'kv', text: r[0] + ' — ' + r[1], bold: r[0].length }));
    if (b.legend) lines.push({ kind: 'legend', text: b.legend.join('   ·   '), legend: b.legend });
    if (b.hint) lines.push({ kind: 'hint', text: '💡 ' + b.hint });
    lines.push({ kind: 'pad', text: '' });
    return lines;
  }

  // Раскладка: широкие карточки — на всю ширину, остальные — парами
  const layout = [];   // { left, right, wide }
  let pending = null;
  blocks.forEach(b => {
    if (b.wide) { if (pending) { layout.push({ left: pending }); pending = null; } layout.push({ wide: b }); }
    else if (pending) { layout.push({ left: pending, right: b }); pending = null; }
    else pending = b;
  });
  if (pending) layout.push({ left: pending });

  const cells = [];   // [row, col, numCols, line]
  const heights = {};
  let row = 1;
  cells.push([row, 2, 5, { kind: 'title', text: title }]); heights[row] = 48; row++;
  cells.push([row, 2, 5, { kind: 'lead', text: lead }]); heights[row] = 34; row++;
  heights[row] = 14; row++;

  const cards = []; // [row, col, numRows, numCols]
  layout.forEach(item => {
    const parts = item.wide ? [[item.wide, 2, 5]] : [[item.left, 2, 2]].concat(item.right ? [[item.right, 5, 2]] : []);
    let maxLen = 0;
    parts.forEach(([b, col, width]) => {
      const lines = cardLines(b);
      maxLen = Math.max(maxLen, lines.length);
      lines.forEach((ln, i) => {
        const r = row + i;
        const textWidth = width === 5 ? GUIDE_TEXT_W * 2 + 80 : GUIDE_TEXT_W;
        const perLine = Math.floor(GUIDE_CHARS_PER_LINE * textWidth / GUIDE_TEXT_W);
        const nLines = Math.max(1, Math.ceil(ln.text.length / perLine));
        const h = ln.kind === 'head' ? 34 : ln.kind === 'pad' ? 8 : 8 + 18 * nLines;
        heights[r] = Math.max(heights[r] || 0, h);
        if (ln.kind === 'step') {
          cells.push([r, col, 1, { kind: 'num', text: ln.num }]);
          cells.push([r, col + 1, width - 1, ln]);
        } else {
          cells.push([r, col, width, ln]);
        }
      });
      cards.push([row, col, lines.length, width]);
    });
    row += maxLen;
    heights[row] = 16; row++; // промежуток между рядами карточек
  });
  const lastRow = row;

  // Фон страницы и белые карточки
  if (sh.getMaxRows() < lastRow + 2) sh.insertRowsAfter(sh.getMaxRows(), lastRow + 2 - sh.getMaxRows());
  sh.getRange(1, 1, lastRow + 2, 7).setBackground(GUIDE_BG).setFontFamily(FONT).setFontSize(10)
    .setFontColor(INK).setVerticalAlignment('middle').setWrap(true);
  cards.forEach(([r, c, nr, nc]) => {
    sh.getRange(r, c, nr, nc).setBackground('#FFFFFF')
      .setBorder(true, true, true, true, false, false, '#DADDE2', SpreadsheetApp.BorderStyle.SOLID);
  });

  cells.forEach(([r, c, nc, ln]) => {
    const range = sh.getRange(r, c, 1, nc);
    if (nc > 1) range.merge();
    const cell = sh.getRange(r, c);
    switch (ln.kind) {
      case 'title':
        cell.setValue(ln.text).setFontSize(18).setFontWeight('bold');
        break;
      case 'lead':
        cell.setValue(ln.text).setFontColor('#5F6B7A').setFontSize(10.5);
        break;
      case 'head':
        cell.setValue(ln.text).setFontSize(12).setFontWeight('bold');
        range.setBorder(null, null, true, null, null, null, '#E6E8EB', SpreadsheetApp.BorderStyle.SOLID);
        break;
      case 'num':
        cell.setValue(ln.text).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(HEAD_BG)
          .setHorizontalAlignment('center');
        break;
      case 'step':
        cell.setValue(ln.text);
        break;
      case 'kv':
        cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(ln.text)
          .setTextStyle(0, ln.bold, SpreadsheetApp.newTextStyle().setBold(true).build()).build());
        break;
      case 'legend': {
        const b = SpreadsheetApp.newRichTextValue().setText(ln.text);
        let pos = 0;
        ln.legend.forEach(name => {
          const pair = STATUS_COLORS[name] || READINESS_COLORS[name] || ['#FFFFFF', INK];
          b.setTextStyle(pos, pos + name.length, SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor(pair[1]).build());
          pos += name.length + 7;
        });
        cell.setRichTextValue(b.build());
        break;
      }
      case 'hint':
        range.setBackground('#FFF8E6');
        cell.setValue(ln.text).setFontColor('#6B5100');
        break;
    }
  });

  Object.keys(heights).forEach(r => sh.setRowHeight(Number(r), heights[r]));
  if (sh.getMaxRows() > lastRow + 2) sh.deleteRows(lastRow + 3, sh.getMaxRows() - lastRow - 2);
  if (sh.getMaxColumns() > 7) sh.deleteColumns(8, sh.getMaxColumns() - 7);
}

function writeManagerGuide_() {
  writeGuide_(MANAGER_GUIDE_SHEET, '📋 Шпаргалка для менеджеров',
    'Всё делается через меню «📋 Менеджеры» или через расширение Chrome «Задачи и сметы». Руками в листе можно менять статус и сроки.',
    [
      { icon: '➕', title: 'Добавить задачу', steps: [
        '«📋 Менеджеры → ➕ Добавить задачу» или вкладка «Задача» в расширении.',
        'Выберите подрядчика — номер появится сам (например, LIT-26-2232).',
        'Тема — только суть. Номер, подрядчик, коды языков и продукт добавятся сами.',
        'Тикет (можно просто цифры), ссылка на Band, продукт, языки, срок сдачи — по возможности сразу.',
        '«Добавить задачу» — строка появится наверху листа. Тему можно скопировать для письма.'],
        hint: 'Смету, итого и SP заполняют позже. Смету удобнее всего добавить через расширение, вкладка «Смета».' },
      { icon: '🔍', title: 'Найти и поправить', steps: [
        '«📋 Менеджеры → 🔍 Поиск и правка».',
        'Впишите номер, кусок темы, тикет или ник — выберите задачу. Пустой поиск показывает ваши задачи.',
        'Статус меняется одной кнопкой, остальное — в блоках ниже.',
        '«Сохранить». Все правки попадают в «📝 Журнал».'],
        hint: 'Удалить задачу — внизу окна, спросит подтверждение. Отменить удаление нельзя.' },
      { icon: '✏️', title: 'Можно править прямо в листе', rows: [
        ['Статус, сроки, SP…', 'Меняйте прямо в ячейке — цвета обновятся сами, правка попадёт в «📝 Журнал».'],
        ['Языки', 'Через запятую, как в других строках: «Грузинский, Казахский». Цвет появится сам.'],
        ['Новая задача', 'Можно вписать в пустую строку: как только есть тема и подрядчик, появится номер.']],
        hint: 'Через «➕ Добавить задачу» всё равно удобнее: тема письма соберётся сама, её можно сразу скопировать.' },
      { icon: '📅', title: 'Сроки и статусы', rows: [
        ['Дедлайн', 'Примерная срочность: ASAP, 1-2 дня, до недели…'],
        ['Срок сдачи', 'Дата, к которой нужно сдать заказчику.'],
        ['Статус', 'Стадия задачи.'],
        ['Статус отдачи', 'Отчитались ли перед заказчиком в срок.']],
        hint: 'Срок прошёл, а задача не «Отдано» и не «Отменено» — строка краснеет.' },
      { icon: '🎨', title: 'Цвета статусов', legend: ['Принято', 'В работе', 'Отдано', 'Отменено', 'Холд'] },
      { icon: '👀', title: 'Фильтры (меню «📋 Менеджеры»)', rows: [
        ['Мои задачи', 'Только ваши. Нужна ваша почта в «Списках», колонка F.'],
        ['Одного менеджера', 'Впишите имя — покажутся только его задачи.'],
        ['Скрыть закрытые', 'Прячет «Отдано» и «Отменено».'],
        ['Показать все', 'Сбрасывает любой фильтр.']],
        hint: 'Колонки «Смета, Итого» и «Статус отдачи, SP, Комментарий» можно свернуть кнопкой «–» над ними и развернуть «+». Языки выбираются в самой ячейке — можно несколько.' },
      { icon: '✉️', title: 'Письма, которые приходят сами', wide: true, rows: [
        ['Понедельник, 9:00', 'Сводка ваших открытых задач.'],
        ['Каждый день, 9:00', 'Просроченные и со сроком сегодня или завтра.'],
        ['26 числа, 10:00', 'Задачи месяца, где не хватает данных для отчёта.']],
        hint: 'Письма идут на почту из «Списков», колонка M. Нет почты — нет писем.' }
    ]);
}

function writeTranslatorGuide_() {
  writeGuide_(TRANSLATOR_GUIDE_SHEET, '👥 Шпаргалка для переводчиков EN',
    'Лист «✍️ Задачи (переводчики)» — отдельный: он не связан с задачами менеджеров, ничего не дублируется.',
    [
      { icon: '➕', title: 'Добавить задачу', steps: [
        '«👥 Переводчики EN → ➕ Добавить задачу переводчика».',
        'Дата, сторона (общий раздел продукта), раздел — список разделов подстроится под сторону.',
        'Задача — что нужно сделать, своими словами. Заказчик и ссылка на Band.',
        'Переводчик и редактор — отметьте, можно несколько.',
        '«Добавить задачу» — строка появится наверху листа.'],
        hint: 'В таблице можно писать короткие имена: Лера, Вика, Дима, Максим, Маша, Соня, Таня — скрипт их понимает.' },
      { icon: '🔍', title: 'Найти и поправить', steps: [
        '«👥 Переводчики EN → 🔍 Поиск и правка».',
        'Впишите кусок задачи, заказчика, раздела или имя — выберите задачу.',
        'Поменяйте что нужно и нажмите «Сохранить».'],
        hint: 'Удалить задачу — внизу окна, спросит подтверждение. Отменить удаление нельзя.' },
      { icon: '🎨', title: 'Готовность', legend: ['Не начато', 'В работе', 'На проверке', 'Готово'] },
      { icon: '👀', title: 'Фильтры (меню «👥 Переводчики EN»)', rows: [
        ['Мои задачи', 'Только ваши. Нужна ваша почта в «Списках», колонка K — попросите менеджера вписать.'],
        ['Одного переводчика', 'Впишите имя — покажутся только его задачи.'],
        ['В работе / Сданные', 'Быстрые фильтры по готовности.'],
        ['Показать все', 'Сбрасывает любой фильтр.']] }
    ]);
}


/*************************************************************
 * Перенос.gs — разовый перенос истории из старой таблицы
 * (лист «Localization Misc») в «📌 Задачи (менеджеры)».
 *
 * Меню: «⚙️ Настройки → 📥 Перенести историю из старой таблицы». В окне — файл .xlsx старой таблицы
 * (Файл → Скачать → Microsoft Excel) или ссылка на неё. Сначала показывает, сколько задач перенесёт
 * и сколько пропустит, и переносит только после подтверждения.
 * Задачи, номера которых уже есть в новой таблице, пропускаются — запускать повторно безопасно.
 * После переноса файл можно удалить.
 *************************************************************/

// OLD_SHEET_NAME и OLD_TABLE_PROP — в Код.gs (нужны и после удаления этого файла).

// Колонки старого листа (A..S)
const OLD = {
  ID: 1, SIDE: 2, DATE: 3, SUBJECT: 4, TASK: 5, BAND: 6, CUSTOMER: 7, LANGS: 8, DEADLINE: 9, DUE: 10,
  STATUS: 11, ESTIMATE: 12, TOTAL: 13, CONTRACTOR: 14, MANAGER: 15, DELIVERY: 16, SP: 17, COMMENT: 18, COMPLAINTS: 19
};
const OLD_COLS = 19;

// Объединённые ячейки в старой таблице: одна задача подрядчику на несколько строк (разные запросы из Band).
// Номер, тему, дату и т. п. раскопируем на каждую строку. Сумму и SP — нет, иначе деньги посчитаются дважды.
const OLD_NO_FILL = [OLD.TOTAL, OLD.SP, OLD.COMMENT, OLD.COMPLAINTS];

// «Сторона» из старой таблицы → «Продукт» (по памятке из инструкции)
const OLD_SIDE_TO_PRODUCT = {
  'Маркетинг': 'Магазинка Пуши и коммуникации: Маркетинг',
  'Маркетплейс': 'Магазинка',
  'Поддержка покупателей': 'Поддержка Покупатели',
  'Поддержка продавца': 'Поддержка Продавцы',
  'Поддержка продавцов': 'Поддержка Продавцы',
  'Справочный центр': 'Поддержка Продавцы',
  'Портал продавца': 'WBP',
  'Портал продавца, Справочный центр': 'WBP',
  '"Портал продавца, Справочный центр"': 'WBP',
  'Контент': 'WBP Контент',
  'Склады': 'Логобъекты WBWH',
  'WB Taxi': 'WB Такси',
  'WB Job': 'Логобъекты WB Job',
  'Логистика': 'Логобъекты Логистика',
  'Юридичка': 'Межнар',
  'Геоданные': 'Межнар',
  'ИБ': 'Инфобез',
  'ДРУГОЕ': 'Межнар'
};

/** Меню: окно переноса — файл Excel (.xlsx) или ссылка на старую таблицу. */
function migrateFromOldTable() {
  showDialog_('MigrateDialog', 'Перенос истории из старой таблицы', 520, 470);
}

function oldTableIdFromLink_(url) {
  const t = String(url || '').trim();
  const m = t.match(/\/d\/([a-zA-Z0-9_-]{20,})/) || t.match(/^([a-zA-Z0-9_-]{20,})$/);
  if (!m) throw new Error('Не похоже на ссылку на Google Таблицу');
  return m[1];
}

function planSummary_(plan) {
  return {
    count: plan.rows.length, skippedExisting: plan.skippedExisting, skippedEmpty: plan.skippedEmpty,
    from: plan.from, to: plan.to
  };
}

/** Окно переноса, способ «ссылка»: сначала сколько перенесётся, потом перенос. */
function migrationPreviewFromLink(url) {
  return planSummary_(planMigration_(SpreadsheetApp.openById(oldTableIdFromLink_(url))));
}

function migrationApplyFromLink(url) {
  const id = oldTableIdFromLink_(url);
  const plan = planMigration_(SpreadsheetApp.openById(id));
  // Запоминаем старую таблицу: новые номера будут сверяться с ней
  PropertiesService.getDocumentProperties().setProperty(OLD_TABLE_PROP, id);
  if (plan.rows.length) applyMigration_(plan);
  return { added: plan.rows.length };
}

/** Способ «файл»: окно само читает .xlsx и присылает лист «Localization Misc» (см. MigrateDialog). */
function migrationPreviewFromData(src) {
  return planSummary_(planFromSource_(sourceFromUpload_(src)));
}

function migrationApplyFromData(src) {
  const plan = planFromSource_(sourceFromUpload_(src));
  if (plan.rows.length) applyMigration_(plan);
  return { added: plan.rows.length };
}

/** Даты из окна приходят как { d: 'yyyy-MM-dd' }: google.script.run не передаёт Date. */
function sourceFromUpload_(src) {
  if (!src || !Array.isArray(src.values)) throw new Error('Файл не прочитался');
  const values = src.values.map(row => {
    const out = [];
    for (let c = 0; c < OLD_COLS; c++) {
      out.push(uploadValue_(row[c]));
    }
    return out;
  });
  return { values: values, links: (src.links || []).map(l => l || ''), merges: src.merges || [] };
}

function uploadValue_(v) {
  if (v && typeof v === 'object' && v.d) {
    const p = String(v.d).split('-').map(Number);
    return new Date(p[0], p[1] - 1, p[2]);
  }
  return v == null ? '' : v;
}

function numOrEmpty_(v) {
  if (v === '' || v === null) return '';
  const n = Number(String(v).replace(/\s/g, '').replace(',', '.'));
  return isFinite(n) ? n : v;
}

/** Читает старый лист из Google Таблицы: значения, ссылки на Band и объединённые ячейки. */
function readOldSheet_(oldSs) {
  const old = oldSs.getSheetByName(OLD_SHEET_NAME);
  if (!old) throw new Error('В старой таблице нет листа «' + OLD_SHEET_NAME + '»');
  const lastRow = old.getLastRow();
  const range = old.getRange(1, 1, lastRow, OLD_COLS);
  return {
    values: range.getValues(),
    links: old.getRange(1, OLD.BAND, lastRow, 1).getRichTextValues().map(r => linksFromRich_(r[0]).link || ''),
    merges: range.getMergedRanges().map(mr => [mr.getRow(), mr.getColumn(), mr.getNumRows(), mr.getNumColumns()])
  };
}

/** Читает старый лист и готовит строки для новой таблицы (ничего не записывает). */
function planMigration_(oldSs) {
  return planFromSource_(readOldSheet_(oldSs));
}

/**
 * source: { values: строки листа с шапкой (A..S), links: ссылка на Band для каждой строки,
 *           merges: [[строка, колонка, строк, колонок], …] — с 1, как в таблице }.
 */
function planFromSource_(src) {
  const values = src.values, links = src.links.slice();
  const lastRow = values.length;

  // Раскопировать объединённые ячейки на все их строки
  src.merges.forEach(m => {
    const r0 = m[0], c0 = m[1];
    if (r0 > lastRow) return;
    for (let c = c0; c < c0 + m[3] && c <= OLD_COLS; c++) {
      if (OLD_NO_FILL.indexOf(c) !== -1) continue;
      for (let r = r0 + 1; r < r0 + m[2] && r <= lastRow; r++) {
        values[r - 1][c - 1] = values[r0 - 1][c - 1];
        if (c === OLD.BAND) links[r - 1] = links[r0 - 1];
      }
    }
  });

  const idx = existingTaskIndex_(getTasksSheet());
  const existing = idx.ids, existingKeys = idx.keys, keyOf = taskKey_;

  const rows = [];
  let skippedExisting = 0, skippedEmpty = 0, from = null, to = null;
  for (let i = lastRow; i >= 2; i--) { // снизу вверх — свежие задачи первыми, как в новой таблице
    const v = values[i - 1];
    const get = c => str_(v[c - 1]);
    const id = get(OLD.ID), subject = get(OLD.SUBJECT), taskText = get(OLD.TASK);
    if (!id && !subject && !taskText && !get(OLD.BAND)) { skippedEmpty++; continue; }
    if (subject === 'Тема письма в аутлуке' || get(OLD.STATUS) === 'Статус') { skippedEmpty++; continue; }
    if (id && existing[id]) { skippedExisting++; continue; }

    // «Задача» в старой таблице — чаще номер тикета (LOCAL-493), иначе описание
    const ticketMatch = taskText.match(/^LOCAL[-\s]?(\d+)$/i);
    const ticket = ticketMatch ? 'LOCAL-' + ticketMatch[1] : '';
    // В одной ячейке бывает несколько ссылок на Band: первая — на тему, остальные — «(ссылка 2)», «(ссылка 3)»…
    const urls = get(OLD.BAND).match(/https?:\/\/[^\s,;]+/g) || [];
    const band = urls[0] || links[i - 1] || ''; // ссылка ячейки — если в тексте вместо адреса подпись
    const extra = urls.filter(u => u !== band);
    const comment = [
      get(OLD.COMMENT),
      taskText && !ticketMatch && taskText !== subject ? 'Задача: ' + taskText : '',
      get(OLD.COMPLAINTS) ? 'Жалобы на заказчика: ' + get(OLD.COMPLAINTS) : ''
    ].filter(Boolean).join('\n');

    if (!id && existingKeys[keyOf(subject || taskText || '(без темы)', v[OLD.DATE - 1], band)]) { skippedExisting++; continue; }
    const deadline = get(OLD.DEADLINE) === 'Дедлайн (если есть)' ? '' : get(OLD.DEADLINE);
    const side = get(OLD.SIDE);
    const date = v[OLD.DATE - 1];

    rows.push({
      link: band,
      link2: extra.join('\n'),
      values: [
        id, ticket, subject || taskText || '(без темы)', date, OLD_SIDE_TO_PRODUCT[side] || side,
        get(OLD.CUSTOMER), get(OLD.LANGS), deadline, v[OLD.DUE - 1] instanceof Date ? v[OLD.DUE - 1] : get(OLD.DUE),
        get(OLD.STATUS), get(OLD.ESTIMATE), numOrEmpty_(v[OLD.TOTAL - 1]), get(OLD.CONTRACTOR), get(OLD.MANAGER),
        get(OLD.DELIVERY), numOrEmpty_(v[OLD.SP - 1]), comment
      ]
    });
    if (date instanceof Date) {
      if (!from || date < from) from = date;
      if (!to || date > to) to = date;
    }
  }
  return {
    rows: rows, skippedExisting: skippedExisting, skippedEmpty: skippedEmpty,
    from: fmtDate_(from, 'dd.MM.yyyy') || '—', to: fmtDate_(to, 'dd.MM.yyyy') || '—'
  };
}

/** Что уже есть на листе: по номеру, а строки без номера — по теме, дате и ссылке. */
function taskKey_(subject, date, link) {
  return [stripLinkMarkers_(subject), fmtDate_(date, 'yyyy-MM-dd') || str_(date), link || ''].join('|');
}

function existingTaskIndex_(sh) {
  const ids = {}, keys = {};
  if (!sh || sh.getLastRow() < 2) return { ids: ids, keys: keys };
  const cur = readRows_(sh, COL.DATE);
  const curLinks = subjectLinks_(sh);
  cur.forEach((r, i) => {
    if (str_(r[COL.ID - 1])) ids[str_(r[COL.ID - 1])] = true;
    else keys[taskKey_(r[COL.SUBJECT - 1], r[COL.DATE - 1], curLinks[i] && curLinks[i].link)] = true;
  });
  return { ids: ids, keys: keys };
}

// ==================== ФАЙЛ В ФОРМАТЕ НОВОЙ ТАБЛИЦЫ ====================
// Например, «Перенос — задачи из старой таблицы.xlsx»: листы с теми же колонками, что «📌 Задачи (менеджеры)».
// Листы «Данные до …» / «Архив…» уходят на отдельный лист с тем же названием, остальные — в задачи.

/** src: { sheets: [{ name, archive, rows: [{ values: 17 значений, link, link2 }] }] } */
function tasksFilePlan_(src) {
  if (!src || !Array.isArray(src.sheets)) throw new Error('Файл не прочитался');
  const ss = SpreadsheetApp.getActive();
  const plan = { main: [], mainSkipped: 0, archives: [] };
  const mainIdx = existingTaskIndex_(getTasksSheet());
  src.sheets.forEach(sheet => {
    const idx = sheet.archive ? existingTaskIndex_(ss.getSheetByName(sheet.name)) : mainIdx;
    const target = sheet.archive ? { name: str_(sheet.name), rows: [], skipped: 0 } : null;
    (sheet.rows || []).forEach(r => {
      const values = [];
      for (let c = 0; c < TASK_COLS; c++) values.push(uploadValue_((r.values || [])[c]));
      if (!values.some(v => v !== '')) return;
      values[COL.TOTAL - 1] = numOrEmpty_(values[COL.TOTAL - 1]);
      values[COL.SP - 1] = numOrEmpty_(values[COL.SP - 1]);
      values[COL.SUBJECT - 1] = stripLinkMarkers_(values[COL.SUBJECT - 1]);
      const id = str_(values[COL.ID - 1]);
      const row = { values: values, link: str_(r.link), link2: extraLinks_(r.link2).join('\n') };
      const known = id ? idx.ids[id] : idx.keys[taskKey_(values[COL.SUBJECT - 1], values[COL.DATE - 1], row.link)];
      if (known) { if (target) target.skipped++; else plan.mainSkipped++; return; }
      (target ? target.rows : plan.main).push(row);
    });
    if (target) plan.archives.push(target);
  });
  return plan;
}

function migrationPreviewFromTasksFile(src) {
  const plan = tasksFilePlan_(src);
  return {
    main: plan.main.length, mainSkipped: plan.mainSkipped,
    archives: plan.archives.map(a => ({ name: a.name, count: a.rows.length, skipped: a.skipped }))
  };
}

function migrationApplyFromTasksFile(src) {
  const plan = tasksFilePlan_(src);
  plan.archives.forEach(a => { if (a.rows.length) writeArchiveSheet_(a.name, a.rows); });
  if (plan.main.length) applyMigration_({ rows: plan.main });
  return { added: plan.main.length, archived: plan.archives.reduce((n, a) => n + a.rows.length, 0) };
}

/** Архив: отдельный лист с теми же колонками. Без выпадающих списков и правил — просто для истории. */
function writeArchiveSheet_(name, rows) {
  withScriptLock_(() => {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, TASK_COLS).setValues(getTasksSheet().getRange(1, 1, 1, TASK_COLS).getValues())
        .setBackground('#2D3340').setFontColor('#FFFFFF').setFontWeight('bold').setFontFamily(FONT).setWrap(true);
      sh.setFrozenRows(1);
    }
    const start = Math.max(sh.getLastRow(), 1) + 1;
    const need = start + rows.length - 1 - sh.getMaxRows();
    if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
    const CHUNK = 1000;
    for (let off = 0; off < rows.length; off += CHUNK) {
      const part = rows.slice(off, off + CHUNK);
      sh.getRange(start + off, 1, part.length, TASK_COLS).setValues(part.map(p => p.values));
      sh.getRange(start + off, COL.SUBJECT, part.length, 1)
        .setRichTextValues(part.map(p => [buildSubjectRich_(String(p.values[COL.SUBJECT - 1]), p.link, p.link2)]));
    }
    const n = sh.getLastRow() - 1;
    sh.getRange(2, 1, n, TASK_COLS).setFontFamily(FONT).setFontSize(10).setVerticalAlignment('middle').setWrap(true);
    sh.getRange(2, COL.DATE, n, 1).setNumberFormat('dd.mm.yyyy');
    sh.getRange(2, COL.DUE, n, 1).setNumberFormat('dd.mm.yyyy');
    sh.getRange(2, COL.TOTAL, n, 1).setNumberFormat('#,##0.00');
    SpreadsheetApp.flush();
    sh.getRange(2, 1, n, TASK_COLS).sort({ column: COL.DATE, ascending: false });
    estimateLinksToLabels_(sh);
    logChange('Перенос истории', '', '', '', rows.length + ' строк на лист «' + name + '»');
  });
}

/** Дописывает строки под текущими задачами и оформляет лист. */
function applyMigration_(plan) {
  withScriptLock_(() => {
    const sh = getTasksSheet();
    const start = sh.getLastRow() + 1;
    const n = plan.rows.length;
    const need = start + n - 1 - sh.getMaxRows();
    if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);

    const CHUNK = 1000;
    for (let off = 0; off < n; off += CHUNK) {
      const part = plan.rows.slice(off, off + CHUNK);
      sh.getRange(start + off, 1, part.length, TASK_COLS).setValues(part.map(p => p.values));
      sh.getRange(start + off, COL.SUBJECT, part.length, 1)
        .setRichTextValues(part.map(p => [buildSubjectRich_(String(p.values[COL.SUBJECT - 1]), p.link, p.link2)]));
    }
    logChange('Перенос истории', '', '', '', n + ' задач из старой таблицы');

    // Весь лист — по дате, свежие сверху. Сортировка устойчивая: строки одной задачи
    // (одна дата) остаются рядом и в прежнем порядке. Строки без даты уходят вниз.
    SpreadsheetApp.flush();
    sh.getRange(2, 1, sh.getLastRow() - 1, TASK_COLS).sort({ column: COL.DATE, ascending: false });
  });
  designTasksSheet_(getTasksSheet());
}


/**
 * Приёмник для расширения «Сметы → таблица».
 * Принимает номер задачи, итог с НДС и ссылку на смету и пишет их в лист задач менеджеров,
 * а в «📝 Журнал» — что поменялось.
 *
 * Ещё умеет добавлять задачи менеджеров из расширения (вкладка «Задача»). Для этого файл должен
 * лежать в проекте самой таблицы: он вызывает те же функции, что и окно «➕ Добавить задачу».
 *
 * Лучше ставить ОТДЕЛЬНЫМ проектом Apps Script (script.google.com → «Создать проект»),
 * чтобы не трогать основной скрипт таблицы: тогда впишите ID таблицы ниже.
 * Если всё-таки кладёте в скрипт самой таблицы — ID можно оставить пустым,
 * но проверьте, что там ещё нет своей функции doPost.
 */

// ID таблицы — кусок адреса между /d/ и /edit. Нужен, только если Smeta.gs — отдельный проект, а не внутри таблицы.
const SMETA_SPREADSHEET_ID = '';

const SMETA_TASKS_SHEET = '📌 Задачи (менеджеры)';
const SMETA_LOG_SHEET = '📝 Журнал';
const SMETA_H = {
  task: '№ задачи',
  subject: 'Тема письма',
  link: 'Смета (ссылка)',
  total: 'Итого с НДС, ₽',
  contractor: 'Подрядчик',
  manager: 'Менеджер',
};

/** Запустите один раз из редактора: создаст токен и покажет его в журнале выполнения. */
function setupSmetaToken() {
  const token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('SMETA_TOKEN', token);
  Logger.log('Токен для расширения: ' + token);
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return smetaJson_({ ok: false, error: 'Некорректный запрос' });
  }
  const token = PropertiesService.getScriptProperties().getProperty('SMETA_TOKEN');
  if (!token || req.token !== token) return smetaJson_({ ok: false, error: 'Неверный токен' });

  // Чтение — без очереди: панель шлёт несколько запросов сразу, и они не должны ждать друг друга
  try {
    if (req.action === 'managers') { requireTableScript_(); return smetaJson_({ ok: true, managers: getListsData().managers }); }
    if (req.action === 'lookup') return smetaJson_(smetaLookup_(req));
    if (req.action === 'taskForm') return smetaJson_(taskForm_());
    if (req.action === 'previewTaskId') return smetaJson_(previewTaskId_(req));
    if (req.action === 'previewTaskIds') { requireTableScript_(); return smetaJson_({ ok: true, ids: previewNextTaskIds() }); }
    // Запросы из расширения выполняются от имени владельца таблицы, поэтому «кто я» берём только из настроек расширения
    if (req.action === 'checkDuplicates') { requireTableScript_(); return smetaJson_({ ok: true, duplicates: findDuplicateTasks(req.task || {}) }); }
    if (req.action === 'myTasks') { requireTableScript_(); return smetaJson_(Object.assign({ ok: true }, getMyOpenTasks(req.manager || '*', req.filter))); }
    if (req.action === 'searchTasks') return smetaJson_(searchTasks_(req));
    if (req.action === 'getTask') return smetaJson_(getTask_(req));
  } catch (err) {
    return smetaJson_({ ok: false, error: String(err && err.message || err) });
  }

  // Запись — по одному, чтобы два человека не получили один номер и не затёрли строку друг другу
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return smetaJson_({ ok: false, error: 'Таблица занята другим запросом, попробуйте ещё раз через минуту' });
  }
  SCRIPT_LOCK_HELD = true; // Код.gs не будет брать блокировку второй раз
  try {
    if (req.action === 'write') return smetaJson_(smetaWrite_(req));
    if (req.action === 'addTask') return smetaJson_(addTask_(req));
    if (req.action === 'setStatus') return smetaJson_(setStatus_(req));
    if (req.action === 'saveTask') return smetaJson_(saveTask_(req));
    return smetaJson_({ ok: false, error: 'Неизвестное действие' });
  } catch (err) {
    return smetaJson_({ ok: false, error: String(err && err.message || err) });
  } finally {
    SCRIPT_LOCK_HELD = false;
    lock.releaseLock();
  }
}

/** Проверка: открыть адрес веб-приложения в браузере — должно написать, что скрипт работает. */
function doGet() {
  return ContentService.createTextOutput('Скрипт работает. Адрес правильный — вставьте его в настройки расширения.');
}

function smetaLookup_(req) {
  const t = smetaFindTask_(req.task);
  if (!t) return { ok: true, found: false };
  const v = t.values;
  return {
    ok: true, found: true, row: t.row, matched: t.matched,
    subject: String(v[t.col.subject] || '').replace(/ \((?:доп\. ссылка|ссылка \d+)\)/g, ''), contractor: v[t.col.contractor], manager: v[t.col.manager],
    link: smetaCellLink_(t, v), total: v[t.col.total],
  };
}


/** Адрес сметы: ссылка под «Ссылка на смету номер» или сам текст ячейки, если там адрес. */
function smetaCellLink_(t, values) {
  const rt = t.sheet.getRange(t.row, t.col.link + 1).getRichTextValue();
  if (rt) {
    const url = rt.getLinkUrl() || rt.getRuns().map(function (r) { return r.getLinkUrl(); }).filter(Boolean)[0];
    if (url) return url;
  }
  const v = String(values[t.col.link] || '').trim();
  return /^(Смета|Ссылка на смету)( \[[^\]]*\]| \S+)?$/.test(v) ? '' : v;
}

function smetaWrite_(req) {
  const total = Number(req.total);
  const link = String(req.link || '').trim();
  if (!(total > 0)) return { ok: false, error: 'Сумма должна быть больше нуля' };
  if (!/^https?:\/\//i.test(link)) return { ok: false, error: 'Ссылка должна начинаться с http' };

  const t = smetaFindTask_(req.task);
  if (!t) return { ok: false, error: 'Задача ' + req.task + ' не найдена в листе «' + SMETA_TASKS_SHEET + '»' };

  const oldLink = smetaCellLink_(t, t.values);
  const oldTotal = t.values[t.col.total];
  if ((oldLink || oldTotal) && !req.overwrite) {
    return { ok: false, error: 'exists', link: oldLink, total: oldTotal };
  }

  const sheet = t.sheet;
  // В ячейке — короткое «Ссылка на смету LIT-26-2232», адрес — ссылкой под ним
  const rowId = String(t.values[t.col.task] || '').trim();
  sheet.getRange(t.row, t.col.link + 1).setRichTextValue(
    SpreadsheetApp.newRichTextValue().setText('Ссылка на смету' + (rowId ? ' ' + rowId : '')).setLinkUrl(link).build());
  sheet.getRange(t.row, t.col.total + 1).setValue(Math.round(total * 100) / 100);

  const who = 'Расширение смет' + (req.user ? ' (' + req.user + ')' : '');
  smetaLog_(who, t.taskId, SMETA_H.link, oldLink, link);
  smetaLog_(who, t.taskId, SMETA_H.total, oldTotal, total);
  return { ok: true, row: t.row };
}

function smetaFindTask_(task) {
  const taskId = String(task || '').trim().toUpperCase();
  if (!taskId || taskId.length > 100) throw new Error('Пустой номер задачи');

  const sheet = smetaSpreadsheet_().getSheetByName(SMETA_TASKS_SHEET);
  if (!sheet) throw new Error('Нет листа «' + SMETA_TASKS_SHEET + '»');
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(function (h) { return String(h).trim(); });

  const col = {};
  Object.keys(SMETA_H).forEach(function (k) {
    col[k] = header.indexOf(SMETA_H[k]);
    if (col[k] < 0) throw new Error('Не найдена колонка «' + SMETA_H[k] + '»');
  });

  const same = function (cell, id) { return String(cell).trim().toUpperCase() === id; };
  const found = function (i, id) {
    return { sheet: sheet, row: i + 1, values: data[i], col: col, taskId: taskId, matched: id };
  };

  // 1) номер в колонке «№ задачи»; 2) номер в любой другой колонке строки (например, «Тикет»);
  // 3) номер без последней части (из «ABC-12-345» — «ABC-12») в колонке «№ задачи».
  for (let i = 1; i < data.length; i++) {
    if (same(data[i][col.task], taskId)) return found(i, taskId);
  }
  for (let i = 1; i < data.length; i++) {
    if (data[i].some(function (cell) { return same(cell, taskId); })) return found(i, taskId);
  }
  const short = taskId.match(/^(.+)-\d+$/);
  if (short) {
    for (let i = 1; i < data.length; i++) {
      if (same(data[i][col.task], short[1])) return found(i, short[1]);
    }
  }
  return null;
}

function smetaLog_(who, taskId, field, before, after) {
  const log = smetaSpreadsheet_().getSheetByName(SMETA_LOG_SHEET);
  if (!log) return;
  log.appendRow([new Date(), who, 'Смета из расширения', taskId, field,
    before === undefined ? '' : before, after]);
}

/**
 * Скрипт лежит в самой таблице — берём её, ID не нужен. openById требует отдельного разрешения Google
 * («нет разрешения на вызов SpreadsheetApp.openById»), поэтому только для отдельного проекта.
 */
function smetaSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  if (!SMETA_SPREADSHEET_ID) throw new Error('Скрипт не в таблице: впишите SMETA_SPREADSHEET_ID в Smeta.gs');
  return SpreadsheetApp.openById(SMETA_SPREADSHEET_ID);
}

function smetaJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
 *  ДОБАВЛЕНИЕ ЗАДАЧИ ИЗ РАСШИРЕНИЯ
 *  Использует функции окна «➕ Добавить задачу» из Код.gs.
 * ============================================================ */

function requireTableScript_() {
  if (typeof submitNewTaskFromDialog !== 'function' || typeof getAddTaskFormLists !== 'function') {
    throw new Error('Добавление задач работает, только когда Smeta.gs лежит в проекте самой таблицы');
  }
}

/** Справочники для формы: подрядчики, продукты, языки и т. д. — те же, что в окне в таблице. */
function taskForm_() {
  requireTableScript_();
  return { ok: true, lists: getAddTaskFormLists() };
}

/** Какой номер получит задача у этого подрядчика (подсказка; настоящий выдаётся при добавлении). */
function previewTaskId_(req) {
  requireTableScript_();
  const contractor = String(req.contractor || '').trim();
  if (!contractor) return { ok: true, id: '' };
  return { ok: true, id: generateNextTaskId(getTasksSheet(), contractor) };
}

function addTask_(req) {
  requireTableScript_();
  const t = req.task || {};
  const str = function (v) { return String(v == null ? '' : v).trim(); };

  const task = {
    ticket: str(t.ticket), contractor: str(t.contractor), subject: str(t.subject),
    link: str(t.link), link2: str(t.link2), date: str(t.date), product: str(t.product),
    customer: str(t.customer), deadline: str(t.deadline), exactDeadline: str(t.exactDeadline),
    status: str(t.status), deliveryStatus: str(t.deliveryStatus), estimateLink: str(t.estimateLink),
    total: str(t.total), sp: str(t.sp), manager: str(t.manager), comment: str(t.comment),
    languages: (Array.isArray(t.languages) ? t.languages : []).map(str).filter(Boolean),
  };
  if (!task.contractor) return { ok: false, error: 'Выберите подрядчика' };
  if (!task.subject) return { ok: false, error: 'Впишите тему' };
  if (task.ticket && !/^LOCAL-\d+$/.test(task.ticket)) return { ok: false, error: 'Тикет должен быть вида LOCAL-1234' };
  // link2 — доп. ссылки на Band, по одной в строке
  [task.link, task.estimateLink].concat(task.link2.split(/[\s,;]+/)).forEach(function (u) {
    if (u && !/^https?:\/\//i.test(u)) throw new Error('Ссылка должна начинаться с http: ' + u);
  });

  // В «Журнал» пишем, что задачу добавили из расширения и кто.
  LOG_ACTOR = 'Расширение' + (req.user ? ' (' + req.user + ')' : '');
  try {
    const id = submitNewTaskFromDialog(task);
    return { ok: true, id: id };
  } finally {
    LOG_ACTOR = '';
  }
}

/** Быстрая смена статуса из вкладки «Мои задачи». */
/** Поиск любой задачи для правки: номер, тикет, тема, ник, комментарий. */
function searchTasks_(req) {
  requireTableScript_();
  const q = String(req.query || '').trim();
  if (!q) return { ok: true, tasks: [] }; // пустой поиск в таблице показывает «мои» владельца — здесь не нужно
  return { ok: true, tasks: searchTasks(q) };
}

/** Задача целиком для формы правки. Строки могли сдвинуться — ищем по номеру. */
function getTask_(req) {
  requireTableScript_();
  const sh = getTasksSheet();
  const id = String(req.id || '').trim();
  const row = id || req.origSubject != null ? locateTaskRow_(sh, Number(req.row), id, req.origSubject) : Number(req.row);
  return { ok: true, task: getTaskForEdit(row) };
}

function saveTask_(req) {
  requireTableScript_();
  const t = req.task || {};
  [t.link, t.estimateLink].concat(String(t.link2 || '').split(/[\s,;]+/)).forEach(function (u) {
    if (u && !/^https?:\/\//i.test(String(u))) throw new Error('Ссылка должна начинаться с http: ' + u);
  });
  LOG_ACTOR = 'Расширение' + (req.user ? ' (' + req.user + ')' : '');
  try {
    return saveTaskEdits(t);
  } finally {
    LOG_ACTOR = '';
  }
}

function setStatus_(req) {
  requireTableScript_();
  LOG_ACTOR = 'Расширение' + (req.user ? ' (' + req.user + ')' : '');
  try {
    return setTaskStatus(req.row, req.id, req.origSubject, req.status);
  } finally {
    LOG_ACTOR = '';
  }
}


// ==================== ОКНА (HTML) ====================
// Текст всех окон — чтобы не создавать 9 отдельных HTML-файлов.
const HTML_FILES = {
  "Общее": "<style>\n  :root {\n    --ink: #1F2328; --soft: #5F6B7A; --line: #E2E5E9; --field: #D5D9DF; --bg: #FFFFFF; --panel: #F6F7F9;\n    --head: #2D3340; --accent: #2563EB; --accent-soft: #EEF3FE;\n    --ok: #15803D; --ok-soft: #ECF7EF; --warn: #7F6000; --warn-soft: #FFF8E6; --err: #B42318; --err-soft: #FDEDEA;\n  }\n  * { box-sizing: border-box; }\n  [hidden] { display: none !important; }\n  html, body { margin: 0; background: var(--bg); color: var(--ink); font: 13px/1.45 Arial, sans-serif; }\n  body { padding: 14px 16px 16px; }\n  body.has-foot { padding-bottom: 76px; }\n  h2 { margin: 0 0 10px; font-size: 15px; }\n  h3 { margin: 16px 0 8px; font-size: 13.5px; }\n  .muted { color: var(--soft); }\n  .small { font-size: 12px; }\n\n  /* поля */\n  .stack { display: flex; flex-direction: column; gap: 11px; }\n  .field { display: flex; flex-direction: column; gap: 4px; min-width: 0; font-size: 12px; font-weight: 700; color: var(--soft); }\n  .req { color: #C0262D; }\n  input[type=text], input[type=url], input[type=date], input[type=number], select, textarea {\n    width: 100%; padding: 7px 10px; font: inherit; font-size: 13px; color: var(--ink); font-weight: 400;\n    border: 1px solid var(--field); border-radius: 7px; background: #fff; min-height: 34px;\n  }\n  textarea { resize: vertical; min-height: 60px; }\n  input:focus, select:focus, textarea:focus { outline: none; border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-soft); }\n  .two { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }\n  .three { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; }\n  .prefixed { display: flex; }\n  .prefixed span { display: grid; place-items: center; padding: 0 9px; border: 1px solid var(--field); border-right: none;\n    border-radius: 7px 0 0 7px; background: #F3F4F6; font-size: 12px; color: var(--soft); font-weight: 700; }\n  .prefixed input { border-radius: 0 7px 7px 0; }\n  .with-btn { display: flex; gap: 6px; }\n\n  /* таблетки-галочки (языки, люди) */\n  .chips { display: flex; flex-wrap: wrap; gap: 4px; }\n  .chip { position: relative; cursor: pointer; font-weight: 400; }\n  .chip input { position: absolute; opacity: 0; pointer-events: none; }\n  .chip span { display: inline-block; padding: 2px 8px; border-radius: 999px; border: 1px solid var(--field); font-size: 11.5px; line-height: 1.5; color: var(--ink); background: #fff; }\n  .chip input:checked + span { background: var(--accent); border-color: var(--accent); color: #fff; }\n  .chip input:focus-visible + span { outline: 2px solid var(--accent); outline-offset: 2px; }\n  .chip-group-title { font-size: 11px; color: var(--soft); font-weight: 400; margin: 4px 0 2px; }\n\n  /* кнопки */\n  button { font: inherit; cursor: pointer; }\n  .btn { padding: 9px 14px; border-radius: 8px; border: 1px solid var(--field); background: #fff; color: var(--ink); font-weight: 700; font-size: 13px; }\n  .btn:hover { background: var(--panel); }\n  .btn:disabled { opacity: .55; cursor: default; }\n  .btn.primary { background: var(--accent); border-color: var(--accent); color: #fff; }\n  .btn.primary:hover { filter: brightness(1.07); }\n  .btn.small { padding: 5px 10px; font-size: 12px; }\n  .btn.full { width: 100%; }\n  .link-danger { background: none; border: none; color: var(--err); font-size: 12px; padding: 6px; }\n  .foot { position: fixed; left: 0; right: 0; bottom: 0; display: flex; gap: 8px; align-items: center;\n    padding: 12px 16px; background: #fff; border-top: 1px solid var(--line); }\n  .foot .grow { flex: 1; }\n  .foot.col { flex-direction: column; align-items: stretch; gap: 2px; }\n\n  /* блоки */\n  .preview { padding: 10px 12px; border-radius: 8px; background: var(--accent-soft); border: 1px solid #D6E2FB; }\n  .preview .lbl { display: block; font-size: 11px; font-weight: 700; color: var(--soft); margin-bottom: 3px; }\n  .preview p { margin: 0; word-break: break-word; }\n  .card { padding: 10px 12px; border-radius: 8px; background: var(--panel); }\n  .card .t { font-weight: 700; }\n  details.more { border-top: 1px solid var(--line); padding-top: 8px; }\n  details.more > summary { cursor: pointer; font-weight: 700; color: var(--soft); font-size: 12.5px; }\n  details.more[open] > summary { margin-bottom: 10px; }\n  details.more > .stack { margin-bottom: 4px; }\n\n  /* выбор статуса одной кнопкой */\n  .seg { display: flex; flex-wrap: wrap; border: 1px solid var(--field); border-radius: 8px; overflow: hidden; }\n  .seg button { flex: 1 1 auto; padding: 7px 6px; border: none; border-right: 1px solid var(--field); background: #fff;\n    font-size: 12px; font-weight: 700; color: var(--soft); }\n  .seg button:last-child { border-right: none; }\n  .seg button.on { color: #111; }\n\n  /* результаты поиска */\n  .results { border: 1px solid var(--line); border-radius: 8px; overflow: hidden; max-height: 260px; overflow-y: auto; }\n  .results button { display: block; width: 100%; text-align: left; padding: 7px 10px; border: none;\n    border-bottom: 1px solid var(--line); background: #fff; font-size: 12.5px; color: var(--ink); }\n  .results button:last-child { border-bottom: none; }\n  .results button:hover, .results button.on { background: var(--accent-soft); }\n  .results b { margin-right: 4px; }\n  .results .d { color: var(--soft); font-size: 11px; }\n\n  /* сообщения */\n  .msg { min-height: 1em; font-size: 12.5px; color: var(--soft); }\n  .msg.ok { color: var(--ok); }\n  .msg.err { color: var(--err); font-weight: 700; }\n  .hint { padding: 8px 10px; border-radius: 7px; background: var(--warn-soft); color: var(--warn); font-size: 12.5px; }\n\n  /* отчёты */\n  .tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(110px, 1fr)); gap: 8px; margin: 10px 0; }\n  .tile { background: var(--panel); border-radius: 8px; padding: 10px; text-align: center; }\n  .tile .num { font-size: 20px; font-weight: 700; }\n  .tile .lbl { font-size: 11px; color: var(--soft); }\n  .tile.bad .num { color: var(--err); }\n  .group { border: 1px solid var(--line); border-radius: 8px; padding: 8px 12px; margin-bottom: 8px; }\n  .group h4 { margin: 0 0 6px; font-size: 13px; }\n  .line { display: flex; justify-content: space-between; gap: 10px; padding: 4px 0; border-bottom: 1px solid #F0F1F3; font-size: 12.5px; }\n  .line:last-child { border-bottom: none; }\n  .line a { color: var(--accent); text-decoration: none; }\n  .line .v { font-weight: 700; white-space: nowrap; }\n  .total { background: var(--panel); border-radius: 8px; padding: 10px; text-align: center; font-weight: 700; margin-top: 8px; }\n  pre.out { white-space: pre-wrap; background: var(--panel); border-radius: 8px; padding: 12px; font-size: 12px; margin: 8px 0; }\n  .badge { font-size: 11px; padding: 2px 8px; border-radius: 999px; font-weight: 700; white-space: nowrap; background: #EFEFEF; color: #434343; }\n  .b-done { background: #D9EAD3; color: #274E13; } .b-progress { background: #D6E4F0; color: #0B5394; } .b-review { background: #FFF2CC; color: #7F6000; }\n  table.kv { width: 100%; border-collapse: collapse; font-size: 12.5px; }\n  table.kv td { padding: 5px 6px; border-bottom: 1px solid #F0F1F3; }\n  table.kv td:last-child { text-align: right; font-weight: 700; }\n  hr.sep { border: none; border-top: 1px solid var(--line); margin: 18px 0; }\n</style>\n<script>\n  const $ = s => document.querySelector(s);\n  const $$ = s => [...document.querySelectorAll(s)];\n\n  function esc(s) {\n    return String(s == null ? '' : s).replace(/[&<>\"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '\"': '&quot;', \"'\": '&#39;' }[c]));\n  }\n\n  /** Вызов функции из Код.gs как промис; ошибки не теряются. */\n  function run(fn, ...args) {\n    return new Promise((resolve, reject) => {\n      google.script.run.withSuccessHandler(resolve).withFailureHandler(reject)[fn](...args);\n    });\n  }\n\n  function fillSelect(el, items, placeholder, selected) {\n    el.innerHTML = (placeholder == null ? '' : `<option value=\"\">${esc(placeholder)}</option>`) +\n      (items || []).map(v => `<option value=\"${esc(v)}\"${v === selected ? ' selected' : ''}>${esc(v)}</option>`).join('');\n  }\n\n  /** Галочки-таблетки. groups: [{ title, items }]; label — как подписать пункт. */\n  function renderChips(container, groups, checked, label) {\n    checked = checked || [];\n    container.innerHTML = groups.filter(g => g.items && g.items.length).map(g =>\n      (g.title ? `<div class=\"chip-group-title\">${esc(g.title)}</div>` : '') +\n      '<div class=\"chips\">' + g.items.map(v =>\n        `<label class=\"chip\"><input type=\"checkbox\" value=\"${esc(v)}\"${checked.includes(v) ? ' checked' : ''}><span>${esc(label ? label(v) : v)}</span></label>`\n      ).join('') + '</div>'\n    ).join('');\n  }\n\n  function chipValues(container) {\n    return [...container.querySelectorAll('input:checked')].map(cb => cb.value);\n  }\n\n  function setMsg(el, text, kind) {\n    el.textContent = text || '';\n    el.className = 'msg' + (kind ? ' ' + kind : '');\n  }\n\n  function todayIso() {\n    const d = new Date();\n    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');\n  }\n\n  function money(x) {\n    return Number(x || 0).toLocaleString('ru-RU', { maximumFractionDigits: 2 });\n  }\n\n  function copyText(text) {\n    return navigator.clipboard.writeText(text).catch(() => {\n      const t = document.createElement('textarea');\n      t.value = text; document.body.appendChild(t); t.select(); document.execCommand('copy'); t.remove();\n    });\n  }\n\n  /** Копирует HTML со ссылками — при вставке в Doc/почту/Band ссылки остаются кликабельными. */\n  function copyHtml(html) {\n    const holder = document.createElement('div');\n    holder.contentEditable = true;\n    holder.style.cssText = 'position:fixed;left:-9999px';\n    holder.innerHTML = html;\n    document.body.appendChild(holder);\n    const range = document.createRange();\n    range.selectNodeContents(holder);\n    const sel = window.getSelection();\n    sel.removeAllRanges(); sel.addRange(range);\n    document.execCommand('copy');\n    sel.removeAllRanges(); holder.remove();\n  }\n</script>\n",
  "AddTaskDialog": "<!DOCTYPE html>\n<html>\n<head>\n<base target=\"_top\">\n<?!= include('Общее') ?>\n<style>\n  /* Тема письма и «Скопировать» всегда видны сверху, даже при прокрутке */\n  .preview { position: sticky; top: 0; z-index: 5; margin: -14px -16px 0; border-radius: 0; border-width: 0 0 1px; padding: 10px 16px;\n    display: flex; gap: 10px; align-items: flex-start; box-shadow: 0 2px 6px rgb(0 0 0 / .06); }\n  .preview .btn { flex: none; }\n  .preview .txt { flex: 1; min-width: 0; }\n  .tpl-actions { display: flex; flex-wrap: wrap; gap: 2px 14px; font-weight: 400; }\n  .link-btn { background: none; border: none; padding: 2px 0; font-size: 12px; color: var(--soft); text-decoration: underline; text-underline-offset: 2px; }\n  .link-btn:hover { color: var(--ink); }\n</style>\n</head>\n<body class=\"has-foot\">\n  <div class=\"stack\">\n    <div class=\"preview\">\n      <div class=\"txt\"><span class=\"lbl\">ТЕМА ПИСЬМА</span><p id=\"preview\">Выберите подрядчика — появится номер и тема</p></div>\n      <button class=\"btn small\" id=\"copyBtn\" type=\"button\">Скопировать</button>\n    </div>\n\n    <div class=\"field\">Шаблон\n      <div class=\"with-btn\">\n        <select id=\"template\" style=\"flex:1;min-width:0\"><option value=\"\">— без шаблона —</option></select>\n        <button class=\"btn\" id=\"saveTpl\" type=\"button\" title=\"Запомнить подрядчика, тикет, продукт, языки и т. д.\">Сохранить</button>\n      </div>\n      <div class=\"tpl-actions\">\n        <button class=\"link-btn\" id=\"delTpl\" type=\"button\" hidden>Удалить шаблон</button>\n        <button class=\"link-btn\" id=\"exportTpl\" type=\"button\">Скачать шаблоны</button>\n        <button class=\"link-btn\" id=\"importTpl\" type=\"button\">Загрузить из файла</button>\n        <input type=\"file\" id=\"importFile\" accept=\".json,application/json\" hidden>\n      </div>\n    </div>\n\n    <div class=\"two\">\n      <label class=\"field\"><span>Подрядчик <span class=\"req\">*</span></span><select id=\"contractor\"></select></label>\n      <label class=\"field\">Тикет\n        <span class=\"prefixed\"><span>LOCAL-</span><input type=\"text\" id=\"ticketNum\" inputmode=\"numeric\" placeholder=\"1234\"></span>\n      </label>\n    </div>\n\n    <label class=\"field\"><span>Тема <span class=\"req\">*</span></span><input type=\"text\" id=\"subject\" placeholder=\"Только суть — номер и коды добавятся сами\"></label>\n    <label class=\"field\">Ссылка на Band<input type=\"url\" id=\"link\" placeholder=\"https://band.wb.ru/…\"></label>\n    <label class=\"field\">Ещё ссылки на Band <span class=\"muted\" style=\"font-weight:400\">если переписок несколько</span>\n      <textarea id=\"link2\" rows=\"1\" style=\"min-height:0;resize:none;overflow:hidden\" placeholder=\"каждая с новой строки — в таблице станут «(ссылка 2)», «(ссылка 3)» после темы\"></textarea></label>\n\n    <div class=\"two\">\n      <label class=\"field\">Продукт<select id=\"product\"></select></label>\n      <label class=\"field\">Дата<input type=\"date\" id=\"date\"></label>\n    </div>\n    <div class=\"two\">\n      <label class=\"field\">Дедлайн<select id=\"deadline\"></select></label>\n      <label class=\"field\">Срок сдачи<input type=\"date\" id=\"exactDeadline\"></label>\n    </div>\n    <label class=\"field\">Ник заказчика<input type=\"text\" id=\"customer\" placeholder=\"@nick\"></label>\n\n    <div class=\"field\">Языки<div id=\"langs\"></div></div>\n\n    <div class=\"two\">\n      <label class=\"field\">Менеджер<select id=\"manager\"></select></label>\n      <label class=\"field\">Статус<select id=\"status\"></select></label>\n    </div>\n\n    <details class=\"more\">\n      <summary>Ещё поля: смета, итого, SP, статус отдачи, комментарий</summary>\n      <div class=\"stack\">\n        <label class=\"field\">Смета (ссылка)<input type=\"url\" id=\"estimateLink\"></label>\n        <div class=\"three\">\n          <label class=\"field\">Итого с НДС, ₽<input type=\"number\" id=\"total\" step=\"0.01\" min=\"0\"></label>\n          <label class=\"field\">SP<input type=\"number\" id=\"sp\" step=\"1\" min=\"0\"></label>\n          <label class=\"field\">Статус отдачи<select id=\"deliveryStatus\"></select></label>\n        </div>\n        <label class=\"field\">Комментарий<input type=\"text\" id=\"comment\"></label>\n      </div>\n    </details>\n    <div id=\"msg\" class=\"msg\"></div>\n  </div>\n\n  <div class=\"foot\">\n    <button class=\"btn primary grow\" id=\"submitBtn\" type=\"button\">Добавить задачу</button>\n  </div>\n\n<script>\n  let LISTS = null;\n  let nextId = '';\n\n  let TEMPLATES = [];\n  $('#date').value = todayIso();\n\n  run('getAddTaskFormLists').then(d => {\n    LISTS = d;\n    fillSelect($('#contractor'), d.contractors, 'Выберите…');\n    fillSelect($('#product'), d.products, '—');\n    fillSelect($('#deadline'), d.deadlines, '—');\n    fillSelect($('#status'), d.statuses, '—', d.statuses.includes('Принято') ? 'Принято' : '');\n    fillSelect($('#deliveryStatus'), d.deliveryStatuses, '—');\n    fillSelect($('#manager'), d.managers, '—', d.currentManager);\n    renderChips($('#langs'), [\n      { title: '', items: d.languages.regular },\n      { title: 'ШТАТ', items: d.languages.shtat },\n      { title: 'Редкие', items: d.languages.rare }\n    ], [], v => v.replace(/^ШТАТ\\s+/i, ''));\n    updatePreview();\n    TEMPLATES = d.templates || [];\n    renderTemplates();\n  }).catch(e => setMsg($('#msg'), 'Не получилось загрузить списки: ' + e.message, 'err'));\n\n  function checkedLangs() { return chipValues($('#langs')); }\n\n  function buildSubject(id) {\n    const contractor = $('#contractor').value, product = $('#product').value;\n    const codes = checkedLangs().map(l => LISTS.langCodes[l]).filter(Boolean);\n    const prefix = `[${id}]` + (contractor ? `[${contractor}]` : '') + codes.map(c => `[${c}]`).join('') + (product ? `[${product}]` : '');\n    const s = subjectText();\n    return s ? prefix + ' ' + s : prefix;\n  }\n\n  /** «Перевод строчек … от» → «… от 01.10»: дата из поля «Дата» (по умолчанию сегодня). */\n  function withDate(s) {\n    const [, m, d] = $('#date').value.split('-');\n    return d && m ? s.replace(/(^|\\s)(от)\\s*$/i, '$1$2 ' + d + '.' + m) : s;\n  }\n  function subjectText() { return withDate($('#subject').value.trim()); }\n\n  // Ушли из поля — «от» в конце сразу превращается в «от 01.10»\n  $('#subject').addEventListener('blur', () => {\n    const v = subjectText();\n    if (v !== $('#subject').value.trim()) { $('#subject').value = v; updatePreview(); }\n  });\n  // Поменяли дату — «от 22.09» в теме меняется на неё же\n  $('#date').addEventListener('change', () => { $('#subject').value = withToday($('#subject').value); updatePreview(); });\n\n  function updatePreview() {\n    if (!LISTS) return;\n    $('#preview').textContent = $('#contractor').value ? buildSubject(nextId || '…') : 'Выберите подрядчика — появится номер и тема';\n  }\n\n  let idSeq = 0;\n  function refreshId() {\n    const seq = ++idSeq;\n    nextId = '';\n    updatePreview();\n    const c = $('#contractor').value;\n    if (!c) return;\n    run('previewNextTaskId', c).then(id => { if (seq === idSeq) { nextId = id; updatePreview(); } })\n      .catch(e => setMsg($('#msg'), 'Не получилось узнать номер: ' + e.message, 'err'));\n  }\n\n  $('#contractor').addEventListener('change', refreshId);\n  ['subject', 'product'].forEach(id => { $('#' + id).addEventListener('input', updatePreview); $('#' + id).addEventListener('change', updatePreview); });\n  $('#langs').addEventListener('change', updatePreview);\n\n  /** «Новые строчки от 22.09» → «Новые строчки от <сегодня>». */\n  function withToday(title) {\n    const [, m, d] = $('#date').value.split('-');\n    // \\b не работает с русскими буквами — границу слова задаём явно\n    return d && m ? title.replace(/(^|\\s)(от\\s+)\\d{1,2}\\.\\d{1,2}(\\.\\d{2,4})?/i, '$1$2' + d + '.' + m) : title;\n  }\n\n  // ---------- Шаблоны: свои у каждого; файл тот же, что у расширения ----------\n  function renderTemplates(selectName) {\n    TEMPLATES.sort((a, b) => a.name.localeCompare(b.name, 'ru'));\n    $('#template').innerHTML = '<option value=\"\">— без шаблона —</option>' +\n      TEMPLATES.map((t, i) => `<option value=\"${i}\">${esc(t.name)}</option>`).join('');\n    const i = TEMPLATES.findIndex(t => t.name === selectName);\n    $('#template').value = i === -1 ? '' : String(i);\n    $('#delTpl').hidden = !$('#template').value;\n  }\n\n  function storeTemplates(selectName, okText) {\n    return run('saveTaskTemplates', TEMPLATES).then(list => {\n      TEMPLATES = list;\n      renderTemplates(selectName);\n      if (okText) setMsg($('#msg'), okText, 'ok');\n    }).catch(e => setMsg($('#msg'), 'Шаблоны не сохранились: ' + e.message, 'err'));\n  }\n\n  $('#template').addEventListener('change', () => {\n    $('#delTpl').hidden = !$('#template').value;\n    const t = TEMPLATES[$('#template').value];\n    if (!t || !LISTS) return;\n    $('#contractor').value = t.contractor; refreshId();\n    $('#ticketNum').value = t.ticket.replace(/^LOCAL-/i, '');\n    $('#product').value = t.product;\n    $('#customer').value = t.customer;\n    $('#deadline').value = t.deadline;\n    $('#comment').value = t.comment;\n    $('#subject').value = withDate(withToday(t.subject));\n    $('#link').value = '';\n    $$('#langs input').forEach(cb => { cb.checked = t.languages.includes(cb.value); });\n    setMsg($('#msg'), 'Заполнено по шаблону «' + t.name + '». Вставьте ссылку на Band и проверьте тему.');\n    updatePreview();\n  });\n\n  $('#saveTpl').addEventListener('click', () => {\n    if (!$('#contractor').value && !$('#subject').value.trim()) return setMsg($('#msg'), 'Заполните хотя бы подрядчика или тему — их и запомнит шаблон', 'err');\n    const current = TEMPLATES[$('#template').value];\n    const name = (prompt('Название шаблона:', current ? current.name : $('#subject').value.trim()) || '').trim();\n    if (!name) return;\n    const existing = TEMPLATES.findIndex(t => t.name === name);\n    if (existing !== -1 && !confirm('Шаблон «' + name + '» уже есть. Заменить?')) return;\n    const num = $('#ticketNum').value.trim().replace(/^local-?\\s*/i, '');\n    const tpl = {\n      name: name, contractor: $('#contractor').value, ticket: num ? 'LOCAL-' + num : '', subject: $('#subject').value.trim(),\n      product: $('#product').value, customer: $('#customer').value.trim(), deadline: $('#deadline').value,\n      comment: $('#comment').value.trim(), languages: checkedLangs()\n    };\n    if (existing !== -1) TEMPLATES[existing] = tpl; else TEMPLATES.push(tpl);\n    storeTemplates(name, 'Шаблон «' + name + '» сохранён');\n  });\n\n  $('#delTpl').addEventListener('click', () => {\n    const t = TEMPLATES[$('#template').value];\n    if (!t || !confirm('Удалить шаблон «' + t.name + '»?')) return;\n    TEMPLATES.splice(Number($('#template').value), 1);\n    storeTemplates('', 'Шаблон «' + t.name + '» удалён');\n  });\n\n  $('#exportTpl').addEventListener('click', () => {\n    if (!TEMPLATES.length) return setMsg($('#msg'), 'Шаблонов пока нет: заполните форму и нажмите «Сохранить» рядом с «Шаблон»', 'err');\n    const blob = new Blob([JSON.stringify({ type: 'wb-task-templates', version: 1, templates: TEMPLATES }, null, 2)], { type: 'application/json' });\n    const a = document.createElement('a');\n    a.href = URL.createObjectURL(blob);\n    a.download = 'task-templates.json';\n    document.body.appendChild(a); a.click(); a.remove();\n    setMsg($('#msg'), 'Скачано шаблонов: ' + TEMPLATES.length + '. Файл можно отправить коллеге — пусть нажмёт «Загрузить из файла».', 'ok');\n  });\n\n  $('#importTpl').addEventListener('click', () => $('#importFile').click());\n  $('#importFile').addEventListener('change', () => {\n    const file = $('#importFile').files[0];\n    $('#importFile').value = '';\n    if (!file) return;\n    file.text().then(text => {\n      const data = JSON.parse(text);\n      const list = (Array.isArray(data) ? data : data.templates || []).filter(t => t && typeof t.name === 'string' && t.name.trim());\n      if (!list.length) throw new Error('в файле нет шаблонов');\n      let replaced = 0;\n      list.forEach(t => {\n        const i = TEMPLATES.findIndex(x => x.name === t.name.trim());\n        if (i !== -1) { TEMPLATES[i] = t; replaced++; } else TEMPLATES.push(t);\n      });\n      return storeTemplates('', 'Загружено шаблонов: ' + list.length + (replaced ? ' (заменено с тем же названием: ' + replaced + ')' : ''));\n    }).catch(e => setMsg($('#msg'), 'Не получилось загрузить: ' + e.message, 'err'));\n  });\n\n  // «Ещё ссылки» растёт по мере вставки: каждая ссылка — своя строка\n  $('#link2').addEventListener('input', () => {\n    const t = $('#link2');\n    t.style.height = 'auto';\n    t.style.height = (t.scrollHeight + 2) + 'px';\n  });\n\n  $('#copyBtn').addEventListener('click', () => {\n    copyText($('#preview').textContent).then(() => {\n      $('#copyBtn').textContent = 'Скопировано ✓';\n      setTimeout(() => { $('#copyBtn').textContent = 'Скопировать'; }, 1500);\n    });\n  });\n\n  $('#submitBtn').addEventListener('click', () => {\n    if (!$('#contractor').value) return setMsg($('#msg'), 'Выберите подрядчика', 'err');\n    if (!$('#subject').value.trim()) return setMsg($('#msg'), 'Впишите тему', 'err');\n    const ticket = $('#ticketNum').value.trim().replace(/^local-?\\s*/i, '');\n    if (ticket && !/^\\d+$/.test(ticket)) return setMsg($('#msg'), 'В тикете — только номер, например 1234', 'err');\n    const task = {\n      ticket: ticket ? 'LOCAL-' + ticket : '', contractor: $('#contractor').value, subject: subjectText(),\n      link: $('#link').value, link2: $('#link2').value, date: $('#date').value, product: $('#product').value,\n      customer: $('#customer').value, deadline: $('#deadline').value, exactDeadline: $('#exactDeadline').value,\n      status: $('#status').value, deliveryStatus: $('#deliveryStatus').value, estimateLink: $('#estimateLink').value,\n      total: $('#total').value, sp: $('#sp').value, manager: $('#manager').value, comment: $('#comment').value,\n      languages: checkedLangs()\n    };\n    const btn = $('#submitBtn');\n    btn.disabled = true; btn.textContent = 'Проверяю…';\n    run('findDuplicateTasks', task).then(dups => {\n      if (dups.length && !confirm('Похожая задача уже есть:\\n\\n' +\n          dups.map(d => `${d.id || '—'} · ${d.title} · ${d.date} · ${d.manager} (${d.why})`).join('\\n') +\n          '\\n\\nВсё равно добавить?')) {\n        throw new Error('Не добавлено — похожая задача уже есть.');\n      }\n      btn.textContent = 'Добавляю…';\n      return run('submitNewTaskFromDialog', task);\n    }).then(id => {\n      const subject = buildSubject(id);\n      copyText(subject);\n      setMsg($('#msg'), 'Добавлено: ' + id + '. Тема скопирована в буфер. Закрываю…', 'ok');\n      $('#preview').textContent = subject;\n      setTimeout(() => google.script.host.close(), 1600);\n    }).catch(e => {\n      setMsg($('#msg'), e.message.startsWith('Не добавлено') ? e.message : 'Не добавилось: ' + e.message, 'err');\n      btn.disabled = false; btn.textContent = 'Добавить задачу';\n    });\n  });\n</script>\n</body>\n</html>\n",
  "SearchEditSidebar": "<!DOCTYPE html>\n<html>\n<head>\n<base target=\"_top\">\n<?!= include('Общее') ?>\n</head>\n<body class=\"has-foot\">\n  <div class=\"stack\">\n    <label class=\"field\">Найти<input type=\"text\" id=\"q\" placeholder=\"Номер, тема, тикет или ник\"></label>\n    <div class=\"small muted\" id=\"qHint\">Ваши задачи:</div>\n    <div class=\"results\" id=\"results\"></div>\n\n    <div id=\"editor\" class=\"stack\" hidden>\n      <div class=\"card\">\n        <div class=\"t\" id=\"cardTitle\"></div><div class=\"small muted\" id=\"cardMeta\"></div>\n        <button class=\"btn small\" id=\"copySubject\" type=\"button\" style=\"margin-top:6px\">Скопировать тему</button>\n      </div>\n\n      <div class=\"field\">Статус<div class=\"seg\" id=\"statusSeg\"></div></div>\n      <div class=\"two\">\n        <label class=\"field\">Срок сдачи<input type=\"date\" id=\"f_exactDeadline\"></label>\n        <label class=\"field\">SP<input type=\"number\" id=\"f_sp\" step=\"1\" min=\"0\"></label>\n      </div>\n      <div class=\"field\">Языки<div id=\"f_langs\"></div></div>\n\n      <details class=\"more\">\n        <summary>Задача: тема, ссылки, тикет, дата, продукт, ник</summary>\n        <div class=\"stack\">\n          <label class=\"field\">Тема письма<input type=\"text\" id=\"f_subject\"></label>\n          <label class=\"field\">Ссылка на Band<input type=\"url\" id=\"f_link\"></label>\n          <label class=\"field\">Ещё ссылки на Band<textarea id=\"f_link2\" rows=\"2\" placeholder=\"каждая с новой строки\"></textarea></label>\n          <div class=\"two\">\n            <label class=\"field\">Тикет<input type=\"text\" id=\"f_ticket\" placeholder=\"LOCAL-1234\"></label>\n            <label class=\"field\">Дата<input type=\"date\" id=\"f_date\"></label>\n          </div>\n          <label class=\"field\">Продукт<select id=\"f_product\"></select></label>\n          <label class=\"field\">Ник заказчика<input type=\"text\" id=\"f_customer\"></label>\n          <label class=\"field\">Дедлайн<select id=\"f_deadline\"></select></label>\n        </div>\n      </details>\n      <details class=\"more\">\n        <summary>Смета и деньги</summary>\n        <div class=\"stack\">\n          <label class=\"field\">Смета (ссылка)<input type=\"url\" id=\"f_estimateLink\"></label>\n          <div class=\"two\">\n            <label class=\"field\">Итого с НДС, ₽<input type=\"number\" id=\"f_total\" step=\"0.01\" min=\"0\"></label>\n            <label class=\"field\">Статус отдачи<select id=\"f_deliveryStatus\"></select></label>\n          </div>\n        </div>\n      </details>\n      <details class=\"more\">\n        <summary>Подрядчик, менеджер, комментарий</summary>\n        <div class=\"stack\">\n          <div class=\"two\">\n            <label class=\"field\">Подрядчик<select id=\"f_contractor\"></select></label>\n            <label class=\"field\">Менеджер<select id=\"f_manager\"></select></label>\n          </div>\n          <label class=\"field\">Комментарий<textarea id=\"f_comment\"></textarea></label>\n        </div>\n      </details>\n      <details class=\"more\" id=\"historyBox\">\n        <summary>История изменений</summary>\n        <div id=\"history\" class=\"small muted\"></div>\n      </details>\n    </div>\n    <div id=\"msg\" class=\"msg\"></div>\n  </div>\n\n  <div class=\"foot col\" id=\"foot\" hidden>\n    <button class=\"btn primary full\" id=\"saveBtn\" type=\"button\">Сохранить</button>\n    <button class=\"link-danger\" id=\"deleteBtn\" type=\"button\">Удалить задачу</button>\n  </div>\n\n<script>\n  const STATUS_COLORS = { 'Принято': '#EDEDED', 'В работе': '#D6E4F0', 'Отдано': '#D9EAD3', 'Отменено': '#F4CCCC', 'Холд': '#FFF2CC' };\n  let LISTS = null, current = null, status = '';\n\n  run('getAddTaskFormLists').then(d => {\n    LISTS = d;\n    fillSelect($('#f_product'), d.products, '—');\n    fillSelect($('#f_deadline'), d.deadlines, '—');\n    fillSelect($('#f_deliveryStatus'), d.deliveryStatuses, '—');\n    fillSelect($('#f_contractor'), d.contractors, '—');\n    fillSelect($('#f_manager'), d.managers, '—');\n    search('');\n  }).catch(e => setMsg($('#msg'), 'Не получилось загрузить списки: ' + e.message, 'err'));\n\n  // ---------- поиск ----------\n  let timer = null, seq = 0;\n  $('#q').addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(() => search($('#q').value), 350); });\n\n  function search(q) {\n    const my = ++seq;\n    $('#qHint').textContent = q.trim() ? 'Найдено:' : 'Ваши задачи (или последние, если вашей почты нет в «Списках»):';\n    run('searchTasks', q).then(list => {\n      if (my !== seq) return;\n      $('#results').innerHTML = list.length ? list.map(r =>\n        `<button type=\"button\" data-row=\"${r.row}\"${current && current.row === r.row ? ' class=\"on\"' : ''}>` +\n        `<b>${esc(r.id || '—')}</b>${esc(r.title)} <span class=\"d\">· ${esc(r.date)}${r.status ? ' · ' + esc(r.status) : ''}</span></button>`\n      ).join('') : '<div class=\"small muted\" style=\"padding:8px 10px\">Ничего не найдено</div>';\n    }).catch(e => setMsg($('#msg'), 'Поиск не сработал: ' + e.message, 'err'));\n  }\n\n  $('#results').addEventListener('click', e => {\n    const b = e.target.closest('button[data-row]');\n    if (!b) return;\n    $$('#results button').forEach(x => x.classList.toggle('on', x === b));\n    open(Number(b.dataset.row));\n  });\n\n  // ---------- карточка ----------\n  function open(row) {\n    setMsg($('#msg'), 'Открываю…');\n    run('getTaskForEdit', row).then(fill).catch(e => setMsg($('#msg'), 'Не открылась: ' + e.message, 'err'));\n  }\n\n  function renderStatus() {\n    $('#statusSeg').innerHTML = LISTS.statuses.map(s =>\n      `<button type=\"button\" data-s=\"${esc(s)}\" class=\"${s === status ? 'on' : ''}\" style=\"${s === status ? 'background:' + (STATUS_COLORS[s] || '#EEF3FE') : ''}\">${esc(s)}</button>`).join('');\n  }\n  $('#statusSeg').addEventListener('click', e => {\n    const b = e.target.closest('button[data-s]');\n    if (!b) return;\n    status = status === b.dataset.s ? '' : b.dataset.s;\n    renderStatus();\n  });\n\n  function fill(t) {\n    current = t;\n    status = t.status;\n    $('#cardTitle').textContent = (t.id || 'без номера') + (t.ticket ? ' · ' + t.ticket : '');\n    $('#cardMeta').textContent = [t.contractor, t.manager, t.product].filter(Boolean).join(' · ');\n    renderStatus();\n    const set = (id, v) => { $('#f_' + id).value = v == null ? '' : v; };\n    ['exactDeadline', 'sp', 'subject', 'link', 'link2', 'ticket', 'date', 'product', 'customer', 'deadline',\n     'estimateLink', 'total', 'deliveryStatus', 'contractor', 'manager', 'comment'].forEach(k => set(k, t[k]));\n\n    const langs = t.languages.split(',').map(s => s.trim()).filter(Boolean);\n    const known = [].concat(LISTS.languages.regular, LISTS.languages.shtat, LISTS.languages.rare);\n    renderChips($('#f_langs'), [\n      { title: '', items: LISTS.languages.regular },\n      { title: 'ШТАТ', items: LISTS.languages.shtat },\n      { title: 'Редкие', items: LISTS.languages.rare },\n      { title: 'Другие', items: langs.filter(l => !known.includes(l)) }\n    ], langs, v => v.replace(/^ШТАТ\\s+/i, ''));\n\n    $('#editor').hidden = false;\n    $('#foot').hidden = false;\n    setMsg($('#msg'), '');\n    $('#history').textContent = 'Загружаю…';\n    run('getTaskHistory', t.id).then(list => {\n      $('#history').innerHTML = list.length\n        ? list.map(h => `<div style=\"padding:3px 0;border-bottom:1px dashed #E2E5E9\"><b>${esc(h.date)}</b> — ${esc(h.who)} — ${esc(h.action)}${h.field ? ' (' + esc(h.field) + ')' : ''}</div>`).join('')\n        : 'Изменений пока не было.';\n    }).catch(() => { $('#history').textContent = 'Историю загрузить не получилось.'; });\n  }\n\n  $('#copySubject').addEventListener('click', () => {\n    if (!current) return;\n    copyText(current.subject).then(() => {\n      $('#copySubject').textContent = 'Скопировано ✓';\n      setTimeout(() => { $('#copySubject').textContent = 'Скопировать тему'; }, 1500);\n    });\n  });\n\n  // ---------- сохранить / удалить ----------\n  $('#saveBtn').addEventListener('click', () => {\n    if (!current) return;\n    const get = id => $('#f_' + id).value;\n    const task = {\n      row: current.row, id: current.id, origSubject: current.origSubject, status: status,\n      languages: chipValues($('#f_langs')).join(', ')\n    };\n    ['exactDeadline', 'sp', 'subject', 'link', 'link2', 'ticket', 'date', 'product', 'customer', 'deadline',\n     'estimateLink', 'total', 'deliveryStatus', 'contractor', 'manager', 'comment'].forEach(k => { task[k] = get(k); });\n    const btn = $('#saveBtn');\n    btn.disabled = true;\n    setMsg($('#msg'), 'Сохраняю…');\n    run('saveTaskEdits', task).then(r => {\n      setMsg($('#msg'), 'Сохранено ✓', 'ok');\n      return run('getTaskForEdit', r.row).then(fill).then(() => setMsg($('#msg'), 'Сохранено ✓', 'ok'));\n    }).catch(e => setMsg($('#msg'), 'Не сохранилось: ' + e.message, 'err'))\n      .finally(() => { btn.disabled = false; });\n  });\n\n  $('#deleteBtn').addEventListener('click', () => {\n    if (!current) return;\n    if (!confirm('Удалить задачу ' + (current.id || '') + '? Отменить будет нельзя.')) return;\n    run('deleteTask', current.row, current.id, current.origSubject).then(() => {\n      current = null;\n      $('#editor').hidden = true; $('#foot').hidden = true;\n      setMsg($('#msg'), 'Удалено.', 'ok');\n      search($('#q').value);\n    }).catch(e => setMsg($('#msg'), 'Не удалилось: ' + e.message, 'err'));\n  });\n</script>\n</body>\n</html>\n",
  "DashboardSidebar": "<!DOCTYPE html>\n<html>\n<head>\n<base target=\"_top\">\n<?!= include('Общее') ?>\n<script src=\"https://www.gstatic.com/charts/loader.js\"></script>\n<style>\n  .row { display: flex; gap: 8px; align-items: end; }\n  .row .field { flex: 1; }\n  .chart { width: 100%; height: 240px; margin-bottom: 8px; }\n  .overdue .line { font-size: 12px; }\n</style>\n</head>\n<body>\n  <div class=\"row\">\n    <label class=\"field\">Год<select id=\"year\"></select></label>\n    <label class=\"field\">Месяц<select id=\"month\">\n      <option value=\"\">Весь год</option><option value=\"1\">Январь</option><option value=\"2\">Февраль</option><option value=\"3\">Март</option>\n      <option value=\"4\">Апрель</option><option value=\"5\">Май</option><option value=\"6\">Июнь</option><option value=\"7\">Июль</option>\n      <option value=\"8\">Август</option><option value=\"9\">Сентябрь</option><option value=\"10\">Октябрь</option>\n      <option value=\"11\">Ноябрь</option><option value=\"12\">Декабрь</option></select></label>\n    <button class=\"btn primary\" id=\"reload\" type=\"button\">Обновить</button>\n  </div>\n  <div id=\"msg\" class=\"msg\"></div>\n\n  <div class=\"tiles\" id=\"tiles\"></div>\n  <div id=\"overdueBox\" class=\"group overdue\" hidden><h4>Просрочено сейчас</h4><div id=\"overdue\"></div></div>\n\n  <h3>По менеджерам, SP</h3><div class=\"chart\" id=\"c_managers\"></div>\n  <h3>По менеджерам, ₽</h3><div class=\"chart\" id=\"c_managers_money\"></div>\n  <h3>По подрядчикам, ₽</h3><div class=\"chart\" id=\"c_contractors\"></div>\n  <h3>По продуктам, задач</h3><div class=\"chart\" id=\"c_products\"></div>\n  <h3>По языкам, упоминаний</h3><div class=\"chart\" id=\"c_langs\"></div>\n  <h3>По месяцам (последние 12): задачи и SP</h3><div class=\"chart\" id=\"c_months\"></div>\n  <h3>По годам: задачи и SP</h3><div class=\"chart\" id=\"c_years\"></div>\n\n<script>\n  const COLORS = ['#2563EB', '#94A3B8'];\n  run('getDashboardYears').then(years => {\n    fillSelect($('#year'), years.map(String), 'Всё время');\n  });\n  // Цифры показываем сразу; графики — когда загрузится библиотека Google Charts\n  let lastData = null, chartsReady = false;\n  if (window.google && google.charts) {\n    google.charts.load('current', { packages: ['corechart'] });\n    google.charts.setOnLoadCallback(() => { chartsReady = true; if (lastData) drawCharts(lastData); });\n  } else {\n    $$('.chart').forEach(el => { el.innerHTML = '<div class=\"small muted\">Графики не загрузились — проверьте интернет и нажмите «Обновить».</div>'; el.style.height = 'auto'; });\n  }\n  load();\n  $('#reload').addEventListener('click', load);\n  $('#year').addEventListener('change', load);\n  $('#month').addEventListener('change', () => { if ($('#year').value) load(); });\n\n  function load() {\n    setMsg($('#msg'), 'Считаю…');\n    run('getDashboardData', $('#year').value, $('#year').value ? $('#month').value : '')\n      .then(render).then(() => setMsg($('#msg'), ''))\n      .catch(e => setMsg($('#msg'), 'Не получилось посчитать: ' + e.message, 'err'));\n  }\n\n  function tile(num, lbl, bad) {\n    return `<div class=\"tile${bad ? ' bad' : ''}\"><div class=\"num\">${esc(num)}</div><div class=\"lbl\">${esc(lbl)}</div></div>`;\n  }\n\n  function render(d) {\n    $('#tiles').innerHTML = tile(d.totalTasks, 'задач') + tile(d.totalSp, 'SP') + tile(money(d.totalMoney), '₽ с НДС') +\n      tile(d.overdueCount, 'просрочено сейчас', d.overdueCount > 0);\n    $('#overdueBox').hidden = !d.overdue.length;\n    $('#overdue').innerHTML = d.overdue.map(o =>\n      `<div class=\"line\"><span><b>${esc(o.id)}</b> ${esc(o.title)}</span><span class=\"v\">${esc(o.due)} · ${esc(o.manager)}</span></div>`).join('');\n    lastData = d;\n    if (chartsReady) drawCharts(d);\n  }\n\n  function drawCharts(d) {\n    bar('c_managers', d.managers, 'Менеджер', 'SP');\n    bar('c_managers_money', d.managersMoney, 'Менеджер', '₽');\n    bar('c_contractors', d.contractorsMoney, 'Подрядчик', '₽');\n    bar('c_products', d.products, 'Продукт', 'Задачи');\n    bar('c_langs', d.languages, 'Язык', 'Упоминаний');\n    columns('c_months', d.byMonth, 'Месяц');\n    columns('c_years', d.byYear, 'Год');\n  }\n\n  function bar(id, rows, a, b) {\n    const dt = new google.visualization.DataTable();\n    dt.addColumn('string', a); dt.addColumn('number', b);\n    const data = rows.filter(r => r[1] > 0).sort((x, y) => y[1] - x[1]).slice(0, 10);\n    if (!data.length) { document.getElementById(id).innerHTML = '<div class=\"small muted\">Нет данных за период</div>'; return; }\n    data.forEach(r => dt.addRow(r));\n    new google.visualization.BarChart(document.getElementById(id)).draw(dt, {\n      legend: 'none', colors: [COLORS[0]], chartArea: { width: '58%', height: '82%' },\n      hAxis: { textStyle: { fontSize: 10 } }, vAxis: { textStyle: { fontSize: 10 } }\n    });\n  }\n\n  function columns(id, rows, a) {\n    const dt = new google.visualization.DataTable();\n    dt.addColumn('string', a); dt.addColumn('number', 'Задачи'); dt.addColumn('number', 'SP');\n    rows.forEach(r => dt.addRow(r));\n    new google.visualization.ColumnChart(document.getElementById(id)).draw(dt, {\n      legend: { position: 'top' }, colors: COLORS, chartArea: { width: '84%', height: '70%' },\n      hAxis: { textStyle: { fontSize: 9 } }, vAxis: { textStyle: { fontSize: 10 } }\n    });\n  }\n</script>\n</body>\n</html>\n",
  "ManagerReportSidebar": "<!DOCTYPE html>\n<html>\n<head>\n<base target=\"_top\">\n<?!= include('Общее') ?>\n</head>\n<body>\n  <h2>Отчёт по менеджеру</h2>\n  <div class=\"stack\">\n    <div class=\"two\">\n      <label class=\"field\">Менеджер<select id=\"manager\"></select></label>\n      <label class=\"field\">Месяц<select id=\"month\"></select></label>\n    </div>\n    <button class=\"btn primary\" id=\"runBtn\" type=\"button\">Сформировать отчёт</button>\n  </div>\n  <div id=\"results\"></div>\n  <div class=\"two\" id=\"actions\" hidden style=\"margin-top:8px\">\n    <button class=\"btn\" id=\"copyBtn\" type=\"button\">Скопировать текстом</button>\n    <button class=\"btn\" id=\"exportBtn\" type=\"button\">📥 Скачать Excel</button>\n  </div>\n  <div id=\"msg\" class=\"msg\" style=\"margin-top:6px\"></div>\n\n  <hr class=\"sep\">\n  <h2>Отчёт для трекера</h2>\n  <div class=\"stack\">\n    <label class=\"field\">Тикет<select id=\"ticket\"></select></label>\n    <button class=\"btn primary\" id=\"trackerBtn\" type=\"button\">Сформировать</button>\n  </div>\n  <pre class=\"out\" id=\"trackerOut\" hidden></pre>\n  <button class=\"btn full\" id=\"trackerCopy\" type=\"button\" hidden>Скопировать</button>\n  <div id=\"trackerMsg\" class=\"msg\"></div>\n\n<script>\n  let report = null, trackerText = '';\n\n  Promise.all([run('getManagersList'), run('getMonthsList'), run('getTicketsForReport'), run('getAddTaskFormLists')])\n    .then(([managers, months, tickets, lists]) => {\n      fillSelect($('#manager'), managers, 'Выберите…', lists.currentManager);\n      $('#month').innerHTML = '<option value=\"\">Все месяцы</option>' + months.map(m => `<option value=\"${m.value}\">${esc(m.label)}</option>`).join('');\n      fillSelect($('#ticket'), tickets, 'Выберите тикет…');\n    }).catch(e => setMsg($('#msg'), 'Не получилось загрузить списки: ' + e.message, 'err'));\n\n  $('#runBtn').addEventListener('click', () => {\n    const manager = $('#manager').value;\n    if (!manager) return setMsg($('#msg'), 'Выберите менеджера', 'err');\n    $('#results').innerHTML = '<p class=\"muted\">Считаю…</p>';\n    $('#actions').hidden = true;\n    setMsg($('#msg'), '');\n    run('getManagerReport', manager, $('#month').value).then(render)\n      .catch(e => { $('#results').innerHTML = ''; setMsg($('#msg'), 'Не получилось: ' + e.message, 'err'); });\n  });\n\n  function render(d) {\n    report = d;\n    if (!d.groups.length) { $('#results').innerHTML = '<p class=\"muted\">Ничего не найдено.</p>'; return; }\n    let html = `<div class=\"tiles\">\n      <div class=\"tile\"><div class=\"num\">${d.taskCount}</div><div class=\"lbl\">тикетов</div></div>\n      <div class=\"tile\"><div class=\"num\">${d.lineCount}</div><div class=\"lbl\">задач</div></div>\n      <div class=\"tile\"><div class=\"num\">${d.grandTotal}</div><div class=\"lbl\">SP</div></div>\n      <div class=\"tile\"><div class=\"num\">${money(d.grandMoney)}</div><div class=\"lbl\">₽ с НДС</div></div></div>`;\n    d.groups.forEach(g => {\n      html += `<div class=\"group\"><h4>${esc(g.ticket || '(без тикета)')}</h4>` +\n        g.lines.map(l => `<div class=\"line\"><span>${l.link ? `<a href=\"${esc(l.link)}\" target=\"_blank\">${esc(l.subject)}</a>` : esc(l.subject)}` +\n          l.link2.split('\\n').filter(Boolean).map((u, i) => ` <a href=\"${esc(u)}\" target=\"_blank\" class=\"small\">(ссылка ${i + 2})</a>`).join('') + `</span><span class=\"v\">${l.sp} SP</span></div>`).join('') +\n        `<div class=\"line\"><span class=\"muted\">Итого по тикету</span><span class=\"v\">${g.ticketTotal} SP</span></div></div>`;\n    });\n    html += `<div class=\"total\">Общий итог по ${esc(d.manager)}${d.monthLabel ? ' за ' + esc(d.monthLabel) : ''}: ${d.grandTotal} SP</div>`;\n    $('#results').innerHTML = html;\n    $('#actions').hidden = false;\n  }\n\n  $('#copyBtn').addEventListener('click', () => {\n    if (!report) return;\n    let html = '';\n    report.groups.forEach(g => g.lines.forEach(l => {\n      html += `<div>${l.link ? `<a href=\"${esc(l.link)}\">${esc(l.subject)}</a>` : esc(l.subject)} — ${l.sp} SP</div>`;\n    }));\n    html += `<div><b>Общий итог по ${esc(report.manager)}: ${report.grandTotal} SP</b></div>`;\n    copyHtml(html);\n    setMsg($('#msg'), 'Скопировано — ссылки кликабельны при вставке в Doc, почту, Band', 'ok');\n  });\n\n  $('#exportBtn').addEventListener('click', () => {\n    setMsg($('#msg'), 'Готовлю файл…');\n    run('exportReportToDoc', $('#manager').value, $('#month').value).then(url => {\n      setMsg($('#msg'), 'Готово — файл скачивается', 'ok');\n      window.open(url, '_blank');\n    }).catch(e => setMsg($('#msg'), 'Не получилось: ' + e.message, 'err'));\n  });\n\n  $('#trackerBtn').addEventListener('click', () => {\n    const ticket = $('#ticket').value;\n    if (!ticket) return setMsg($('#trackerMsg'), 'Выберите тикет', 'err');\n    setMsg($('#trackerMsg'), 'Считаю…');\n    run('getTrackerReportText', ticket).then(r => {\n      trackerText = r.text;\n      $('#trackerOut').hidden = !r.count;\n      $('#trackerCopy').hidden = !r.count;\n      $('#trackerOut').textContent = r.text;\n      setMsg($('#trackerMsg'), r.count ? '' : 'По этому тикету задач нет');\n    }).catch(e => setMsg($('#trackerMsg'), 'Не получилось: ' + e.message, 'err'));\n  });\n  $('#trackerCopy').addEventListener('click', () => copyText(trackerText).then(() => setMsg($('#trackerMsg'), 'Скопировано', 'ok')));\n</script>\n</body>\n</html>\n",
  "CustomReportSidebar": "<!DOCTYPE html>\n<html>\n<head>\n<base target=\"_top\">\n<?!= include('Общее') ?>\n</head>\n<body>\n  <h2>Кастомный отчёт</h2>\n  <div class=\"stack\">\n    <div class=\"two\">\n      <label class=\"field\">Год<select id=\"year\"></select></label>\n      <label class=\"field\">Месяц<select id=\"month\">\n        <option value=\"\">Весь год</option><option value=\"1\">Январь</option><option value=\"2\">Февраль</option><option value=\"3\">Март</option>\n        <option value=\"4\">Апрель</option><option value=\"5\">Май</option><option value=\"6\">Июнь</option><option value=\"7\">Июль</option>\n        <option value=\"8\">Август</option><option value=\"9\">Сентябрь</option><option value=\"10\">Октябрь</option>\n        <option value=\"11\">Ноябрь</option><option value=\"12\">Декабрь</option></select></label>\n    </div>\n    <button class=\"btn primary\" id=\"runBtn\" type=\"button\">Сформировать</button>\n  </div>\n  <div id=\"results\"></div>\n  <div class=\"two\" id=\"actions\" hidden style=\"margin-top:8px\">\n    <button class=\"btn\" id=\"excelBtn\" type=\"button\">📥 Скачать Excel</button>\n    <button class=\"btn\" id=\"docBtn\" type=\"button\">📄 Открыть в Google Doc</button>\n  </div>\n  <div id=\"msg\" class=\"msg\" style=\"margin-top:6px\"></div>\n\n<script>\n  run('getDashboardYears').then(years => fillSelect($('#year'), years.map(String), 'Всё время'))\n    .catch(e => setMsg($('#msg'), 'Не получилось загрузить годы: ' + e.message, 'err'));\n\n  const period = () => [$('#year').value, $('#year').value ? $('#month').value : ''];\n\n  function table(rows, fmt) {\n    if (!rows.length) return '<p class=\"small muted\">Нет данных</p>';\n    return '<table class=\"kv\">' + rows.map(r => `<tr><td>${esc(r[0])}</td><td>${esc(fmt ? fmt(r) : r[1])}</td></tr>`).join('') + '</table>';\n  }\n\n  $('#runBtn').addEventListener('click', () => {\n    $('#results').innerHTML = '<p class=\"muted\">Считаю…</p>';\n    $('#actions').hidden = true;\n    setMsg($('#msg'), '');\n    run('getCustomReport', ...period()).then(d => {\n      let html = `<p class=\"small muted\">Период: <b>${esc(d.periodLabel)}</b></p><div class=\"tiles\">\n        <div class=\"tile\"><div class=\"num\">${d.totalTasks}</div><div class=\"lbl\">задач</div></div>\n        <div class=\"tile\"><div class=\"num\">${d.totalSp}</div><div class=\"lbl\">SP</div></div>\n        <div class=\"tile\"><div class=\"num\">${d.avgSp}</div><div class=\"lbl\">SP на задачу</div></div>\n        <div class=\"tile\"><div class=\"num\">${money(d.totalMoney)}</div><div class=\"lbl\">₽ с НДС</div></div></div>`;\n      html += '<h3>Продукты по числу задач</h3>' + table(d.topProducts);\n      html += '<h3>Языки по частоте</h3>' + table(d.languages);\n      html += '<h3>Статус отдачи заказчику</h3>' + table(d.deliveryStatus, r => r[1] + ' (' + r[2] + '%)');\n      html += '<h3>Топ-5 менеджеров по SP</h3>' + table(d.topManagers, r => r[2] + ' SP · ' + r[1] + ' задач');\n      html += '<h3>Топ-5 подрядчиков</h3>' + table(d.topContractors);\n      html += '<h3>Топ-10 самых трудоёмких задач</h3>' + (d.topTasksBySP.length ? d.topTasksBySP.map(t =>\n        `<div class=\"line\"><span>${t.link ? `<a href=\"${esc(t.link)}\" target=\"_blank\">${esc(t.subject)}</a>` : esc(t.subject)}</span><span class=\"v\">${t.sp} SP</span></div>`).join('')\n        : '<p class=\"small muted\">Нет задач с SP</p>');\n      $('#results').innerHTML = html;\n      $('#actions').hidden = false;\n    }).catch(e => { $('#results').innerHTML = ''; setMsg($('#msg'), 'Не получилось: ' + e.message, 'err'); });\n  });\n\n  $('#excelBtn').addEventListener('click', () => {\n    setMsg($('#msg'), 'Готовлю файл…');\n    run('exportCustomReportToExcel', ...period()).then(url => { setMsg($('#msg'), 'Готово — файл скачивается', 'ok'); window.open(url, '_blank'); })\n      .catch(e => setMsg($('#msg'), 'Не получилось: ' + e.message, 'err'));\n  });\n  $('#docBtn').addEventListener('click', () => {\n    setMsg($('#msg'), 'Создаю документ…');\n    run('exportCustomReportToDoc', ...period()).then(url => {\n      $('#msg').className = 'msg ok';\n      $('#msg').innerHTML = `<a href=\"${esc(url)}\" target=\"_blank\">Открыть документ →</a>`;\n    }).catch(e => setMsg($('#msg'), 'Не получилось: ' + e.message, 'err'));\n  });\n</script>\n</body>\n</html>\n",
  "AddTranslatorTaskDialog": "<!DOCTYPE html>\n<html>\n<head>\n<base target=\"_top\">\n<?!= include('Общее') ?>\n</head>\n<body class=\"has-foot\">\n  <div class=\"stack\">\n    <div class=\"two\">\n      <label class=\"field\">Дата<input type=\"date\" id=\"date\"></label>\n      <label class=\"field\"><span>Сторона <span class=\"req\">*</span></span><select id=\"side\"></select></label>\n    </div>\n    <label class=\"field\">Раздел<select id=\"razdel\"><option value=\"\">Сначала выберите сторону</option></select></label>\n    <label class=\"field\"><span>Задача <span class=\"req\">*</span></span><input type=\"text\" id=\"task\" placeholder=\"Например: перевести баннер главной страницы\"></label>\n    <div class=\"two\">\n      <label class=\"field\">Заказчик<input type=\"text\" id=\"customer\" placeholder=\"@nick\"></label>\n      <label class=\"field\">Ссылка на Band<input type=\"url\" id=\"link\" placeholder=\"https://band.wb.ru/…\"></label>\n    </div>\n    <div class=\"field\">Переводчик<div id=\"translators\"></div></div>\n    <div class=\"field\">Редактор<div id=\"editors\"></div></div>\n    <div class=\"two\">\n      <label class=\"field\">Готовность<select id=\"readiness\">\n        <option>Не начато</option><option>В работе</option><option>Готово</option><option>На проверке</option></select></label>\n      <label class=\"field\">Комментарий<input type=\"text\" id=\"comment\"></label>\n    </div>\n    <div id=\"msg\" class=\"msg\"></div>\n  </div>\n  <div class=\"foot\"><button class=\"btn primary grow\" id=\"submitBtn\" type=\"button\">Добавить задачу</button></div>\n\n<script>\n  let CATEGORIES = {};\n  $('#date').value = todayIso();\n\n  run('getContentCategories').then(d => {\n    CATEGORIES = d;\n    fillSelect($('#side'), Object.keys(d), 'Выберите…');\n  }).catch(e => setMsg($('#msg'), 'Не получилось загрузить разделы: ' + e.message, 'err'));\n\n  run('getTranslatorNames').then(names => {\n    renderChips($('#translators'), [{ items: names }], []);\n    renderChips($('#editors'), [{ items: names }], []);\n  }).catch(e => setMsg($('#msg'), 'Не получилось загрузить людей: ' + e.message, 'err'));\n\n  $('#side').addEventListener('change', () => {\n    const items = CATEGORIES[$('#side').value] || [];\n    fillSelect($('#razdel'), items, items.length ? 'Выберите…' : '—', items.length === 1 ? items[0] : '');\n  });\n\n  $('#submitBtn').addEventListener('click', () => {\n    if (!$('#side').value) return setMsg($('#msg'), 'Выберите сторону', 'err');\n    if (!$('#task').value.trim()) return setMsg($('#msg'), 'Опишите задачу', 'err');\n    const task = {\n      date: $('#date').value, side: $('#side').value, razdel: $('#razdel').value, task: $('#task').value,\n      customer: $('#customer').value, link: $('#link').value,\n      translator: chipValues($('#translators')).join(', '), editor: chipValues($('#editors')).join(', '),\n      readiness: $('#readiness').value, comment: $('#comment').value\n    };\n    const btn = $('#submitBtn');\n    btn.disabled = true; btn.textContent = 'Добавляю…';\n    run('submitNewTranslatorTask', task).then(() => {\n      setMsg($('#msg'), 'Добавлено! Закрываю…', 'ok');\n      setTimeout(() => google.script.host.close(), 900);\n    }).catch(e => {\n      setMsg($('#msg'), 'Не добавилось: ' + e.message, 'err');\n      btn.disabled = false; btn.textContent = 'Добавить задачу';\n    });\n  });\n</script>\n</body>\n</html>\n",
  "SearchEditTranslatorSidebar": "<!DOCTYPE html>\n<html>\n<head>\n<base target=\"_top\">\n<?!= include('Общее') ?>\n</head>\n<body class=\"has-foot\">\n  <div class=\"stack\">\n    <label class=\"field\">Найти<input type=\"text\" id=\"q\" placeholder=\"Задача, заказчик, раздел или имя\"></label>\n    <div class=\"small muted\" id=\"qHint\">Ваши задачи:</div>\n    <div class=\"results\" id=\"results\"></div>\n\n    <div id=\"editor\" class=\"stack\" hidden>\n      <div class=\"card\"><div class=\"t\" id=\"cardTitle\"></div><div class=\"small muted\" id=\"cardMeta\"></div></div>\n      <div class=\"field\">Готовность<div class=\"seg\" id=\"readySeg\"></div></div>\n      <label class=\"field\">Задача<input type=\"text\" id=\"f_task\"></label>\n      <label class=\"field\">Ссылка на Band<input type=\"url\" id=\"f_link\"></label>\n      <div class=\"field\">Переводчик<div id=\"f_translators\"></div></div>\n      <div class=\"field\">Редактор<div id=\"f_editors\"></div></div>\n      <details class=\"more\">\n        <summary>Дата, сторона, раздел, заказчик, комментарий</summary>\n        <div class=\"stack\">\n          <div class=\"two\">\n            <label class=\"field\">Дата<input type=\"date\" id=\"f_date\"></label>\n            <label class=\"field\">Сторона<select id=\"f_side\"></select></label>\n          </div>\n          <label class=\"field\">Раздел<select id=\"f_razdel\"></select></label>\n          <label class=\"field\">Заказчик<input type=\"text\" id=\"f_customer\"></label>\n          <label class=\"field\">Комментарий<textarea id=\"f_comment\"></textarea></label>\n        </div>\n      </details>\n    </div>\n    <div id=\"msg\" class=\"msg\"></div>\n  </div>\n\n  <div class=\"foot col\" id=\"foot\" hidden>\n    <button class=\"btn primary full\" id=\"saveBtn\" type=\"button\">Сохранить</button>\n    <button class=\"link-danger\" id=\"deleteBtn\" type=\"button\">Удалить задачу</button>\n  </div>\n\n<script>\n  const READY = ['Не начато', 'В работе', 'На проверке', 'Готово'];\n  const READY_COLORS = { 'Не начато': '#EFEFEF', 'В работе': '#D6E4F0', 'Готово': '#D9EAD3', 'На проверке': '#FFF2CC' };\n  let CATEGORIES = {}, PEOPLE = [], current = null, readiness = '';\n\n  Promise.all([run('getContentCategories'), run('getTranslatorNames')]).then(([cats, names]) => {\n    CATEGORIES = cats; PEOPLE = names;\n    fillSelect($('#f_side'), Object.keys(cats), '—');\n    search('');\n  }).catch(e => setMsg($('#msg'), 'Не получилось загрузить списки: ' + e.message, 'err'));\n\n  let timer = null, seq = 0;\n  $('#q').addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(() => search($('#q').value), 350); });\n\n  function search(q) {\n    const my = ++seq;\n    $('#qHint').textContent = q.trim() ? 'Найдено:' : 'Ваши задачи (или последние, если вашей почты нет в «Списках»):';\n    run('searchTranslatorTasks', q).then(list => {\n      if (my !== seq) return;\n      $('#results').innerHTML = list.length ? list.map(r =>\n        `<button type=\"button\" data-row=\"${r.row}\">${esc(r.title)} <span class=\"d\">· ${esc(r.side)} · ${esc(r.date)}${r.readiness ? ' · ' + esc(r.readiness) : ''}</span></button>`\n      ).join('') : '<div class=\"small muted\" style=\"padding:8px 10px\">Ничего не найдено</div>';\n    }).catch(e => setMsg($('#msg'), 'Поиск не сработал: ' + e.message, 'err'));\n  }\n\n  $('#results').addEventListener('click', e => {\n    const b = e.target.closest('button[data-row]');\n    if (!b) return;\n    $$('#results button').forEach(x => x.classList.toggle('on', x === b));\n    run('getTranslatorTaskForEdit', Number(b.dataset.row)).then(fill).catch(err => setMsg($('#msg'), 'Не открылась: ' + err.message, 'err'));\n  });\n\n  function renderReady() {\n    $('#readySeg').innerHTML = READY.map(s =>\n      `<button type=\"button\" data-s=\"${esc(s)}\" class=\"${s === readiness ? 'on' : ''}\" style=\"${s === readiness ? 'background:' + READY_COLORS[s] : ''}\">${esc(s)}</button>`).join('');\n  }\n  $('#readySeg').addEventListener('click', e => {\n    const b = e.target.closest('button[data-s]');\n    if (b) { readiness = b.dataset.s; renderReady(); }\n  });\n\n  function peopleChips(el, chosen) {\n    renderChips(el, [{ items: PEOPLE }, { title: 'Другие', items: chosen.filter(n => !PEOPLE.includes(n)) }], chosen);\n  }\n\n  function updateRazdel(selected) {\n    const items = CATEGORIES[$('#f_side').value] || [];\n    fillSelect($('#f_razdel'), items.includes(selected) || !selected ? items : items.concat([selected]), '—', selected);\n  }\n  $('#f_side').addEventListener('change', () => updateRazdel(''));\n\n  function fill(t) {\n    current = t;\n    readiness = t.readiness || 'Не начато';\n    $('#cardTitle').textContent = t.task || '(задача не заполнена)';\n    $('#cardMeta').textContent = [t.side, t.razdel, t.customer].filter(Boolean).join(' · ');\n    renderReady();\n    $('#f_task').value = t.task; $('#f_link').value = t.link; $('#f_date').value = t.date;\n    $('#f_side').value = t.side; updateRazdel(t.razdel);\n    $('#f_customer').value = t.customer; $('#f_comment').value = t.comment;\n    peopleChips($('#f_translators'), t.translatorNames);\n    peopleChips($('#f_editors'), t.editorNames);\n    $('#editor').hidden = false; $('#foot').hidden = false;\n    setMsg($('#msg'), '');\n  }\n\n  $('#saveBtn').addEventListener('click', () => {\n    if (!current) return;\n    const task = {\n      row: current.row, origKey: current.origKey, readiness: readiness,\n      task: $('#f_task').value, link: $('#f_link').value, date: $('#f_date').value, side: $('#f_side').value,\n      razdel: $('#f_razdel').value, customer: $('#f_customer').value, comment: $('#f_comment').value,\n      translator: chipValues($('#f_translators')).join(', '), editor: chipValues($('#f_editors')).join(', ')\n    };\n    const btn = $('#saveBtn');\n    btn.disabled = true;\n    setMsg($('#msg'), 'Сохраняю…');\n    run('saveTranslatorTaskEdits', task)\n      .then(r => run('getTranslatorTaskForEdit', r.row)).then(fill)\n      .then(() => { setMsg($('#msg'), 'Сохранено ✓', 'ok'); search($('#q').value); })\n      .catch(e => setMsg($('#msg'), 'Не сохранилось: ' + e.message, 'err'))\n      .finally(() => { btn.disabled = false; });\n  });\n\n  $('#deleteBtn').addEventListener('click', () => {\n    if (!current) return;\n    if (!confirm('Удалить задачу «' + (current.task || '') + '»? Отменить будет нельзя.')) return;\n    run('deleteTranslatorTask', current.row, current.origKey).then(() => {\n      current = null;\n      $('#editor').hidden = true; $('#foot').hidden = true;\n      setMsg($('#msg'), 'Удалено.', 'ok');\n      search($('#q').value);\n    }).catch(e => setMsg($('#msg'), 'Не удалилось: ' + e.message, 'err'));\n  });\n</script>\n</body>\n</html>\n",
  "TranslatorReportSidebar": "<!DOCTYPE html>\n<html>\n<head>\n<base target=\"_top\">\n<?!= include('Общее') ?>\n</head>\n<body>\n  <h2>Отчёт по переводчику</h2>\n  <div class=\"stack\">\n    <div class=\"two\">\n      <label class=\"field\">Переводчик<select id=\"translator\"></select></label>\n      <label class=\"field\">Месяц<select id=\"month\"></select></label>\n    </div>\n    <div class=\"two\">\n      <button class=\"btn primary\" id=\"runBtn\" type=\"button\">Сформировать отчёт</button>\n      <button class=\"btn\" id=\"trackerBtn\" type=\"button\">Список для трекера</button>\n    </div>\n  </div>\n  <div id=\"results\"></div>\n  <pre class=\"out\" id=\"trackerOut\" hidden></pre>\n  <div class=\"two\" style=\"margin-top:8px\">\n    <button class=\"btn\" id=\"exportBtn\" type=\"button\" hidden>📥 Скачать Excel</button>\n    <button class=\"btn\" id=\"trackerCopy\" type=\"button\" hidden>Скопировать список</button>\n  </div>\n  <div id=\"msg\" class=\"msg\" style=\"margin-top:6px\"></div>\n\n<script>\n  let trackerText = '';\n  const badge = r => r === 'Готово' ? 'b-done' : r === 'В работе' ? 'b-progress' : r === 'На проверке' ? 'b-review' : '';\n\n  Promise.all([run('getTranslatorNames'), run('getMonthsList')]).then(([names, months]) => {\n    fillSelect($('#translator'), names, 'Выберите…');\n    $('#month').innerHTML = '<option value=\"\">Все месяцы</option>' + months.map(m => `<option value=\"${m.value}\">${esc(m.label)}</option>`).join('');\n  }).catch(e => setMsg($('#msg'), 'Не получилось загрузить списки: ' + e.message, 'err'));\n\n  function need() {\n    if ($('#translator').value) return true;\n    setMsg($('#msg'), 'Выберите переводчика', 'err');\n    return false;\n  }\n\n  $('#runBtn').addEventListener('click', () => {\n    if (!need()) return;\n    $('#trackerOut').hidden = true; $('#trackerCopy').hidden = true;\n    $('#results').innerHTML = '<p class=\"muted\">Считаю…</p>';\n    setMsg($('#msg'), '');\n    run('getTranslatorReport', $('#translator').value, $('#month').value).then(d => {\n      if (!d.total) { $('#results').innerHTML = '<p class=\"muted\">Ничего не найдено.</p>'; $('#exportBtn').hidden = true; return; }\n      let html = `<div class=\"tiles\">\n        <div class=\"tile\"><div class=\"num\">${d.total}</div><div class=\"lbl\">всего</div></div>\n        <div class=\"tile\"><div class=\"num\">${d.done}</div><div class=\"lbl\">готово</div></div>\n        <div class=\"tile\"><div class=\"num\">${d.inProgress}</div><div class=\"lbl\">в работе</div></div></div>`;\n      Object.keys(d.bySide).forEach(side => {\n        const tasks = d.bySide[side];\n        html += `<div class=\"group\"><h4>${esc(side)} (${tasks.length})</h4>` + tasks.map(t =>\n          `<div class=\"line\"><span>${t.link ? `<a href=\"${esc(t.link)}\" target=\"_blank\">${esc(t.task)}</a>` : esc(t.task)}</span>` +\n          `<span class=\"badge ${badge(t.readiness)}\">${esc(t.readiness)}</span></div>`).join('') + '</div>';\n      });\n      $('#results').innerHTML = html;\n      $('#exportBtn').hidden = false;\n    }).catch(e => { $('#results').innerHTML = ''; setMsg($('#msg'), 'Не получилось: ' + e.message, 'err'); });\n  });\n\n  $('#exportBtn').addEventListener('click', () => {\n    setMsg($('#msg'), 'Готовлю файл…');\n    run('exportTranslatorReportToExcel', $('#translator').value, $('#month').value)\n      .then(url => { setMsg($('#msg'), 'Готово — файл скачивается', 'ok'); window.open(url, '_blank'); })\n      .catch(e => setMsg($('#msg'), 'Не получилось: ' + e.message, 'err'));\n  });\n\n  $('#trackerBtn').addEventListener('click', () => {\n    if (!need()) return;\n    setMsg($('#msg'), 'Считаю…');\n    run('getTranslatorTrackerReportText', $('#translator').value).then(r => {\n      trackerText = r.text;\n      $('#trackerOut').textContent = r.text;\n      $('#trackerOut').hidden = !r.count; $('#trackerCopy').hidden = !r.count;\n      setMsg($('#msg'), r.count ? '' : 'Задач не найдено');\n    }).catch(e => setMsg($('#msg'), 'Не получилось: ' + e.message, 'err'));\n  });\n  $('#trackerCopy').addEventListener('click', () => copyText(trackerText).then(() => setMsg($('#msg'), 'Скопировано', 'ok')));\n</script>\n</body>\n</html>\n",
  "MigrateDialog": "<!DOCTYPE html>\n<html>\n<head>\n<base target=\"_top\">\n<?!= include('Общее') ?>\n<style>\n  .drop { display: block; border: 2px dashed var(--field); border-radius: 10px; padding: 22px 14px; text-align: center;\n    cursor: pointer; font-size: 13px; font-weight: 400; color: var(--soft); }\n  .drop:hover, .drop.over { border-color: var(--accent); background: var(--accent-soft); }\n  .drop b { color: var(--ink); }\n  .drop input { display: none; }\n  .or { text-align: center; font-size: 12px; color: var(--soft); }\n</style>\n</head>\n<body class=\"has-foot\">\n  <div class=\"stack\">\n    <label class=\"drop\" id=\"drop\">\n      <input type=\"file\" id=\"file\" accept=\".xlsx\">\n      <b id=\"fileName\">Выберите файл Excel</b><br>\n      или перетащите его сюда\n      <div class=\"small\" style=\"margin-top:6px\">Старая таблица (Файл → Скачать → Microsoft Excel) или файл с колонками новой таблицы</div>\n    </label>\n    <div class=\"or\">— или —</div>\n    <label class=\"field\">Ссылка на старую таблицу\n      <input type=\"url\" id=\"link\" placeholder=\"https://docs.google.com/spreadsheets/d/…\"></label>\n    <div id=\"plan\" class=\"preview\" hidden></div>\n    <div id=\"msg\" class=\"msg\"></div>\n  </div>\n  <div class=\"foot\">\n    <button class=\"btn\" id=\"checkBtn\" type=\"button\">Проверить</button>\n    <button class=\"btn primary grow\" id=\"goBtn\" type=\"button\" disabled>Перенести</button>\n  </div>\n\n<script src=\"https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js\"></script>\n<script>\n// xlsx-parse:start\n  /** Лист «Localization Misc» из файла Excel → { values, links, merges } для migrationPreviewFromData. */\n  function oldSheetFromWorkbook(XLSX, wb) {\n    const COLS = 19, BAND = 5;\n    const name = wb.SheetNames.find(n => n.trim() === 'Localization Misc');\n    if (!name) throw new Error('В файле нет листа «Localization Misc». Это точно старая таблица?');\n    const ws = wb.Sheets[name];\n    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');\n    const values = [], links = [];\n    let last = 0;\n    for (let r = 0; r <= range.e.r; r++) {\n      const row = [];\n      let link = '';\n      for (let c = 0; c < COLS; c++) {\n        const cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];\n        row.push(xlsxCellValue(XLSX, cell));\n        if (c === BAND && cell && cell.l && /^https?:\\/\\//i.test(cell.l.Target || '')) link = cell.l.Target;\n      }\n      if (link || row.some(v => v !== '')) last = r + 1;\n      values.push(row);\n      links.push(link);\n    }\n    values.length = last;\n    links.length = last;\n    const merges = (ws['!merges'] || [])\n      .filter(m => m.s.c < COLS && m.s.r < last)\n      .map(m => [m.s.r + 1, m.s.c + 1, m.e.r - m.s.r + 1, m.e.c - m.s.c + 1]);\n    return { values: values, links: links, merges: merges };\n  }\n\n  /** Как getValues(): число, строка, логическое; дата — { d: 'yyyy-MM-dd' } (Date через google.script.run не пройдёт). */\n  function xlsxCellValue(XLSX, cell) {\n    if (!cell || cell.v == null || cell.t === 'e' || cell.t === 'z') return '';\n    if (cell.t === 'n' && cell.z && XLSX.SSF.is_date(cell.z)) {\n      const p = XLSX.SSF.parse_date_code(cell.v);\n      const pad = n => String(n).padStart(2, '0');\n      return { d: p.y + '-' + pad(p.m) + '-' + pad(p.d) };\n    }\n    return typeof cell.v === 'string' ? cell.v.replace(/\\r\\n?/g, '\\n') : cell.v;\n  }\n  /**\n   * Файл в формате новой таблицы (шапка «№ задачи», «Тикет», «Тема письма»…), например «Перенос — задачи из старой таблицы.xlsx».\n   * Возвращает { kind: 'tasks', sheets: [{ name, archive, rows: [{ values, link, link2 }] }] } или null.\n   */\n  function tasksFileFromWorkbook(XLSX, wb) {\n    const COLS = 17, SUBJ = 2, EST = 10, COMMENT = 16;\n    const sheets = [];\n    wb.SheetNames.forEach(name => {\n      const ws = wb.Sheets[name];\n      if (!ws || !ws['!ref']) return;\n      const at = (r, c) => ws[XLSX.utils.encode_cell({ r: r, c: c })];\n      const head = c => String((at(0, c) || {}).v || '').trim();\n      if (head(0) !== '№ задачи' || head(SUBJ) !== 'Тема письма') return;\n      const range = XLSX.utils.decode_range(ws['!ref']);\n      const rows = [];\n      for (let r = 1; r <= range.e.r; r++) {\n        const values = [];\n        for (let c = 0; c < COLS; c++) values.push(xlsxCellValue(XLSX, at(r, c)));\n        if (!values.some(v => v !== '')) continue;\n        const target = c => { const cell = at(r, c); return cell && cell.l && /^https?:\\/\\//i.test(cell.l.Target || '') ? cell.l.Target : ''; };\n        if (target(EST)) values[EST] = target(EST); // «Ссылка на смету …» → сам адрес\n        // Доп. ссылки, сложенные в комментарий, — обратно в ссылки после темы\n        const extra = [];\n        values[COMMENT] = String(values[COMMENT] || '').split('\\n').filter(line => {\n          const m = line.match(/^\\s*(?:Доп\\. ссылк[аи] на Band|Ещё ссылки)\\s*:\\s*(.*)$/i);\n          if (m) extra.push(...(m[1].match(/https?:\\/\\/\\S+/g) || []));\n          return !m;\n        }).join('\\n');\n        rows.push({ values: values, link: target(SUBJ), link2: extra.join('\\n') });\n      }\n      sheets.push({ name: name, archive: /^(данные до|архив)/i.test(name.trim()), rows: rows });\n    });\n    return sheets.length ? { kind: 'tasks', sheets: sheets } : null;\n  }\n// xlsx-parse:end\n\n  let source = null;   // данные из файла\n  let mode = '';       // 'file' | 'link'\n  let fileError = '';  // почему файл не прочитался — показываем и при «Проверить»\n\n  function resetPlan() {\n    $('#plan').hidden = true;\n    $('#goBtn').disabled = true;\n  }\n\n  function readFile(file) {\n    resetPlan();\n    source = null;\n    fileError = '';\n    if (!file) return;\n    if (!/\\.xlsx$/i.test(file.name)) return setMsg($('#msg'), 'Нужен файл .xlsx', 'err');\n    $('#fileName').textContent = file.name;\n    $('#link').value = '';\n    setMsg($('#msg'), 'Читаю файл…');\n    if (typeof XLSX === 'undefined') return setMsg($('#msg'), 'Не загрузилась читалка Excel. Проверьте интернет или вставьте ссылку.', 'err');\n    file.arrayBuffer().then(buf => {\n      const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellNF: true, cellDates: false });\n      source = tasksFileFromWorkbook(XLSX, wb);\n      if (!source) {\n        if (!wb.SheetNames.some(n => n.trim() === 'Localization Misc')) {\n          throw new Error('в нём нет ни листа «Localization Misc» (старая таблица), ни листа с колонками новой таблицы («№ задачи», «Тикет», «Тема письма»…)');\n        }\n        source = oldSheetFromWorkbook(XLSX, wb);\n      }\n      mode = 'file';\n      check();\n    }).catch(e => {\n      fileError = 'Не получилось прочитать файл: ' + e.message;\n      setMsg($('#msg'), fileError, 'err');\n    });\n  }\n\n  function check() {\n    resetPlan();\n    if (!source && $('#link').value.trim()) mode = 'link';\n    if (!mode || (mode === 'file' && !source)) return setMsg($('#msg'), fileError || 'Выберите файл или вставьте ссылку', 'err');\n    setMsg($('#msg'), 'Считаю, что перенести…');\n    if (source && source.kind === 'tasks') return checkTasksFile();\n    const call = mode === 'file' ? run('migrationPreviewFromData', source) : run('migrationPreviewFromLink', $('#link').value);\n    call.then(p => {\n      $('#plan').hidden = false;\n      if (!p.count) {\n        $('#plan').innerHTML = 'Переносить нечего: ' + p.skippedExisting + ' задач уже есть в новой таблице.';\n        return setMsg($('#msg'), '');\n      }\n      $('#plan').innerHTML =\n        '<b>Будет добавлено задач: ' + p.count + '</b> (с ' + esc(p.from) + ' по ' + esc(p.to) + ')<br>' +\n        '<span class=\"small\">Уже есть в новой таблице — пропущу: ' + p.skippedExisting +\n        '. Пустые строки — пропущу: ' + p.skippedEmpty + '.</span><br>' +\n        '<span class=\"small\">После переноса весь лист отсортируется по дате: свежие сверху.</span>';\n      $('#goBtn').disabled = false;\n      setMsg($('#msg'), '');\n    }).catch(e => setMsg($('#msg'), e.message, 'err'));\n  }\n\n  /** Файл в формате новой таблицы: что уйдёт в задачи, что — в архив. */\n  function checkTasksFile() {\n    run('migrationPreviewFromTasksFile', source).then(p => {\n      $('#plan').hidden = false;\n      const total = p.main + p.archives.reduce((n, a) => n + a.count, 0);\n      const lines = ['<b>В «📌 Задачи (менеджеры)»: ' + p.main + '</b>' + (p.mainSkipped ? ' <span class=\"small\">(уже есть — пропущу ' + p.mainSkipped + ')</span>' : '')]\n        .concat(p.archives.map(a => '<b>На лист «' + esc(a.name) + '»: ' + a.count + '</b>' + (a.skipped ? ' <span class=\"small\">(уже есть — пропущу ' + a.skipped + ')</span>' : '')));\n      $('#plan').innerHTML = total\n        ? lines.join('<br>') + '<br><span class=\"small\">Ссылки на Band и сметы станут ссылками, листы отсортируются по дате: свежие сверху.</span>'\n        : 'Переносить нечего: всё из файла уже есть в таблице.';\n      $('#goBtn').disabled = !total;\n      setMsg($('#msg'), '');\n    }).catch(e => setMsg($('#msg'), e.message, 'err'));\n  }\n\n  $('#file').addEventListener('change', e => readFile(e.target.files[0]));\n  const drop = $('#drop');\n  ['dragenter', 'dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));\n  ['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));\n  drop.addEventListener('drop', e => readFile(e.dataTransfer.files[0]));\n  $('#link').addEventListener('input', () => {\n    source = null; mode = ''; resetPlan();\n    $('#fileName').textContent = 'Выберите файл Excel';\n  });\n  $('#checkBtn').addEventListener('click', check);\n\n  $('#goBtn').addEventListener('click', () => {\n    const btn = $('#goBtn');\n    btn.disabled = true; $('#checkBtn').disabled = true;\n    btn.textContent = 'Переношу… (до пары минут)';\n    const tasksFile = source && source.kind === 'tasks';\n    const call = tasksFile ? run('migrationApplyFromTasksFile', source)\n      : mode === 'file' ? run('migrationApplyFromData', source) : run('migrationApplyFromLink', $('#link').value);\n    call.then(r => {\n      btn.textContent = 'Готово';\n      if (tasksFile) {\n        $('#plan').innerHTML = '<b>Готово: в задачи — ' + r.added + ', в архив — ' + r.archived + '.</b> Цвета, списки и ссылки обновлены.';\n        return setMsg($('#msg'), '');\n      }\n      $('#plan').innerHTML = '<b>Перенесено задач: ' + r.added + '.</b> Лист отсортирован по дате, цвета и списки обновлены.' +\n        (mode === 'link' ? '<br><span class=\"small\">Новые номера теперь сверяются со старой таблицей. Появятся там новые задачи — запустите перенос ещё раз, он добавит только новые.</span>'\n                         : '<br><span class=\"small\">Появятся в старой таблице новые задачи — скачайте файл заново и перенесите ещё раз: добавятся только новые.</span>');\n      setMsg($('#msg'), '');\n    }).catch(e => {\n      setMsg($('#msg'), 'Не перенеслось: ' + e.message, 'err');\n      btn.disabled = false; $('#checkBtn').disabled = false; btn.textContent = 'Перенести';\n    });\n  });\n</script>\n</body>\n</html>\n"
};
