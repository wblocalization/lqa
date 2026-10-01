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
    if (cache) cache.put(key, String(max), 60);
  } catch (e) {
    // старая таблица недоступна (или правка прямо в листе, где её открывать нельзя) — считаем только по новой
  }
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
function findDuplicateTasks(task) {
  const sh = getTasksSheet();
  const rows = readRows_(sh, TASK_COLS);
  const links = subjectLinks_(sh);
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
  const links = subjectLinks_(sh);
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
      link: links[i] ? links[i].link : '',
      sort: filter === 'open' ? (dueTime === null ? Infinity : dueTime) : -(date instanceof Date ? date.getTime() : 0)
    });
  });
  tasks.sort((a, b) => a.sort - b.sort);
  tasks.forEach(t => { delete t.sort; });
  const lists = getListsData();
  return { manager: manager, filter: filter, tasks: filter === 'open' ? tasks : tasks.slice(0, MY_TASKS_LIMIT),
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
