/**
 * Приёмник для расширения «Сметы → таблица».
 * Принимает номер задачи, итог с НДС и ссылку на смету и пишет их в лист задач менеджеров.
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
    // Отчёты, сверка и письма — вкладки «Отчёты» и «Письма» в расширении
    if (req.action === 'reportLists') return smetaJson_(reportLists_());
    if (req.action === 'managerReport') { requireTableScript_(); return smetaJson_(Object.assign({ ok: true }, getManagerReport(String(req.manager || ''), String(req.month || ''), req.by))); }
    if (req.action === 'trackerReport') { requireTableScript_(); return smetaJson_(Object.assign({ ok: true }, getTrackerReportText(String(req.ticket || '')))); }
    if (req.action === 'customReport') { requireTableScript_(); return smetaJson_(Object.assign({ ok: true }, getCustomReport(req.year || '', req.month || '', req.by))); }
    if (req.action === 'moneyPreview') { requireTableScript_(); return smetaJson_(Object.assign({ ok: true }, moneyExportPreview(req.opts || {}))); }
    if (req.action === 'exportXlsx') return smetaJson_(exportXlsx_(req));
    if (req.action === 'mailStatus') return smetaJson_(mailStatus_(req));
    if (req.action === 'mailTest') return smetaJson_(mailTest_(req));
    if (req.action === 'personalGet') return smetaJson_(personalGet_(req));
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
    if (req.action === 'addTaskRow') return smetaJson_(addTaskRow_(req));
    if (req.action === 'setStatus') return smetaJson_(setStatus_(req));
    if (req.action === 'saveTask') return smetaJson_(saveTask_(req));
    if (req.action === 'setDelivery') { requireTableScript_(); return smetaJson_(setTaskDelivery(req.row, req.id, req.origSubject, req.value)); }
    if (req.action === 'mailToggle') return smetaJson_(mailToggle_(req));
    if (req.action === 'mailEmail') return smetaJson_(mailEmail_(req));
    if (req.action === 'personalSave') return smetaJson_(personalSave_(req));
    if (req.action === 'deleteTask') { requireTableScript_(); deleteTask(req.row, req.id, req.origSubject); return smetaJson_({ ok: true }); }
    return smetaJson_({ ok: false, error: 'Неизвестное действие' });
  } catch (err) {
    return smetaJson_({ ok: false, error: String(err && err.message || err) });
  } finally {
    SCRIPT_LOCK_HELD = false;
    SpreadsheetApp.flush(); // записать номер до того, как следующий его прочитает
    lock.releaseLock();
  }
}

/** Проверка: открыть адрес веб-приложения в браузере — должно написать, что скрипт работает. */
function doGet() {
  return ContentService.createTextOutput('Скрипт работает. Адрес правильный — вставьте его в настройки расширения.\n\nВерсия: 9 октября, ночь (личные шаблоны в таблице, строки в заказ, дедлайн по рабочим дням).');
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
  // В ячейке — сам адрес ссылкой: так таблица читает его мгновенно (сверка, отчёты), а в Excel он виден как есть
  sheet.getRange(t.row, t.col.link + 1).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(link).setLinkUrl(link).build());
  // Сумма — ровно как в смете, без округления до копеек (в сметах бывает три знака после запятой)
  const exact = Math.round(total * 1000) / 1000;
  const cell = sheet.getRange(t.row, t.col.total + 1);
  cell.setValue(exact).setNumberFormat(smetaMoneyFormat_(exact));
  // Подрядчик обновил смету — прежняя сумма и ссылка остаются в заметке к ячейке (для сверки)
  if (req.overwrite && (oldLink || oldTotal !== '')) {
    const changed = oldTotal === '' || Number(oldTotal) !== exact || (oldLink && oldLink !== link);
    if (changed) {
      const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
      const ver = req.version ? ' (версия ' + req.version + ')' : '';
      const sum = oldTotal === '' ? 'сумма ' + smetaRub_(exact) + ' ₽'
        : Number(oldTotal) === exact ? 'сумма та же: ' + smetaRub_(exact) + ' ₽'
        : 'было ' + smetaRub_(oldTotal) + ' ₽ → стало ' + smetaRub_(exact) + ' ₽';
      const line = date + ' — смета обновлена' + ver + ': ' + sum + (oldLink && oldLink !== link ? '\nпрежняя смета: ' + oldLink : '');
      const note = cell.getNote();
      cell.setNote(note ? note + '\n\n' + line : line);
    }
  }

  return { ok: true, row: t.row, url: tableUrl_(t.row), was: req.overwrite ? oldTotal : '' };
}

/** 102641.64 → «102 641,64»; доли копейки не прячем: 110929.155 → «110 929,155». */
function smetaRub_(x) {
  if (x === '' || x == null || !isFinite(Number(x))) return String(x == null ? '' : x); // сумма текстом — как есть
  const n = Math.round(Number(x) * 1000) / 1000;
  const digits = Math.round(Math.abs(n) * 1000) % 10 ? 3 : 2;
  const parts = Math.abs(n).toFixed(digits).split('.');
  return (n < 0 ? '−' : '') + parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ',' + parts[1];
}

/** Формат ячейки суммы: два знака, а если в сумме доли копейки — три, чтобы таблица их не прятала. */
function smetaMoneyFormat_(x) {
  return Math.round(Math.abs(Number(x)) * 1000) % 10 ? '#,##0.000' : '#,##0.00';
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
  const lists = getAddTaskFormLists();
  lists.tableUrl = tableUrl_(); // кнопка «Таблица ↗» в расширении
  return { ok: true, lists: lists };
}

/** Адрес листа задач; row — сразу на нужную строку. */
function tableUrl_(row) {
  const url = SpreadsheetApp.getActive().getUrl().replace(/#.*$/, '') + '#gid=' + getTasksSheet().getSheetId();
  return row ? url + '&range=A' + row : url;
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
    complaints: str(t.complaints),
    languages: (Array.isArray(t.languages) ? t.languages : []).map(str).filter(Boolean),
  };
  if (!task.contractor) return { ok: false, error: 'Выберите подрядчика' };
  if (!task.subject) return { ok: false, error: 'Впишите тему' };
  if (task.ticket && !/^LOCAL-\d+$/.test(task.ticket)) return { ok: false, error: 'Тикет должен быть вида LOCAL-1234' };
  // link2 — доп. ссылки на Band, по одной в строке
  [task.link, task.estimateLink].concat(task.link2.split(/[\s,;]+/)).forEach(function (u) {
    if (u && !/^https?:\/\//i.test(u)) throw new Error('Ссылка должна начинаться с http: ' + u);
  });

  const id = submitNewTaskFromDialog(task);
  return { ok: true, id: id, url: tableUrl_(2) }; // новая задача — всегда вторая строка, под шапкой
}

/** Ещё строка в существующий заказ (свои продукт, Band, ник, языки, сроки) — сразу под его строками. */
function addTaskRow_(req) {
  requireTableScript_();
  const t = req.task || {};
  const str = function (v) { return String(v == null ? '' : v).trim(); };
  const task = {
    ticket: str(t.ticket), link: str(t.link), link2: str(t.link2), date: str(t.date), product: str(t.product),
    customer: str(t.customer), deadline: str(t.deadline), exactDeadline: str(t.exactDeadline), status: str(t.status),
    deliveryStatus: str(t.deliveryStatus), sp: str(t.sp), manager: str(t.manager), comment: str(t.comment),
    languages: (Array.isArray(t.languages) ? t.languages : []).map(str).filter(Boolean),
  };
  if (task.ticket && !/^LOCAL-\d+$/.test(task.ticket)) return { ok: false, error: 'Тикет должен быть вида LOCAL-1234' };
  [task.link].concat(task.link2.split(/[\s,;]+/)).forEach(function (u) {
    if (u && !/^https?:\/\//i.test(u)) throw new Error('Ссылка должна начинаться с http: ' + u);
  });
  const r = addRowToTask(str(req.id), task);
  return { ok: true, id: r.id, row: r.row, subject: r.subject, url: tableUrl_(r.row) };
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
  return saveTaskEdits(t);
}

function setStatus_(req) {
  requireTableScript_();
  return setTaskStatus(req.row, req.id, req.origSubject, req.status);
}

/* ============================================================
 *  ОТЧЁТЫ, СВЕРКА И ПИСЬМА ИЗ РАСШИРЕНИЯ
 *  Те же функции, что у окон в таблице (Код.gs).
 * ============================================================ */

/** Всё для фильтров вкладки «Отчёты» одним запросом. */
function reportLists_() {
  requireTableScript_();
  const lists = getListsData();
  return {
    ok: true, managers: lists.managers, contractors: lists.contractors, langs: lists.langs,
    months: getMonthsList(), years: getDashboardYears(), tickets: getTicketsForReport(),
    current: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM')
  };
}

/**
 * Excel прямо в расширение: файл собирается как в таблице, скачивается сюда (base64), а временный файл
 * уходит в корзину Диска. Так ссылка не нужна — у коллеги может не быть доступа к Диску владельца.
 */
function exportXlsx_(req) {
  requireTableScript_();
  let url;
  if (req.kind === 'manager') url = exportReportToDoc(String(req.manager || ''), String(req.month || ''), req.by);
  else if (req.kind === 'custom') url = exportCustomReportToExcel(req.year || '', req.month || '', req.by);
  else if (req.kind === 'money') url = exportMoneyExcel(req.opts || {}).url;
  else return { ok: false, error: 'Неизвестный отчёт' };
  const id = url.match(/\/d\/([^/]+)/)[1];
  const file = DriveApp.getFileById(id);
  const name = file.getName();
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  file.setTrashed(true);
  return { ok: true, name: name.replace(/[\\/:*?"<>|]+/g, ' ') + '.xlsx', b64: Utilities.base64Encode(blob.getBytes()) };
}

// Письма: расписание общее на всю команду, почта — у каждого своя (колонка M «Списков»)
const MAIL_KINDS = {
  weekly: { handler: 'sendWeeklyDigests', send: function (m) { return sendWeeklyDigests(m); }, make: function (b) { return b.onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9); } },
  due: { handler: 'sendDueSoonAlerts', send: function (m) { return sendDueSoonAlerts(m); }, make: function (b) { return b.everyDays(1).atHour(9); } },
  month: { handler: 'sendMonthEndReminders', send: function (m) { return sendMonthEndReminders(m); }, make: function (b) { return b.onMonthDay(26).atHour(10); } }
};

function mailStatus_(req) {
  requireTableScript_();
  const me = req.manager ? managerMailList_(String(req.manager))[0] : null;
  const out = { ok: true, known: !!me, email: me ? me.email : '' };
  // Расписания видны только с разрешением «script.scriptapp»; без него почта и «прислать мне» всё равно работают
  let handlers = null;
  try { handlers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }); } catch (e) {
    out.noTriggers = String(e.message || e);
    // от чьего имени работает веб-приложение — разрешение нужно дать именно этому аккаунту
    try { out.runAs = Session.getEffectiveUser().getEmail(); } catch (e2) { /* не страшно */ }
  }
  Object.keys(MAIL_KINDS).forEach(function (k) { out[k] = handlers ? handlers.indexOf(MAIL_KINDS[k].handler) !== -1 : false; });
  return out;
}

function mailToggle_(req) {
  requireTableScript_();
  const kind = MAIL_KINDS[req.kind];
  if (!kind) return { ok: false, error: 'Неизвестное письмо' };
  if (req.on) replaceTrigger_(kind.handler, kind.make);
  else ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === kind.handler) ScriptApp.deleteTrigger(t); });
  return mailStatus_(req);
}

/** «Прислать мне сейчас» — письмо только этому менеджеру. */
function mailTest_(req) {
  requireTableScript_();
  const kind = MAIL_KINDS[req.kind];
  if (!kind) return { ok: false, error: 'Неизвестное письмо' };
  const name = String(req.manager || '').trim();
  const me = name ? managerMailList_(name)[0] : null; // без имени managerMailList_ вернул бы всех
  if (!me) return { ok: false, error: 'Выберите себя в настройках расширения (⚙️ → «Кто вы»)' };
  if (!me.email) return { ok: false, error: 'Сначала впишите почту для писем' };
  return { ok: true, sent: kind.send(me.manager), email: me.email };
}

/** Почта для писем этого менеджера — «Списки», колонка M. Пусто — письма не приходят. */
function mailEmail_(req) {
  requireTableScript_();
  const manager = String(req.manager || '').trim();
  const email = String(req.email || '').trim();
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'Не похоже на почту: ' + email };
  const lst = listsSheet_();
  const names = lst.getRange(3, 5, 200, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    if (!names[i][0]) break;
    if (String(names[i][0]).trim() === manager) {
      lst.getRange(3 + i, 13).setValue(email);
      return mailStatus_(req);
    }
  }
  return { ok: false, error: 'Нет менеджера «' + manager + '» в «Списках»' };
}

// ---------- Личные шаблоны: шаблоны задач, письма подрядчикам и подпись каждого менеджера ----------
// Скрытый лист «Личные шаблоны»: A — менеджер, B — когда сохранено, C… — данные (JSON, кусками: в ячейку влезает 50 000 знаков).
const PERSONAL_SHEET = 'Личные шаблоны';
const PERSONAL_CHUNK = 45000;

function personalSheet_(create) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PERSONAL_SHEET);
  if (!sh && create) {
    sh = ss.insertSheet(PERSONAL_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['Менеджер', 'Сохранено', 'Данные (не редактировать руками)']]);
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function personalRow_(sh, manager) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const names = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) if (String(names[i][0]).trim() === manager) return i + 2;
  return 0;
}

function personalGet_(req) {
  const manager = String(req.manager || '').trim();
  if (!manager) return { ok: false, error: 'Не выбран менеджер (⚙️ → «Кто вы»)' };
  const sh = personalSheet_(false);
  const row = sh ? personalRow_(sh, manager) : 0;
  if (!row) return { ok: true, data: null };
  const width = Math.max(sh.getLastColumn() - 2, 1);
  const json = sh.getRange(row, 3, 1, width).getValues()[0].map((v) => String(v).replace(/^~/, '')).join('');
  let data = null;
  try { data = json ? JSON.parse(json) : null; } catch (e) { return { ok: false, error: 'Личные шаблоны в таблице повреждены' }; }
  return { ok: true, data: data };
}

function personalSave_(req) {
  const manager = String(req.manager || '').trim();
  if (!manager) return { ok: false, error: 'Не выбран менеджер (⚙️ → «Кто вы»)' };
  if (!req.data || typeof req.data !== 'object') return { ok: false, error: 'Нечего сохранять' };
  const json = JSON.stringify(req.data);
  const parts = [];
  for (let i = 0; i < json.length; i += PERSONAL_CHUNK) parts.push(json.slice(i, i + PERSONAL_CHUNK));
  if (parts.length > 40) return { ok: false, error: 'Слишком много шаблонов — не помещается в таблицу' };
  const sh = personalSheet_(true);
  let row = personalRow_(sh, manager);
  if (!row) row = Math.max(sh.getLastRow(), 1) + 1;
  const width = Math.max(sh.getLastColumn() - 2, parts.length, 1);
  const cells = [];
  // «~» впереди — чтобы кусок, начавшийся с «=» или «+», таблица не приняла за формулу
  for (let i = 0; i < width; i++) cells.push(parts[i] ? '~' + parts[i] : '');
  sh.getRange(row, 1, 1, 2).setValues([[manager, new Date()]]);
  sh.getRange(row, 3, 1, width).setNumberFormat('@').setValues([cells]);
  return { ok: true };
}
