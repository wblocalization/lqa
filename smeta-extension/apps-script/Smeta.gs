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

// ID таблицы — кусок адреса между /d/ и /edit.
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
    // Запросы из расширения выполняются от имени владельца таблицы, поэтому «кто я» берём только из настроек расширения
    if (req.action === 'checkDuplicates') { requireTableScript_(); return smetaJson_({ ok: true, duplicates: findDuplicateTasks(req.task || {}) }); }
    if (req.action === 'myTasks') { requireTableScript_(); return smetaJson_(Object.assign({ ok: true }, getMyOpenTasks(req.manager || '*'))); }
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
    subject: v[t.col.subject], contractor: v[t.col.contractor], manager: v[t.col.manager],
    link: smetaCellLink_(t, v), total: v[t.col.total],
  };
}


/** Адрес сметы: ссылка под «Смета [номер]» или сам текст ячейки, если там адрес. */
function smetaCellLink_(t, values) {
  const rt = t.sheet.getRange(t.row, t.col.link + 1).getRichTextValue();
  if (rt) {
    const url = rt.getLinkUrl() || rt.getRuns().map(function (r) { return r.getLinkUrl(); }).filter(Boolean)[0];
    if (url) return url;
  }
  const v = String(values[t.col.link] || '').trim();
  return /^Смета( \[[^\]]*\])?$/.test(v) ? '' : v;
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
  // В ячейке — короткое «Смета [LIT-26-2232]», адрес — ссылкой под ним
  const rowId = String(t.values[t.col.task] || '').trim();
  sheet.getRange(t.row, t.col.link + 1).setRichTextValue(
    SpreadsheetApp.newRichTextValue().setText('Смета' + (rowId ? ' [' + rowId + ']' : '')).setLinkUrl(link).build());
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

function smetaSpreadsheet_() {
  return SMETA_SPREADSHEET_ID
    ? SpreadsheetApp.openById(SMETA_SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
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
  ['link', 'link2', 'estimateLink'].forEach(function (k) {
    if (task[k] && !/^https?:\/\//i.test(task[k])) throw new Error('Ссылка должна начинаться с http: ' + task[k]);
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
function setStatus_(req) {
  requireTableScript_();
  LOG_ACTOR = 'Расширение' + (req.user ? ' (' + req.user + ')' : '');
  try {
    return setTaskStatus(req.row, req.id, req.origSubject, req.status);
  } finally {
    LOG_ACTOR = '';
  }
}
