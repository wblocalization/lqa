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
  return ContentService.createTextOutput('Скрипт работает. Адрес правильный — вставьте его в настройки расширения.\n\nВерсия: 5 октября (удаление задач из расширения).');
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
