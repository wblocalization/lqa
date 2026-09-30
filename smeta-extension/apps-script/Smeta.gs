/**
 * Приёмник для расширения «Сметы → таблица».
 * Принимает номер задачи, итог с НДС и ссылку на смету и пишет их в лист задач менеджеров,
 * а в «📝 Журнал» — что поменялось.
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

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (req.action === 'lookup') return smetaJson_(smetaLookup_(req));
    if (req.action === 'write') return smetaJson_(smetaWrite_(req));
    return smetaJson_({ ok: false, error: 'Неизвестное действие' });
  } catch (err) {
    return smetaJson_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function smetaLookup_(req) {
  const t = smetaFindTask_(req.task);
  if (!t) return { ok: true, found: false };
  const v = t.values;
  return {
    ok: true, found: true, row: t.row,
    subject: v[t.col.subject], contractor: v[t.col.contractor], manager: v[t.col.manager],
    link: v[t.col.link], total: v[t.col.total],
  };
}

function smetaWrite_(req) {
  const total = Number(req.total);
  const link = String(req.link || '').trim();
  if (!(total > 0)) return { ok: false, error: 'Сумма должна быть больше нуля' };
  if (!/^https?:\/\//i.test(link)) return { ok: false, error: 'Ссылка должна начинаться с http' };

  const t = smetaFindTask_(req.task);
  if (!t) return { ok: false, error: 'Задача ' + req.task + ' не найдена в листе «' + SMETA_TASKS_SHEET + '»' };

  const oldLink = t.values[t.col.link];
  const oldTotal = t.values[t.col.total];
  if ((oldLink || oldTotal) && !req.overwrite) {
    return { ok: false, error: 'exists', link: oldLink, total: oldTotal };
  }

  const sheet = t.sheet;
  sheet.getRange(t.row, t.col.link + 1).setValue(link);
  sheet.getRange(t.row, t.col.total + 1).setValue(Math.round(total * 100) / 100);

  const who = 'Расширение смет' + (req.user ? ' (' + req.user + ')' : '');
  smetaLog_(who, t.taskId, SMETA_H.link, oldLink, link);
  smetaLog_(who, t.taskId, SMETA_H.total, oldTotal, total);
  return { ok: true, row: t.row };
}

function smetaFindTask_(task) {
  const taskId = String(task || '').trim().toUpperCase();
  if (!/^[A-Z]+-\d+$/.test(taskId)) throw new Error('Номер задачи должен быть вида LIT-26');

  const sheet = smetaSpreadsheet_().getSheetByName(SMETA_TASKS_SHEET);
  if (!sheet) throw new Error('Нет листа «' + SMETA_TASKS_SHEET + '»');
  const data = sheet.getDataRange().getValues();
  const header = data[0].map(function (h) { return String(h).trim(); });

  const col = {};
  Object.keys(SMETA_H).forEach(function (k) {
    col[k] = header.indexOf(SMETA_H[k]);
    if (col[k] < 0) throw new Error('Не найдена колонка «' + SMETA_H[k] + '»');
  });

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col.task]).trim().toUpperCase() === taskId) {
      return { sheet: sheet, row: i + 1, values: data[i], col: col, taskId: taskId };
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
