/*************************************************************
 * Перенос.gs — разовый перенос истории из старой таблицы
 * (лист «Localization Misc») в «📌 Задачи (менеджеры)».
 *
 * Меню: «⚙️ Настройки → 📥 Перенести историю из старой таблицы».
 * Сначала показывает, сколько задач перенесёт и сколько пропустит, и переносит только после подтверждения.
 * Задачи, номера которых уже есть в новой таблице, пропускаются — запускать повторно безопасно.
 * После переноса файл можно удалить.
 *************************************************************/

const OLD_SHEET_NAME = 'Localization Misc';

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

function migrateFromOldTable() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Перенос истории',
    'Вставьте ссылку на старую таблицу (где лист «' + OLD_SHEET_NAME + '»):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const m = resp.getResponseText().match(/\/d\/([a-zA-Z0-9_-]{20,})/) || resp.getResponseText().trim().match(/^([a-zA-Z0-9_-]{20,})$/);
  if (!m) { ui.alert('Не похоже на ссылку на Google Таблицу.'); return; }

  const plan = planMigration_(SpreadsheetApp.openById(m[1]));
  if (!plan.rows.length) {
    ui.alert('Переносить нечего: ' + plan.skippedExisting + ' задач уже есть в новой таблице, ' + plan.skippedEmpty + ' строк пустые.');
    return;
  }
  const ok = ui.alert('Перенос истории',
    'Будет добавлено задач: ' + plan.rows.length + ' (с ' + plan.from + ' по ' + plan.to + ').\n' +
    'Уже есть в новой таблице — пропущу: ' + plan.skippedExisting + '.\n' +
    'Пустые строки и повторы шапки — пропущу: ' + plan.skippedEmpty + '.\n\n' +
    'Задачи встанут под текущими, свежие выше. Продолжить?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  applyMigration_(plan);
  ui.alert('Готово: перенесено ' + plan.rows.length + ' задач. Цвета и выпадающие списки обновлены.');
}

function isUrl_(s) {
  return /^https?:\/\//i.test(String(s || '').trim());
}

function numOrEmpty_(v) {
  if (v === '' || v === null) return '';
  const n = Number(String(v).replace(/\s/g, '').replace(',', '.'));
  return isFinite(n) ? n : v;
}

/** Читает старый лист и готовит строки для новой таблицы (ничего не записывает). */
function planMigration_(oldSs) {
  const old = oldSs.getSheetByName(OLD_SHEET_NAME);
  if (!old) throw new Error('В старой таблице нет листа «' + OLD_SHEET_NAME + '»');
  const lastRow = old.getLastRow();
  const range = old.getRange(1, 1, lastRow, OLD_COLS);
  const values = range.getValues();
  const bandRich = old.getRange(1, OLD.BAND, lastRow, 1).getRichTextValues().map(r => r[0]);

  // Раскопировать объединённые ячейки на все их строки
  range.getMergedRanges().forEach(mr => {
    const r0 = mr.getRow(), c0 = mr.getColumn();
    for (let c = c0; c < c0 + mr.getNumColumns(); c++) {
      if (OLD_NO_FILL.indexOf(c) !== -1) continue;
      const v = values[r0 - 1][c - 1];
      for (let r = r0; r < r0 + mr.getNumRows(); r++) {
        values[r - 1][c - 1] = v;
        if (c === OLD.BAND) bandRich[r - 1] = bandRich[r0 - 1];
      }
    }
  });

  // Что уже есть в новой таблице: по номеру, а строки без номера — по теме, дате и ссылке
  const tasks = getTasksSheet();
  const existing = {}, existingKeys = {};
  const cur = readRows_(tasks, COL.DATE);
  const curLinks = subjectLinks_(tasks);
  const keyOf = (subject, date, link) => [str_(subject).replace(LINK2_MARKER, ''), fmtDate_(date, 'yyyy-MM-dd') || str_(date), link || ''].join('|');
  cur.forEach((r, i) => {
    if (str_(r[COL.ID - 1])) existing[str_(r[COL.ID - 1])] = true;
    else existingKeys[keyOf(r[COL.SUBJECT - 1], r[COL.DATE - 1], curLinks[i] && curLinks[i].link)] = true;
  });

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
    const comment = [
      get(OLD.COMMENT),
      taskText && !ticketMatch && taskText !== subject ? 'Задача: ' + taskText : '',
      get(OLD.COMPLAINTS) ? 'Жалобы на заказчика: ' + get(OLD.COMPLAINTS) : ''
    ].filter(Boolean).join('\n');

    const band = linksFromRich_(bandRich[i - 1]).link || (isUrl_(get(OLD.BAND)) ? get(OLD.BAND) : '');
    if (!id && existingKeys[keyOf(subject || taskText || '(без темы)', v[OLD.DATE - 1], band)]) { skippedExisting++; continue; }
    const deadline = get(OLD.DEADLINE) === 'Дедлайн (если есть)' ? '' : get(OLD.DEADLINE);
    const side = get(OLD.SIDE);
    const date = v[OLD.DATE - 1];

    rows.push({
      link: band,
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
        .setRichTextValues(part.map(p => [buildSubjectRich_(String(p.values[COL.SUBJECT - 1]), p.link, '')]));
    }
    logChange('Перенос истории', '', '', '', n + ' задач из старой таблицы');
  });
  designTasksSheet_(getTasksSheet());
}
