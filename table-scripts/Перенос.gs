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
      const v = row[c];
      if (v && typeof v === 'object' && v.d) {
        const p = String(v.d).split('-').map(Number);
        out.push(new Date(p[0], p[1] - 1, p[2]));
      } else out.push(v == null ? '' : v);
    }
    return out;
  });
  return { values: values, links: (src.links || []).map(l => l || ''), merges: src.merges || [] };
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

  // Что уже есть в новой таблице: по номеру, а строки без номера — по теме, дате и ссылке
  const tasks = getTasksSheet();
  const existing = {}, existingKeys = {};
  const cur = readRows_(tasks, COL.DATE);
  const curLinks = subjectLinks_(tasks);
  const keyOf = (subject, date, link) => [stripLinkMarkers_(subject), fmtDate_(date, 'yyyy-MM-dd') || str_(date), link || ''].join('|');
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
