/*
 * Две правки для Код.gs. Найдите в Код.gs эти функции и замените их ЦЕЛИКОМ на то, что ниже
 * (от строки «function …» до её закрывающей «}»). Константы вставьте рядом с функцией.
 *
 * 1) generateNextTaskId — номер задачи в формате ПОДРЯДЧИК-ГГ-N (LIT-26-2232),
 *    как в старой таблице и в именах смет.
 * 2) logChange — чтобы в «Журнале» было видно, что задачу добавили из расширения.
 */

// ===== 1. Номер задачи =====

// Нумерация продолжается со старой таблицы: здесь последние номера оттуда (на 30.09.2026).
// Новая таблица берёт больший из двух: это число или последний номер на листе.
const TASK_ID_START = {
  'LIT-26': 2231,
  'BP-26': 366,
  'LoG-26': 75,
  'JWW-26': 29,
  'AWT-26': 1
};

function generateNextTaskId(tgt, contractor) {
  const prefix = CONTRACTOR_PREFIX[contractor] || 'TASK';
  const yy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yy');
  const base = prefix + '-' + yy; // LIT-26
  const lastRow = tgt.getLastRow();
  const values = lastRow >= 2 ? tgt.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
  let maxSeq = TASK_ID_START[base] || 0;
  const re = new RegExp('^' + base + '-(\\d+)$');
  values.forEach(v => {
    const m = String(v).trim().match(re);
    if (m) maxSeq = Math.max(maxSeq, Number(m[1]));
  });
  return base + '-' + (maxSeq + 1);
}

// ===== 2. Журнал =====

// Кто действует, если это не человек в таблице (например, расширение). Заполняет Smeta.gs.
var LOG_ACTOR = '';

function logChange(action, refId, field, oldVal, newVal) {
  try {
    const log = SpreadsheetApp.getActive().getSheetByName('📝 Журнал');
    const email = LOG_ACTOR || Session.getActiveUser().getEmail() || '(неизвестно)';
    log.appendRow([new Date(), email, action, refId, field || '', oldVal || '', newVal || '']);
  } catch (e) {}
}
