// Разбор текста сметы: номер задачи и итоговая сумма с НДС.
// Модуль без зависимостей от браузера — его же гоняют тесты в Node.

// Ставки НДС, которые встречаются в сметах (доли от суммы без НДС).
const VAT_RATES = [0.22, 0.2, 0.1, 0.07, 0.05];

// «110 929,150», «5 937,50», «90 925,53» — пробелы (в т.ч. неразрывные) между тысячами, запятая перед копейками.
const MONEY_RE = /(?<![\d,])\d{1,3}(?:[   ]\d{3})*,\d{2,3}(?![\d,])/g;

export function parseMoney(s) {
  return Number(s.replace(/[   ]/g, '').replace(',', '.'));
}

const toCents = (x) => Math.round(x * 100);

/** Все денежные суммы из кусков текста PDF (уникальные, в копейках). */
export function extractAmounts(texts) {
  const cents = new Set();
  for (const t of texts) {
    for (const m of String(t).matchAll(MONEY_RE)) cents.add(toCents(parseMoney(m[0])));
  }
  return [...cents].filter((c) => c > 0);
}

/**
 * Ищет тройку «без НДС + НДС = с НДС», где НДС — одна из стандартных ставок.
 * В текстовом слое PDF подписи оторваны от чисел, поэтому опираемся на арифметику,
 * а не на соседство со словами. Если троек несколько — берём самую большую сумму.
 */
export function findTotalWithVat(amountsCents) {
  const set = new Set(amountsCents);
  let best = null;
  for (const net of amountsCents) {
    for (const rate of VAT_RATES) {
      const expectedVat = net * rate;
      for (const vat of amountsCents) {
        if (Math.abs(vat - expectedVat) > 5) continue; // допуск 5 копеек на округления
        for (const d of [0, -1, 1, -2, 2]) {
          const gross = net + vat + d;
          if (set.has(gross) && (!best || gross > best.gross)) {
            best = { net, vat, gross, rate };
          }
        }
      }
    }
  }
  if (!best) return null;
  return { net: best.net / 100, vat: best.vat / 100, gross: best.gross / 100, rate: best.rate };
}

// Код подрядчика — 2–6 латинских букв перед номером: LIT-26, GLB-5.
// Номер сметы «LIT-26-2203» (код-задача-смета) берём целиком, поэтому ищем его первым.
const ESTIMATE_NO_RE = /(?<![A-Za-z])([A-Za-z]{2,6})[-_](\d+)[-_](\d+)(?!\d)/;
const TASK_RE = /(?<![A-Za-z0-9])([A-Za-z]{2,6})[-_](\d+)(?!\d)/;

const toTask = (m) => [m[1].toUpperCase(), m[2], m[3]].filter(Boolean).join('-');

// «-РВБ» / «_RWB» / « РВБ» в имени файла: всё, что до него, и есть номер.
// Хвост вроде « (1)» от повторной загрузки тоже отрезается.
const RWB_SUFFIX_RE = /[\s_-]*(?:РВБ|RWB)(?![A-Za-zА-Яа-яЁё]).*$/i;

/**
 * Номер из имени файла: «HELLO-3123-2133132-RWB.pdf» → HELLO-3123-2133132.
 * Берётся всё до «-РВБ»/«-RWB». Если такого хвоста нет — ищем номер вида LIT-26-2203.
 */
export function taskFromFileName(fileName) {
  const name = String(fileName).replace(/\.pdf$/i, '').trim();
  if (RWB_SUFFIX_RE.test(name)) {
    const id = name.replace(RWB_SUFFIX_RE, '').trim();
    return id ? id.toUpperCase() : null;
  }
  const atStart = name.match(new RegExp('^' + ESTIMATE_NO_RE.source)) || name.match(new RegExp('^' + TASK_RE.source));
  const m = atStart || name.match(ESTIMATE_NO_RE) || name.match(TASK_RE);
  return m ? toTask(m) : null;
}

/** Запасной вариант — номер сметы в тексте PDF (там же есть чужие коды вроде WB-2063, поэтому только трёхчастный). */
export function taskFromTexts(texts) {
  for (const t of texts) {
    const m = String(t).match(ESTIMATE_NO_RE);
    if (m) return toTask(m);
  }
  return null;
}

/**
 * Главная функция: принимает куски текста PDF и имя файла.
 * Номер задачи берём из имени файла, а если его там нет — из номера сметы в тексте.
 */
export function parseEstimate(texts, fileName = '') {
  const task = taskFromFileName(fileName) || taskFromTexts(texts);
  const total = findTotalWithVat(extractAmounts(texts));
  return { task, total };
}

/** 110929.15 → «110 929,15» */
export function formatMoney(x) {
  return x.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
