// Разбор текста сметы: номер задачи и итоговая сумма с НДС.
// Модуль без зависимостей от браузера — его же гоняют тесты в Node.

// Ставки НДС, которые встречаются в сметах (доли от суммы без НДС).
const VAT_RATES = [0.22, 0.2, 0.1, 0.07, 0.05];

// «110 929,150», «5 937,50», «90 925,53» — пробелы (в т.ч. неразрывные) между тысячами, запятая перед копейками.
const MONEY_RE = /(?<![\d,])\d{1,3}(?:[   ]\d{3})*,\d{2,3}(?![\d,])/g;

export function parseMoney(s) {
  return Number(s.replace(/[   ]/g, '').replace(',', '.'));
}

/** «110 929,155» → 110929155 тысячных рубля — целым числом, без ошибок дробей вроде 1,005 × 100 = 100,4999… */
function toMilli(s) {
  const [int, frac = ''] = s.replace(/[   ]/g, '').split(',');
  return Number(int) * 1000 + Number((frac + '000').slice(0, 3));
}
const milliToCents = (m) => Math.round(m / 10); // целые — округление точное, половинка вверх

/** Суммы в тысячных рубля, как напечатаны в PDF (уникальные). */
function extractMilli(texts) {
  const milli = new Set();
  for (const t of texts) {
    for (const m of String(t).matchAll(MONEY_RE)) milli.add(toMilli(m[0]));
  }
  return [...milli].filter((x) => x > 0);
}

/** Все денежные суммы из кусков текста PDF (уникальные, в копейках) — для поиска тройки «без НДС + НДС = с НДС». */
export function extractAmounts(texts) {
  return [...new Set(extractMilli(texts).map(milliToCents))];
}

/**
 * Сумма ровно как в PDF, без округления: копейки → какое число было напечатано.
 * В сметах три знака после запятой («102 641,640»); если третий не ноль — вернём его как есть.
 */
function exactOf(cents, milliList) {
  const hits = milliList.filter((m) => milliToCents(m) === cents);
  const m = hits.find((x) => x % 10 !== 0) ?? hits[0] ?? cents * 10;
  return m / 1000;
}

/** Есть ли у суммы ненулевой третий знак после запятой (доли копейки). */
export function hasExtraDigits(x) {
  return Math.round(Math.abs(x) * 1000) % 10 !== 0;
}

/**
 * Строки страницы по координатам: в PDF подпись «Версия» и её значение «1» лежат в разных местах текстового слоя,
 * но на одной высоте. items — { str, x, y, page }.
 */
export function linesFromItems(items) {
  const sorted = items.filter((i) => String(i.str).trim()).slice().sort((a, b) => (a.page - b.page) || (b.y - a.y) || (a.x - b.x));
  const lines = [];
  let cur = null;
  for (const it of sorted) {
    if (!cur || it.page !== cur.page || Math.abs(it.y - cur.y) > 2) {
      cur = { page: it.page, y: it.y, items: [] };
      lines.push(cur);
    }
    cur.items.push(it);
  }
  return lines.map((l) => l.items.sort((a, b) => a.x - b.x).map((i) => String(i.str).trim()).join(' '));
}

/** «Версия 2» из шапки сметы. */
export function versionFromTexts(texts) {
  const m = texts.join(' ').match(/Версия\s*:?\s*(\d{1,3})(?!\d)/i);
  return m ? Number(m[1]) : null;
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
export function parseEstimate(texts, fileName = '', lines = []) {
  const task = taskFromFileName(fileName) || taskFromTexts(texts);
  const found = findTotalWithVat(extractAmounts(texts));
  let total = null;
  if (found) {
    // Числа — ровно как напечатаны (до тысячных), чтобы ничего не округлилось молча
    const milli = extractMilli(texts);
    const exact = (x) => exactOf(Math.round(x * 100), milli);
    total = { net: exact(found.net), vat: exact(found.vat), gross: exact(found.gross), rate: found.rate };
    total.extraDigits = hasExtraDigits(total.gross);
  }
  return { task, total, version: versionFromTexts(lines) ?? versionFromTexts(texts) };
}

/** 110929.15 → «110 929,15»; доли копейки не прячем: 110929.155 → «110 929,155» */
export function formatMoney(x) {
  const digits = hasExtraDigits(x) ? 3 : 2;
  return x.toLocaleString('ru-RU', { minimumFractionDigits: digits, maximumFractionDigits: digits });
}
