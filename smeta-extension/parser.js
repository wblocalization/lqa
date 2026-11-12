// Разбор текста сметы: номер задачи и итоговая сумма с НДС.
// Модуль без зависимостей от браузера — его же гоняют тесты в Node.

// Ставки НДС, которые встречаются в сметах (доли от суммы без НДС).
const VAT_RATES = [0.22, 0.2, 0.1, 0.07, 0.05];

// Русский формат: «110 929,150», «5 937,50», «26134,40» — пробелы (в т.ч. неразрывные) между тысячами, запятая перед копейками.
const MONEY_RE = /(?<![\d,.])(?:\d{1,3}(?:[   ]\d{3})+|\d+),\d{2,3}(?![\d,])/g;
// Английский формат (Logrus Global): «21,262.50», «1,012.50» — запятая между тысячами, точка перед копейками.
const MONEY_EN_RE = /(?<![\d,.])(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2,3}(?![\d.])/g;
// Смета в английском формате — если в ней есть хоть одна сумма вида «1,012.50» (иначе «22.09» — это дата, а не деньги)
const isEnglishMoney = (texts) => texts.some((t) => /\d,\d{3}\.\d{2}(?!\d)/.test(String(t)));

export function parseMoney(s) {
  s = String(s).replace(/[   ]/g, '');
  if (/\.\d{2,3}$/.test(s)) return Number(s.replace(/,/g, '')); // 21,262.50
  return Number(s.replace(',', '.'));
}

/** «110 929,155» → 110929155 тысячных рубля — целым числом, без ошибок дробей вроде 1,005 × 100 = 100,4999… */
function toMilli(s, english = false) {
  const clean = s.replace(/[   ]/g, '');
  const [int, frac = ''] = english ? clean.replace(/,/g, '').split('.') : clean.split(',');
  return Number(int) * 1000 + Number((frac + '000').slice(0, 3));
}
const milliToCents = (m) => Math.round(m / 10); // целые — округление точное, половинка вверх

/** Суммы в тысячных рубля, как напечатаны в PDF (уникальные). */
function extractMilli(texts) {
  const milli = new Set();
  const english = isEnglishMoney(texts);
  for (const t of texts) {
    for (const m of String(t).matchAll(english ? MONEY_EN_RE : MONEY_RE)) milli.add(toMilli(m[0], english));
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
  if (m) return Number(m[1]);
  // Logrus Global: в шапке «Дата Версия», а значения — ниже: «2026-09-30 01»
  if (texts.some((t) => /Версия/i.test(t))) {
    for (const t of texts) { const d = String(t).match(/\b\d{4}-\d{2}-\d{2}\s+0?(\d{1,2})\s*$/); if (d) return Number(d[1]); }
  }
  return null;
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
  if (best) return { net: best.net / 100, vat: best.vat / 100, gross: best.gross / 100, rate: best.rate };
  return findVatIncluded(amountsCents);
}

/**
 * «Итого 26134,40 · в том числе НДС 5% 1244,50» (Janus): суммы без НДС в смете нет — НДС уже внутри итога.
 * Ищем пару «итог + НДС», где НДС = итог × ставка / (1 + ставка). Берём самый большой итог.
 */
function findVatIncluded(amountsCents) {
  let best = null;
  for (const gross of amountsCents) {
    for (const rate of VAT_RATES) {
      const vat = Math.round(gross * rate / (1 + rate));
      for (const v of amountsCents) {
        if (v !== gross && Math.abs(v - vat) <= 1 && (!best || gross > best.gross)) best = { gross, vat: v, rate };
      }
    }
  }
  if (!best) return null;
  return { net: (best.gross - best.vat) / 100, vat: best.vat / 100, gross: best.gross / 100, rate: best.rate, included: true };
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

// Номер задачи из таблицы: префикс подрядчика, год двумя цифрами, номер — LIT-26-2217, LoG-26-75, JWW-26-29
const STRICT_ID_RE = /(?<![A-Za-z])([A-Za-z]{2,6})-(\d{2})-(\d+)(?!\d)/;
const strictId = (s) => { const m = String(s).match(STRICT_ID_RE); return m ? `${m[1]}-${m[2]}-${m[3]}` : null; };

/**
 * Номер задачи для сметы любого подрядчика. По порядку:
 * 1) имя файла с «-РВБ» (как раньше: всё до него); 2) номер вида LoG-26-75 в имени файла;
 * 3) номер в квадратных скобках в тексте — «[JWW-26-29] [Janus] …» (у Janus имя файла — «Смета_Janus_LSP_480_00030»);
 * 4) такой же номер просто в тексте; 5) прежние запасные варианты.
 */
export function taskFor(fileName, texts = []) {
  const name = String(fileName).replace(/\.pdf$/i, '').trim();
  if (RWB_SUFFIX_RE.test(name)) return taskFromFileName(name);
  const inName = strictId(name);
  if (inName) return inName;
  for (const t of texts) { const m = String(t).match(/\[\s*([A-Za-z]{2,6}-\d{2}-\d+)\s*\]/); if (m) return m[1]; }
  for (const t of texts) { const id = strictId(t); if (id) return id; }
  return taskFromFileName(name) || taskFromTexts(texts);
}

/**
 * Главная функция: принимает куски текста PDF и имя файла.
 * Номер задачи — см. taskFor; сумма — «без НДС + НДС = с НДС» или «итог, в том числе НДС».
 */
/**
 * Слова, разбитые на отдельные буквы. Жирный шрифт в сметах Logrus набран «вразрядку» и ещё и дважды
 * (обводка поверх текста): «Итого с НДС» → «2»,«2»,«5»,«5»,«,»,«,»,«2»,«2»… Склеиваем такие серии
 * обратно в «25,200.00руб», убирая удвоение.
 */
export function mergeCharRuns(texts) {
  const out = [];
  let run = [];
  const flush = () => {
    if (run.length >= 4) {
      const doubled = run.length % 2 === 0 && run.every((c, i) => i % 2 || c === run[i + 1]);
      out.push((doubled ? run.filter((_, i) => i % 2 === 0) : run).join(''));
    } else out.push(...run);
    run = [];
  };
  for (const t of texts) {
    const s = String(t).trim();
    if (s.length === 1) run.push(s);
    else { flush(); out.push(t); }
  }
  flush();
  return out;
}

export function parseEstimate(texts, fileName = '', lines = []) {
  texts = mergeCharRuns(texts);
  const task = taskFor(fileName, [...lines, ...texts]);
  const found = findTotalWithVat(extractAmounts(texts));
  let total = null;
  if (found) {
    // Числа — ровно как напечатаны (до тысячных), чтобы ничего не округлилось молча
    const milli = extractMilli(texts);
    const exact = (x) => exactOf(Math.round(x * 100), milli);
    total = { net: exact(found.net), vat: exact(found.vat), gross: exact(found.gross), rate: found.rate };
    if (found.included) total.included = true; // НДС уже внутри итога («в том числе НДС»)
    total.extraDigits = hasExtraDigits(total.gross);
  }
  return { task, total, version: versionFromTexts(lines) ?? versionFromTexts(texts) };
}

/** 110929.15 → «110 929,15»; доли копейки не прячем: 110929.155 → «110 929,155» */
export function formatMoney(x) {
  const digits = hasExtraDigits(x) ? 3 : 2;
  return x.toLocaleString('ru-RU', { minimumFractionDigits: digits, maximumFractionDigits: digits });
}

/**
 * Одно число, как его пишут люди и подрядчики: «4.538» и «4.538₽» (точка — тысячи), «5 263,50», «21,262.50», «3000.5».
 * Точка ровно с тремя цифрами после неё — это тысячи, а не копейки: так пишет «Бюро переводов».
 */
export function parseNumber(s) {
  let t = String(s).replace(/[\s  ₽]|руб\.?|р\./gi, '');
  if (!t) return NaN;
  if (/^\d{1,3}(?:\.\d{3})+(?:,\d{1,3})?$/.test(t)) t = t.replace(/\./g, '').replace(',', '.'); // 4.538 / 1.234.567,50
  else if (/^\d{1,3}(?:,\d{3})+(?:\.\d{1,3})?$/.test(t)) t = t.replace(/,/g, ''); // 21,262.50
  else t = t.replace(',', '.');
  return /^\d+(?:\.\d+)?$/.test(t) ? Number(t) : NaN;
}

const round3 = (x) => Math.round(x * 1000) / 1000;

/**
 * Сумма из поля «Итого»: «5263+5336», «=3000+2500,50» или кусок письма
 * «английский — 5.263₽, китайский — 5.336₽» (берём все суммы с ₽ / руб).
 * → { parts: [5263, 5336], total: 10599, formula: '5263+5336' } · пусто → null · непонятно → { error }.
 */
export function parseSum(input) {
  const s = String(input ?? '').trim().replace(/^=/, '');
  if (!s) return null;
  let pieces;
  const withRub = [...s.matchAll(/(\d[\d\s  .,]*?)\s*(?:₽|руб|р\.)/gi)].map((m) => m[1]);
  if (/[а-яё]/i.test(s.replace(/руб\.?|р\./gi, '')) || (withRub.length > 1 && !s.includes('+'))) {
    if (!withRub.length) return { error: 'Не нашла сумм с ₽' };
    pieces = withRub;
  } else {
    pieces = s.split('+');
  }
  const parts = pieces.map(parseNumber);
  if (parts.some((p) => !isFinite(p))) return { error: 'Сумма не похожа на число' };
  const total = round3(parts.reduce((a, b) => a + b, 0));
  return { parts, total, formula: parts.length > 1 ? parts.map((p) => String(round3(p))).join('+') : '' };
}

// ---------- Акт / счёт подрядчика за месяц ----------
// Кириллица, похожая на латиницу (в PDF «LIT» иногда набрано русскими буквами)
const LOOKALIKE = { А: 'A', В: 'B', Е: 'E', К: 'K', М: 'M', Н: 'H', О: 'O', Р: 'P', С: 'C', Т: 'T', Х: 'X', У: 'Y', а: 'a', е: 'e', о: 'o', р: 'p', с: 'c', х: 'x', у: 'y' };
const ACT_ID_RE = /(?<![A-Za-zА-Яа-яЁё])([A-Za-zА-Яа-яЁё]{2,6})\s?[-–—]\s?(\d{2})\s?[-–—]\s?(\d{1,6})(?!\d)/;
const ACT_SKIP_RE = /итого|всего|в\s+том\s+числе|к\s+оплате|сумма\s+ндс/i;

function actId(line) {
  const m = String(line).match(ACT_ID_RE);
  if (!m) return null;
  const pref = m[1].replace(/[А-Яа-яЁё]/g, (c) => LOOKALIKE[c] || c);
  if (/[^A-Za-z]/.test(pref)) return null; // настоящее русское слово, а не номер
  return `${pref}-${m[2]}-${m[3]}`;
}
const lineMilli = (line) => [...String(line).matchAll(MONEY_RE)].map((m) => toMilli(m[0]));

/**
 * Сумма с НДС из строки акта: самая большая сумма, для которой в строке есть «без НДС» (с НДС = без НДС × 1,2 и т. п.)
 * или «без НДС + НДС». Так количество «1» перед ценой «386,00» (читается как 1 386,00) не собьёт итог.
 * Пары нет (акт без НДС) — самая большая сумма.
 */
function grossOf(milli) {
  const c = milli.map((m) => Math.round(m / 10));
  let best = null;
  for (const g of c) {
    const ok = c.some((n) => n < g && (VAT_RATES.some((r) => Math.abs(n * (1 + r) - g) <= 2) || c.some((v) => v < g && Math.abs(n + v - g) <= 1 && VAT_RATES.some((r) => Math.abs(n * r - v) <= 2))));
    if (ok && (best == null || g > best)) best = g;
  }
  return best != null ? milli[c.indexOf(best)] : Math.max(...milli);
}

/**
 * Номер, разорванный переносом в узкой колонке: «Перевод LIT-26-» и строкой ниже в той же колонке «2240 китайский».
 * Склеиваем по координатам (items из PDF: { str, x, y, page }), иначе между частями влезет «1» из колонки «№».
 */
export function repairIdBreaks(items) {
  const out = items.map((it) => ({ ...it }));
  for (const it of out) {
    if (!/[A-Za-zА-Яа-я]{2,6}-(?:\d{2}-)?\s*$/.test(it.str)) continue;
    let next = null;
    for (const o of out) {
      if (o === it || o.page !== it.page || o.y >= it.y - 1 || Math.abs(o.x - it.x) > 4 || !/^\s*\d/.test(o.str)) continue;
      if (!next || o.y > next.y) next = o;
    }
    if (!next || it.y - next.y > 30) continue;
    const m = next.str.match(/^\s*(\d+(?:-\d+)?)/);
    it.str = it.str.replace(/\s+$/, '') + m[1];
    next.str = next.str.slice(m[0].length);
  }
  return out.filter((it) => it.str.trim());
}

/**
 * Акт / счёт за месяц: строки с номером задачи (LIT-26-2240) → сумма С НДС по каждой строке акта,
 * строки одного номера складываются. В строке акта обычно «без НДС … НДС … с НДС» — берём с НДС.
 * В узкой колонке название переносится («Перевод LIT-26-» / «2240 китайский»), а суммы стоят на одной из строк.
 * Поэтому строка акта = строка с суммами + текст над и под ней (до соседней строки с суммами).
 * → { items: [{ id, gross, lines: [{ text, gross }] }], total, unread: [что не разобрали — показать человеку] }
 */
export function parseAct(lines) {
  const L = lines.map(String);
  const anchors = [];
  L.forEach((l, i) => { if (lineMilli(l).length) anchors.push(i); });
  const blocks = new Map(anchors.map((a) => [a, []]));
  const orphans = [], ties = [];
  L.forEach((l, i) => {
    if (blocks.has(i)) { blocks.get(i).push(i); return; }
    const up = anchors.filter((a) => a < i).pop(), down = anchors.find((a) => a > i);
    const du = up == null ? Infinity : i - up, dd = down == null ? Infinity : down - i;
    if (Math.min(du, dd) > 4) { if (actId(l)) orphans.push(l); return; }
    if (du === dd) ties.push([i, up, down]);
    else blocks.get(du < dd ? up : down).push(i);
  });
  // Ровно посередине: если у верхней строки номер уже есть — это начало следующей
  const hasId = (a) => blocks.get(a).some((j) => ACT_ID_RE.test(L[j]));
  ties.forEach(([i, up, down]) => blocks.get(hasId(up) ? down : up).push(i));
  const byId = new Map();
  const unread = [];
  const add = (id, a, text) => {
    const gross = grossOf(lineMilli(L[a]));
    const g = byId.get(id) || { id, milli: 0, lines: [] };
    g.milli += gross;
    g.lines.push({ text, gross: gross / 1000 });
    byId.set(id, g);
  };
  const finish = () => {
    orphans.forEach((l) => unread.push(`номер без суммы: ${l}`));
    const groups = [...byId.values()];
    const items = groups.map((g) => ({ id: g.id, gross: g.milli / 1000, lines: g.lines }));
    return { items, total: groups.reduce((a, g) => a + g.milli, 0) / 1000, unread };
  };

  // 1) Главный способ: номера и строки с суммами идут в одном порядке — k-й номер ↔ k-я строка с суммами.
  //    Не важно, где в ячейке стоит текст (сверху, по центру), и сколько строк занимает название.
  const occ = [];
  L.forEach((l, i) => {
    if (ACT_SKIP_RE.test(l)) return;
    [...new Set([...l.matchAll(new RegExp(ACT_ID_RE.source, 'g'))].map((m) => actId(m[0])).filter(Boolean))].forEach((id) => occ.push({ i, id }));
  });
  if (occ.length) {
    const lo = occ[0].i - 4, hi = occ[occ.length - 1].i + 4;
    const rowsA = anchors.filter((a) => a >= lo && a <= hi && !ACT_SKIP_RE.test(L[a]));
    if (rowsA.length === occ.length && occ.every((o, k) => Math.abs(o.i - rowsA[k]) <= 6)) {
      occ.forEach((o, k) => add(o.id, rowsA[k], o.i === rowsA[k] ? L[o.i].trim() : `${L[Math.min(o.i, rowsA[k])].trim()} ${L[Math.max(o.i, rowsA[k])].trim()}`));
      return finish();
    }
  }

  // 2) Запасной: строка акта = строка с суммами + ближайший текст над и под ней
  for (const a of anchors) {
    if (ACT_SKIP_RE.test(L[a])) continue;
    const text = blocks.get(a).sort((x, y) => x - y).map((i) => L[i].trim()).join(' ');
    const ids = [...new Set([...text.matchAll(new RegExp(ACT_ID_RE.source, 'g'))].map((m) => actId(m[0])).filter(Boolean))];
    if (!ids.length) { if (/перев|редакт|локализ|услуг|работ/i.test(text)) unread.push(`без номера: ${text}`); continue; }
    if (ids.length > 1) { unread.push(`несколько номеров в одной строке: ${text}`); continue; }
    add(ids[0], a, text);
  }
  return finish();
}
