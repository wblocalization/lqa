// Разбор сообщения-заказа из Band: продукт, языки (с кодами и странами), ссылки — и название задачи по сути.
// Модуль без зависимостей от браузера: его используют панель, фон и тесты в Node.

// Язык → код и страна (страна — для поля «Страны» и для описания задачи)
export const LANGS = {
  английский: ['en', 'Международный'], грузинский: ['ka', 'Грузия'], казахский: ['kk', 'Казахстан'], кыргызский: ['ky', 'Кыргызстан'],
  киргизский: ['ky', 'Кыргызстан'], узбекский: ['uz', 'Узбекистан'], таджикский: ['tg', 'Таджикистан'], армянский: ['hy', 'Армения'],
  азербайджанский: ['az', 'Азербайджан'], амхарский: ['am', 'Эфиопия'], белорусский: ['be', 'Беларусь'], китайский: ['zh', 'Китай'],
  турецкий: ['tr', 'Турция'], иврит: ['he', 'Израиль'], арабский: ['ar', 'ОАЭ'], французский: ['fr', 'Франция'], испанский: ['es', 'Испания'],
  немецкий: ['de', 'Германия'], корейский: ['ko', 'Корея'], японский: ['ja', 'Япония'], суахили: ['sw', 'Кения'], русский: ['ru', 'Россия'],
  португальский: ['pt', 'Португалия'], итальянский: ['it', 'Италия'], вьетнамский: ['vi', 'Вьетнам'], монгольский: ['mn', 'Монголия'],
  польский: ['pl', 'Польша'], румынский: ['ro', 'Молдова'], сербский: ['sr', 'Сербия'], хинди: ['hi', 'Индия'], индонезийский: ['id', 'Индонезия'],
};
const BY_CODE = Object.fromEntries(Object.entries(LANGS).map(([name, [code, country]]) => [code, { name, code, country }]));
// «kk_cyr», «kk-Latn» → kk
const baseCode = (c) => String(c).toLowerCase().split(/[-_]/)[0];

const plain = (l) => String(l).replace(/\[([^\]]*)\]\([^)]*\)/g, '$1').replace(/\*\*|__|~~|`/g, '').trim();

/** Значение поля «Языки: …», «Продукт: …» из сообщения (строки вида «Название: значение»). */
function field(lines, ...names) {
  for (const l of lines) {
    const m = l.match(/^\s*[-*•]?\s*([А-Яа-яЁёA-Za-z ]{2,40}?)\s*:\s*(.+)$/);
    if (m && names.includes(m[1].trim().toLowerCase())) return m[2].trim();
  }
  return '';
}

/** «английский (en), грузинский, kk» → [{ name, code, country }] без повторов. */
export function parseLanguages(s) {
  const out = [];
  const add = (x) => { if (x && !out.some((o) => o.code === x.code)) out.push(x); };
  for (const part of String(s || '').split(/[,;/]|\sи\s/)) {
    const t = part.toLowerCase().replace(/[.()]/g, ' ').trim();
    if (!t) continue;
    const name = Object.keys(LANGS).find((n) => t.includes(n));
    if (name) { add({ name, code: LANGS[name][0], country: LANGS[name][1] }); continue; }
    const code = (t.match(/\b([a-z]{2}(?:[-_][a-z]{2,4})?)\b/) || [])[1];
    if (code && BY_CODE[baseCode(code)]) add(BY_CODE[baseCode(code)]);
  }
  return out;
}

const GREETING = /^(коллеги|привет|добрый|здравствуй|доброе|всем привет|hi\b|hello\b)/i;

/**
 * Сообщение-заказ → { product, languages, links, verb, title, text }.
 * Название: «Tracker: переводы — en, ka, kk». Нет полей — первая содержательная строка (без «Коллеги, добрый день!»).
 */
export function parseOrder(md) {
  const text = String(md || '');
  const lines = text.split('\n').map(plain).filter(Boolean);
  const product = field(lines, 'продукт', 'сервис', 'проект').replace(/[.;]+$/, '');
  const languages = parseLanguages(field(lines, 'языки', 'язык', 'languages'));
  const links = [...text.matchAll(/\[([^\]]+)\]\((https?:[^)\s]+)\)|(https?:\/\/[^\s)]+)/g)]
    .map((m) => ({ label: m[1] || '', url: m[2] || m[3] }))
    .filter((l, i, a) => a.findIndex((x) => x.url === l.url) === i);
  const body = lines.filter((l) => !GREETING.test(l) && !/^[А-Яа-яЁёA-Za-z ]{2,40}:/.test(l)).join(' ');
  const verb = /провер|вычит|ревью|lqa/i.test(body) && !/перев[оеё]д|перевест/i.test(body) ? 'проверка'
    : /перев[оеё]д|перевест/i.test(body) ? 'переводы' : 'локализация';
  const due = parseDue(field(lines, 'дедлайн', 'срок', 'срок сдачи', 'deadline', 'дедлайн (если есть)'));
  const codes = languages.map((l) => l.code).join(', ');
  let title = '';
  if (product) title = `${product}: ${verb}${codes ? ` — ${codes}` : ''}`;
  else {
    const first = lines.find((l) => !GREETING.test(l) && !/^[@#]/.test(l)) || lines[0] || '';
    title = first.replace(/^прошу\s+/i, '').replace(/^(взять в работу|сделать)\s+/i, '');
    title = title.charAt(0).toUpperCase() + title.slice(1);
    if (codes && !title.includes(codes)) title += ` — ${codes}`;
  }
  if (title.length > 120) title = `${title.slice(0, 117).replace(/\s+\S*$/, '')}…`;
  return { product, languages, links, verb, title, text, due };
}

/**
 * «28.10», «28.10.2026», «до 28.10.26» → '2026-10-28T00:00:00+03:00' (так трекер хранит дедлайн). Нет даты — ''.
 * Год не указан — ближайшая такая дата не в прошлом.
 */
export function parseDue(s, now = new Date()) {
  const m = String(s || '').match(/(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?/);
  if (!m) return '';
  const d = Number(m[1]), mo = Number(m[2]);
  if (d < 1 || d > 31 || mo < 1 || mo > 12) return '';
  let y = m[3] ? Number(m[3]) : now.getFullYear();
  if (y < 100) y += 2000;
  if (!m[3] && new Date(y, mo - 1, d) < new Date(now.getFullYear(), now.getMonth(), now.getDate())) y += 1;
  const p2 = (n) => String(n).padStart(2, '0');
  return `${y}-${p2(mo)}-${p2(d)}T00:00:00+03:00`;
}

/** Название подзадачи для одного языка: «Tracker: перевод — ka». */
export function subtaskTitle(order, lang) {
  const base = order.product || order.title.replace(/\s+—\s+[a-z, _-]+$/i, '');
  const verb = order.verb === 'проверка' ? 'проверка' : 'перевод';
  return `${base}: ${verb} — ${lang.code}`;
}

/** Блок «что из заказа» для описания задачи: продукт, языки, страны, ссылки. */
export function orderSummary(order) {
  const rows = [];
  if (order.product) rows.push(`**Продукт / БЗ:** ${order.product}`);
  if (order.languages.length) {
    rows.push(`**Языки:** ${order.languages.map((l) => `${l.name} (${l.code})`).join(', ')}`);
    const countries = [...new Set(order.languages.map((l) => l.country))];
    rows.push(`**Страны:** ${countries.join(', ')}`);
  }
  if (order.links.length) rows.push(`**Ссылки:** ${order.links.map((l) => (l.label ? `[${l.label}](${l.url})` : l.url)).join(' · ')}`);
  return rows.join('\n');
}

/** Номера задач трекера в тексте (ответ в треде «Завели: LOC-101…» или ссылки tracker.wb.ru/i/…/LOC-101). */
export function trackerKeys(text) {
  const keys = new Set();
  for (const m of String(text || '').matchAll(/tracker\.wb\.ru\/i\/[^/\s)]+\/([A-Z][A-Z0-9_]*-\d+)/g)) keys.add(m[1]);
  return [...keys];
}
