// Создание задач в tracker.wb.ru — общее для панели и фона (карточка в треде Band).
// Трекер пускает запросы только со своей страницы (вход + csrf_token из cookie), поэтому задачу создаём
// во вкладке tracker.wb.ru — от имени того, кто в ней вошёл. Нет открытой вкладки — откроем в фоне.

export const TRACKER = 'https://tracker.wb.ru';
// Тип задачи, который есть у всех сразу (Редакция / Локализация / Локализация). Остальные расширение запоминает само (tracker-hook.js).
export const PRESETS = [{
  path: '/api/gateway/catalog/redaktsiya/lokalizatsiya/lokalizatsiya',
  name: 'Локализация · Локализация', space: 'Редакция', ws: 'EDITORS', project: 'LOCAL', preset: true,
  fields: {
    '63f946ab-d7d0-4b5f-a0d6-a1f5982c3448': [],
    '1ee7026d-8aae-4bb1-bf9c-3ebc9ea3a7bc': [{ id: 'DgqUHf16il6y12Jms6LPN', label: 'Normal' }],
    'b5fd258b-fe14-45d0-b877-f38c585f6e38': [],
    '7f54c34c-2992-4817-b86c-03310905c8d1': [],
    '883735ab-be95-4d44-a830-62740ae23fcb': [],
    'c907e2d8-c109-4147-ae67-f78fbdc3bdfb': [],
  },
}];

// Поля трекера, которые заполняем по заказу (коды — из «Скопировать для разработчика», тип «Локализация»).
// Остальные поля берём из запомненного типа как есть (приоритет и т. п.).
export const ROLE = {
  due: '63f946ab-d7d0-4b5f-a0d6-a1f5982c3448',       // дедлайн: ['2026-10-28T00:00:00+03:00']
  bz: '883735ab-be95-4d44-a830-62740ae23fcb',        // БЗ: один вариант [{ id, label }]
  countries: 'c907e2d8-c109-4147-ae67-f78fbdc3bdfb', // страны: несколько [{ id, label }]
};
// Как страна может называться в трекере
const COUNTRY_ALIASES = { кыргызстан: ['киргизия', 'киргизстан'], беларусь: ['белоруссия'], молдова: ['молдавия'], оаэ: ['объединённые арабские эмираты'] };

/** Варианты полей, которые расширение уже видело в трекере: { [поле]: { [название в нижнем регистре]: { id, label } } }. */
// Что уже известно из трекера (от первой задачи, созданной руками) — у всех сразу; остальное расширение выучит само
export const SEED_OPTIONS = {
  'c907e2d8-c109-4147-ae67-f78fbdc3bdfb': { беларусь: { id: 'PQYSAXgqcl3w8Gxcoas1s', label: 'Беларусь' }, грузия: { id: '_3TyyGB7ox9s9H5-KNumf', label: 'Грузия' } },
  '883735ab-be95-4d44-a830-62740ae23fcb': { межнар: { id: 'kp9u0p5qrqk60pkUJfeoG', label: 'Межнар' } },
};
export async function knownOptions() {
  const { fieldOptions = {} } = await chrome.storage.local.get('fieldOptions');
  const out = JSON.parse(JSON.stringify(SEED_OPTIONS));
  for (const [fid, o] of Object.entries(fieldOptions)) out[fid] = { ...(out[fid] || {}), ...o };
  return out;
}
const norm = (s) => String(s || '').trim().toLowerCase().replace(/ё/g, 'е');

/**
 * Поля задачи по заказу: страны (по языкам), БЗ (по продукту), дедлайн. Чего трекер ещё не «показал» — оставляем пустым
 * и сообщаем в missing, чтобы человек видел. → { fields, missing: { countries: [...], bz: '' }, picked: { countries: [...], bz } }
 */
export async function fieldsFor(board, { countries = [], bz = '', due = '' } = {}) {
  const opts = await knownOptions();
  const f = JSON.parse(JSON.stringify((board && board.fields) || {}));
  const find = (fid, label) => {
    const o = opts[fid] || {};
    const n = norm(label);
    if (o[n]) return o[n];
    for (const alias of COUNTRY_ALIASES[n] || []) if (o[norm(alias)]) return o[norm(alias)];
    return null;
  };
  const missing = { countries: [], bz: '' };
  const picked = { countries: [], bz: '' };
  if (ROLE.countries in f) {
    const real = countries.filter((c) => norm(c) !== 'международный');
    f[ROLE.countries] = real.map((c) => { const o = find(ROLE.countries, c); if (!o) missing.countries.push(c); return o; }).filter(Boolean).map(({ id, label }) => ({ id, label }));
    picked.countries = f[ROLE.countries].map((o) => o.label);
  }
  if (ROLE.bz in f) {
    let o = bz ? find(ROLE.bz, bz) : null;
    // «Tracker» ↔ «Tracker (трекер задач)» — похожее название тоже подходит
    if (!o && bz) o = Object.values(opts[ROLE.bz] || {}).find((x) => norm(x.label).includes(norm(bz)) || norm(bz).includes(norm(x.label))) || null;
    f[ROLE.bz] = o ? [{ id: o.id, label: o.label }] : [];
    if (bz && !o) missing.bz = bz;
    picked.bz = o ? o.label : '';
  }
  if (ROLE.due in f) f[ROLE.due] = due ? [due] : [];
  return { fields: f, missing, picked };
}
export const issueUrl = (board, key) => `${TRACKER}/i/${encodeURIComponent((board && board.ws) || 'EDITORS')}/${encodeURIComponent(key)}`;

/** Типы задач: запомненные + встроенные; выбранный — первым. */
export async function loadBoards() {
  const { boards: saved = [], hiddenPresets = [], boardPath = '' } = await chrome.storage.local.get(['boards', 'hiddenPresets', 'boardPath']);
  const all = [...saved, ...PRESETS.filter((p) => !hiddenPresets.includes(p.path) && !saved.some((b) => b.path === p.path))];
  return { boards: all, current: all.find((b) => b.path === boardPath) || all[0] || null };
}

function waitLoaded(tabId, timeoutMs = 30000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(onUpdated);
      reject(new Error('Трекер не открылся за 30 секунд — проверьте сеть/VPN'));
    }, timeoutMs);
    function onUpdated(id, info) {
      if (id === tabId && info.status === 'complete') {
        clearTimeout(timer);
        chrome.tabs.onUpdated.removeListener(onUpdated);
        resolve();
      }
    }
    chrome.tabs.onUpdated.addListener(onUpdated);
  });
}

async function trackerTab() {
  const [tab] = await chrome.tabs.query({ url: `${TRACKER}/*` });
  if (tab && tab.status === 'complete') return tab;
  if (tab) { await waitLoaded(tab.id); return tab; }
  const fresh = await chrome.tabs.create({ url: `${TRACKER}/`, active: false });
  await waitLoaded(fresh.id);
  return fresh;
}

/** Создать задачу этого типа. Возвращает номер, например «LOCAL-1817». */
export async function createIssue(board, title, description, fields) {
  const body = {
    executorId: null,
    issue: { initId: crypto.randomUUID(), title, description },
    fields: fields || board.fields || {},
    files: { add: [], remove: [] },
    agile: { sprintIds: [] },
  };
  const tab = await trackerTab();
  let res;
  try {
    [res] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, world: 'MAIN', func: createInPage, args: [board.path, body] });
  } catch (e) {
    if (/error page|Cannot access/i.test(e.message)) throw new Error('Трекер не открылся во вкладке — проверьте сеть/VPN и что tracker.wb.ru открывается');
    throw e;
  }
  const r = res && res.result;
  if (!r) throw new Error('Трекер не ответил');
  if (r.error) throw new Error(r.error);
  return r.key;
}

/** Выполняется на странице tracker.wb.ru: там вход и csrf_token. */
async function createInPage(path, body) {
  const m = document.cookie.match(/(?:^|;\s*)csrf_token=([^;]*)/);
  if (!m) return { error: 'Не вижу входа в трекер — откройте tracker.wb.ru, войдите и попробуйте ещё раз' };
  try {
    const r = await fetch(path, {
      method: 'POST',
      credentials: 'include',
      headers: { 'content-type': 'application/json', accept: 'application/json', 'x-csrf-token': decodeURIComponent(m[1]) },
      body: JSON.stringify(body),
      bandTrackerOwn: true, // метка для tracker-hook.js: эту задачу создали мы, тип не перезаписывать
    });
    let j = {};
    try { j = await r.json(); } catch { /* не JSON — ниже скажем код */ }
    if (r.status === 401 || r.status === 403) return { error: `Трекер не пустил (код ${r.status}) — войдите в tracker.wb.ru заново` };
    if (!r.ok || !j.success || typeof j.data !== 'string') {
      const why = j.message || j.error || (j.errors && JSON.stringify(j.errors)) || '';
      return { error: `Трекер не создал задачу (код ${r.status})${why ? `: ${why}` : ''}` };
    }
    return { key: j.data };
  } catch (e) {
    return { error: `Нет связи с трекером: ${e.message}` };
  }
}


/** Журнал: что, когда и по какому сообщению создано (последние 200) — для «уже есть» и разбора ошибок. */
export async function logCreated(entry) {
  const { journal = [] } = await chrome.storage.local.get('journal');
  await chrome.storage.local.set({ journal: [{ ...entry, at: Date.now() }, ...journal].slice(0, 200) });
}
export async function journalFor(postId) {
  if (!postId) return null;
  const { journal = [] } = await chrome.storage.local.get('journal');
  return journal.find((j) => j.postId === postId) || null;
}

/**
 * Основная задача + подзадачи по языкам. plan: { board, title, description, subs: [{ title, description }] }.
 * Подзадачи — отдельные задачи того же типа со ссылкой на основную (как подзадачи в трекере свяжем, когда узнаем, как он их создаёт).
 * Возвращает { main: { key, url }, subs: [{ key, url, lang }] }.
 */
export async function createPlan(plan) {
  const key = await createIssue(plan.board, plan.title, plan.description, plan.fields);
  const main = { key, url: issueUrl(plan.board, key) };
  const subs = [];
  for (const s of plan.subs || []) {
    const desc = `Основная задача: [${main.key}](${main.url})\n\n${s.description || ''}`.trim();
    const k = await createIssue(plan.board, s.title, desc, s.fields || plan.fields);
    subs.push({ key: k, url: issueUrl(plan.board, k), lang: s.lang || '' });
  }
  return { main, subs };
}
