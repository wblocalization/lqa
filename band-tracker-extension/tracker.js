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
  // «Локализация · Тип задачи» — подпись формы попала в название по ошибке (2.0): убираем
  const clean = (b) => ({ ...b, name: String(b.name || '').split(' · ').filter((x) => !/^(тип задачи|проект|пространство)$/i.test(x.trim())).join(' · ') || b.name });
  const all = [...saved.map(clean), ...PRESETS.filter((p) => !hiddenPresets.includes(p.path) && !saved.some((b) => b.path === p.path))];
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

// ---------- Справочники трекера: варианты БЗ, стран и исполнители ----------
// Трекер отдаёт форму создания задачи и список исполнителей по адресам (видно в «Скопировать для разработчика»):
//   /api/gateway/entrypoint/<пространство>/<проект>/<тип>?categoryID=…  — поля формы с вариантами
//   /api/gateway/namespace/<ws>/entrypoint_group/<проект>/entrypoint/<тип>/executors — исполнители
// Названия ключей у вариантов заранее не знаем, поэтому ищем похожее: id|value|key|uuid + label|name|title|fullName…
const OPT_ID = ['id', 'value', 'key', 'uuid', 'code', 'userId', 'user_id'];
const OPT_LABEL = ['label', 'name', 'title', 'fullName', 'full_name', 'displayName', 'text', 'value', 'login', 'email'];
function asOption(x) {
  if (!x || typeof x !== 'object' || Array.isArray(x)) return null;
  if (x.user && typeof x.user === 'object') { const u = asOption(x.user); if (u) return u; }
  const idk = OPT_ID.find((k) => typeof x[k] === 'string' && x[k]);
  const lk = OPT_LABEL.find((k) => k !== idk && typeof x[k] === 'string' && x[k].trim());
  return idk && lk ? { id: x[idk], label: x[lk].trim() } : null;
}
function optionArrays(o, depth = 0, out = []) {
  if (!o || typeof o !== 'object' || depth > 8) return out;
  if (Array.isArray(o)) {
    const opts = o.map(asOption).filter(Boolean);
    if (opts.length >= 1 && opts.length >= o.length / 2) out.push(opts);
    o.slice(0, 300).forEach((x) => optionArrays(x, depth + 1, out));
  } else Object.values(o).forEach((x) => optionArrays(x, depth + 1, out));
  return out;
}
/** Варианты для каждого из нужных полей: объект, где где-то есть id поля, а внутри — массив вариантов. */
function optionsByField(json, fieldIds) {
  const found = {};
  (function walk(o, depth) {
    if (!o || typeof o !== 'object' || depth > 10) return;
    if (!Array.isArray(o)) {
      const fid = fieldIds.find((f) => Object.values(o).some((v) => v === f));
      if (fid && !found[fid]) {
        const lists = optionArrays(o).filter((l) => !l.some((x) => x.id === fid));
        if (lists.length) found[fid] = lists.sort((a, b) => b.length - a.length)[0];
      }
    }
    (Array.isArray(o) ? o : Object.values(o)).forEach((x) => walk(x, depth + 1));
  })(json, 0);
  return found;
}
/** Устройство ответа без значений — для разработчика, если варианты не нашлись. */
function shapeOf(o, depth = 0) {
  if (depth > 6) return '…';
  if (Array.isArray(o)) return o.length ? [`${o.length}×`, shapeOf(o[0], depth + 1)] : [];
  if (o && typeof o === 'object') return Object.fromEntries(Object.entries(o).slice(0, 25).map(([k, v]) => [k, shapeOf(v, depth + 1)]));
  return typeof o;
}

/** Выполняется на странице tracker.wb.ru: GET от имени вошедшего. */
async function getJsonInPage(paths) {
  const out = {};
  for (const p of paths) {
    try {
      const r = await fetch(p, { credentials: 'include', headers: { accept: 'application/json' } });
      let j = null; try { j = await r.json(); } catch { /* не JSON */ }
      out[p] = { status: r.status, json: j };
    } catch (e) { out[p] = { status: 0, error: e.message }; }
  }
  return out;
}

/**
 * Загрузить справочники для типа задачи: варианты БЗ и стран → fieldOptions, исполнители → executors.
 * → { bz, countries, executors, report } (числа — сколько вариантов нашлось; report — для «Скопировать для разработчика»)
 */
export async function refreshDictionaries(board) {
  const m = /^\/api\/gateway\/catalog\/([^/]+)\/([^/]+)\/([^/]+)$/.exec((board && board.path) || '');
  if (!m) throw new Error('Не понимаю адрес типа задачи — пришлите «Скопировать для разработчика»');
  const [, space, project, type] = m;
  const { catalogTypes = {} } = await chrome.storage.local.get('catalogTypes');
  const cat = catalogTypes[type] || '';
  const ws = board.ws || 'EDITORS', code = board.project || 'LOCAL';
  const entry = `/api/gateway/entrypoint/${space}/${project}/${type}`;
  const paths = [entry, ...(cat ? [`${entry}?categoryID=${encodeURIComponent(cat)}`] : []), `/api/catalog/${space}/${project}`,
    `/api/gateway/namespace/${ws}/entrypoint_group/${code}/entrypoint/${type}/executors`];
  const tab = await trackerTab();
  const [res] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, world: 'MAIN', func: getJsonInPage, args: [paths] });
  let got = res && res.result || {};
  // categoryID — id типа в каталоге проекта (Локализация, Перевод, LQA…): узнали — спросим форму ещё раз
  const catalog = got[`/api/catalog/${space}/${project}`];
  if (!cat && catalog && catalog.json) {
    const types = optionArrays(catalog.json).flat();
    const typeName = (board.name || '').split(' · ')[0].toLowerCase();
    const t = types.find((x) => x.label.toLowerCase() === typeName) || types.find((x) => x.label.toLowerCase() === 'локализация');
    if (t) {
      await chrome.storage.local.set({ catalogTypes: { ...catalogTypes, [type]: t.id } });
      const [again] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, world: 'MAIN', func: getJsonInPage, args: [[`${entry}?categoryID=${encodeURIComponent(t.id)}`]] });
      got = { ...got, ...(again && again.result) };
    }
  }
  const fieldIds = [ROLE.bz, ROLE.countries];
  const found = {};
  for (const [p, r] of Object.entries(got)) {
    if (!p.includes('/entrypoint/') || p.endsWith('/executors') || !r.json) continue;
    const f = optionsByField(r.json, fieldIds);
    for (const fid of fieldIds) if (f[fid] && (!found[fid] || f[fid].length > found[fid].length)) found[fid] = f[fid];
  }
  const { fieldOptions = {} } = await chrome.storage.local.get('fieldOptions');
  for (const [fid, list] of Object.entries(found)) {
    fieldOptions[fid] = fieldOptions[fid] || {};
    for (const x of list) fieldOptions[fid][norm(x.label)] = { id: x.id, label: x.label };
  }
  const exResp = Object.entries(got).find(([p]) => p.endsWith('/executors'));
  const people = exResp && exResp[1].json ? optionArrays(exResp[1].json).sort((a, b) => b.length - a.length)[0] || [] : [];
  await chrome.storage.local.set({ fieldOptions, executors: people.slice(0, 500), dictAt: Date.now() });
  const report = Object.fromEntries(Object.entries(got).map(([p, r]) => [p, { status: r.status, error: r.error, shape: r.json ? shapeOf(r.json) : null }]));
  await chrome.storage.local.set({ dictReport: report });
  return { bz: (found[ROLE.bz] || []).length, countries: (found[ROLE.countries] || []).length, executors: people.length, report };
}
