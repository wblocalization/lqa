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
  },
}];
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
export async function createIssue(board, title, description) {
  const body = {
    executorId: null,
    issue: { initId: crypto.randomUUID(), title, description },
    fields: board.fields || {},
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
  const key = await createIssue(plan.board, plan.title, plan.description);
  const main = { key, url: issueUrl(plan.board, key) };
  const subs = [];
  for (const s of plan.subs || []) {
    const desc = `Основная задача: [${main.key}](${main.url})\n\n${s.description || ''}`.trim();
    const k = await createIssue(plan.board, s.title, desc);
    subs.push({ key: k, url: issueUrl(plan.board, k), lang: s.lang || '' });
  }
  return { main, subs };
}
