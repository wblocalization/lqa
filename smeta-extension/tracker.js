// Задача в трекере (tracker.wb.ru) прямо из панели. Попадаем сюда двумя путями:
// правый клик по сообщению в Band → «📮 Создать задачу в трекере» или «＋ создать» у поля «Тикет в трекере».
// Трекер пускает запросы только со своей страницы (вход + csrf_token из cookie), поэтому создаём задачу
// во вкладке tracker.wb.ru — от имени того, кто в ней вошёл. Нет открытой вкладки — откроем в фоне.
import { toast } from './core.js';

export const TRACKER_ORIGIN = 'https://tracker.wb.ru';
// Доска «Редакция → Локализация» (EDITORS / LOCAL)
const CREATE_PATH = '/api/gateway/catalog/redaktsiya/lokalizatsiya/lokalizatsiya';
const BOARD_URL = `${TRACKER_ORIGIN}/n/EDITORS/p/LOCAL/b/BOARD-3715/board`;
export const issueUrl = (key) => `${TRACKER_ORIGIN}/i/EDITORS/${encodeURIComponent(key)}`;
// Поля доски — как их отправляет сама форма трекера: приоритет Normal, остальное пусто
const FIELDS = {
  '63f946ab-d7d0-4b5f-a0d6-a1f5982c3448': [],
  '1ee7026d-8aae-4bb1-bf9c-3ebc9ea3a7bc': [{ id: 'DgqUHf16il6y12Jms6LPN', label: 'Normal' }],
  'b5fd258b-fe14-45d0-b877-f38c585f6e38': [],
  '7f54c34c-2992-4817-b86c-03310905c8d1': [],
  '883735ab-be95-4d44-a830-62740ae23fcb': [],
};
const DRAFT_TTL = 10 * 60 * 1000; // черновик из Band старше 10 минут уже не подхватываем

const $ = (s) => document.querySelector(s);
const els = {
  form: $('#trForm'), from: $('#trFrom'), title: $('#trTitle'), desc: $('#trDesc'), submit: $('#trSubmit'),
  back: $('#trBack'), msg: $('#trMsg'), done: $('#trDone'), doneKey: $('#trDoneKey'), doneTitle: $('#trDoneTitle'),
  copy: $('#trCopy'), toTask: $('#trToTask'), again: $('#trAgain'),
};

let source = null;   // { link, author } — откуда пришли (сообщение Band)
let created = null;  // { key, title, link } — что создали
let hooks = {};      // { back(), toTask(info) } — задаёт app.js

export function setHooks(h) { hooks = h; }

function setMsg(text, kind = 'info') {
  els.msg.textContent = text;
  els.msg.className = `status ${kind}`;
}

/** Название — первая строка сообщения, без лишнего и не длиннее 120 знаков. */
export function titleFromText(text) {
  const line = String(text || '').split('\n').map((s) => s.trim()).find(Boolean) || '';
  return line.length > 120 ? `${line.slice(0, 117).replace(/\s+\S*$/, '')}…` : line;
}

/** Описание в Markdown (трекер хранит его так): текст сообщения и ссылка на него в Band. */
export function buildDescription(text, src) {
  const parts = [String(text || '').trim()];
  if (src && src.link) parts.push(`[Сообщение в Band](${src.link})${src.author ? ` — ${src.author}` : ''}`);
  return parts.filter(Boolean).join('\n\n');
}

/** Открыть форму: пустую или из черновика Band { text, link, author }. */
export function openForm(draft = null) {
  created = null;
  source = draft && (draft.link || draft.author) ? { link: draft.link || '', author: draft.author || '' } : null;
  const text = draft ? (draft.selection || draft.text || '') : '';
  els.title.value = titleFromText(text);
  els.desc.value = buildDescription(text, source);
  els.from.hidden = !source;
  if (source) {
    els.from.innerHTML = '';
    els.from.append('Из Band' + (source.author ? ` · ${source.author}` : '') + ' · ');
    const a = document.createElement('a');
    a.target = '_blank'; a.rel = 'noopener';
    a.href = source.link || '#';
    a.textContent = source.link ? 'сообщение ↗' : 'канал';
    els.from.append(a);
  }
  els.form.hidden = false;
  els.done.hidden = true;
  setMsg(draft && !text ? 'Текст сообщения не нашёлся — впишите название сами.' : '');
  (els.title.value ? els.desc : els.title).focus();
}

/** Свежий черновик из Band (правый клик) — забрать и стереть, чтобы не открылся второй раз. */
export async function takeDraft() {
  const { trackerDraft: d } = await chrome.storage.local.get('trackerDraft');
  if (!d) return null;
  await chrome.storage.local.remove('trackerDraft');
  return Date.now() - (d.at || 0) < DRAFT_TTL ? d : null;
}

// ---------- Вкладка трекера ----------
function waitLoaded(tabId, timeoutMs = 30000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(onUpdated);
      reject(new Error('Трекер не открылся за 30 секунд'));
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
  const [tab] = await chrome.tabs.query({ url: `${TRACKER_ORIGIN}/*` });
  if (tab && tab.status === 'complete') return tab;
  if (tab) { await waitLoaded(tab.id); return tab; }
  const fresh = await chrome.tabs.create({ url: BOARD_URL, active: false });
  await waitLoaded(fresh.id);
  return fresh;
}

/** Создать задачу. Возвращает ключ, например «LOCAL-1817». */
export async function createIssue(title, description) {
  const body = {
    executorId: null,
    issue: { initId: crypto.randomUUID(), title, description },
    fields: FIELDS,
    files: { add: [], remove: [] },
    agile: { sprintIds: [] },
  };
  const tab = await trackerTab();
  let res;
  try {
    [res] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, world: 'MAIN', func: createInPage, args: [CREATE_PATH, body] });
  } catch (e) {
    if (/error page|Cannot access/i.test(e.message)) {
      throw new Error('Трекер не открылся во вкладке — проверьте сеть/VPN и что tracker.wb.ru открывается');
    }
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
    });
    let j = {};
    try { j = await r.json(); } catch { /* не JSON — ниже скажем код */ }
    if (r.status === 401 || r.status === 403) return { error: `Трекер не пустил (код ${r.status}) — войдите в tracker.wb.ru заново` };
    if (!r.ok || !j.success || !j.data) {
      const why = j.message || j.error || (j.errors && JSON.stringify(j.errors)) || '';
      return { error: `Трекер не создал задачу (код ${r.status})${why ? `: ${why}` : ''}` };
    }
    return { key: String(j.data) };
  } catch (e) {
    return { error: `Нет связи с трекером: ${e.message}` };
  }
}

// ---------- Кнопки ----------
els.form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const title = els.title.value.trim();
  if (!title) return setMsg('Впишите название задачи', 'err');
  els.submit.disabled = true;
  els.submit.textContent = 'Создаю…';
  setMsg('');
  try {
    const key = await createIssue(title, els.desc.value.trim());
    created = { key, title, link: source ? source.link : '' };
    els.doneKey.textContent = key;
    els.doneKey.href = issueUrl(key);
    els.doneTitle.textContent = title;
    els.form.hidden = true;
    els.done.hidden = false;
    toast(`✓ Создано: ${key}`);
  } catch (err) {
    setMsg(err.message, 'err');
  } finally {
    els.submit.disabled = false;
    els.submit.textContent = 'Создать в трекере';
  }
});

els.copy.addEventListener('click', async () => {
  if (!created) return;
  await navigator.clipboard.writeText(issueUrl(created.key));
  toast('✓ Ссылка на задачу скопирована');
});
els.toTask.addEventListener('click', () => { if (created && hooks.toTask) hooks.toTask(created); });
els.again.addEventListener('click', () => openForm());
els.back.addEventListener('click', () => { if (hooks.back) hooks.back(); });
