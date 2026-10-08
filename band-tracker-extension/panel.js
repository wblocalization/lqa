// Боковая панель «Band → Трекер»: форма новой задачи и список досок.
// Трекер пускает запросы только со своей страницы (вход + csrf_token из cookie), поэтому задачу создаём
// во вкладке tracker.wb.ru — от имени того, кто в ней вошёл. Нет открытой вкладки — откроем в фоне.

const TRACKER = 'https://tracker.wb.ru';
// Тип задачи, который есть у всех сразу (Редакция / Локализация / Локализация). Остальные расширение запоминает само (tracker-hook.js).
const PRESETS = [{
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
const DRAFT_TTL = 10 * 60 * 1000; // черновик из Band старше 10 минут уже не подхватываем

const $ = (s) => document.querySelector(s);
const els = {
  intro: $('#intro'), helpBtn: $('#helpBtn'), form: $('#form'), from: $('#from'), board: $('#board'), title: $('#title'), desc: $('#desc'),
  submit: $('#submit'), clear: $('#clear'), done: $('#done'), doneKey: $('#doneKey'), doneTitle: $('#doneTitle'),
  copyLink: $('#copyLink'), copyKey: $('#copyKey'), again: $('#again'),
  boards: $('#boards'), boardsBtn: $('#boardsBtn'), boardList: $('#boardList'), boardsClose: $('#boardsClose'),
  msg: $('#msg'), toast: $('#toast'),
};

let boards = [];     // типы задач (адрес «пространство/проект/тип»): запомненные + встроенные
let source = null;   // { link, author } — сообщение Band, из которого пришли
let created = null;  // { key, url }

// ---------- Мелочи ----------
function setMsg(text, kind = 'info') {
  els.msg.textContent = text;
  els.msg.className = `status ${kind}`;
}
let toastTimer;
function toast(text) {
  els.toast.textContent = text;
  els.toast.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { els.toast.hidden = true; }, 2200);
}
const issueUrl = (board, key) => `${TRACKER}/i/${encodeURIComponent((board && board.ws) || 'EDITORS')}/${encodeURIComponent(key)}`;

/** Название — первая строка сообщения, не длиннее 120 знаков. */
export function titleFromText(text) {
  const plain = (l) => l.replace(/\[([^\]]*)\]\([^)]*\)/g, '$1').replace(/\*\*|__|~~|`/g, '').replace(/^([-*>]|\d+\.)\s+/, '').trim();
  // строки только из @упоминаний, #тегов и ссылок — не название (у ботов так часто начинается)
  const junk = (l) => !l.replace(/@[\wА-Яа-яЁё.-]+(\s+[А-ЯЁA-Z][а-яёa-z-]+)?|#[\wА-Яа-яЁё.-]+|https?:\S+|[\s,;:🔗]/gu, '');
  const lines = String(text || '').split('\n').map(plain).filter(Boolean);
  // Сообщение бота («Раздел: …», «Задача: <ссылка на MP-9119>») — название «Маркетплейс (FBS) · MP-9119»
  const field = (name) => { const l = lines.find((x) => x.toLowerCase().startsWith(`${name}:`)); return l ? l.slice(name.length + 1).trim() : ''; };
  const section = field('раздел') || field('команда');
  const issue = (/\/issues?\/([A-Z][A-Z0-9]*-\d+)/.exec(field('задача')) || /\/browse\/([A-Z][A-Z0-9]*-\d+)/.exec(field('задача')) || [])[1];
  if (section && issue) return `${section} · ${issue}`.slice(0, 120);
  const line = lines.find((l) => !junk(l)) || lines[0] || '';
  return line.length > 120 ? `${line.slice(0, 117).replace(/\s+\S*$/, '')}…` : line;
}

/**
 * Переносы строк как в Band: в Markdown одиночный перенос склеивает строки, поэтому ставим «жёсткий» (два пробела).
 * Внутри блоков кода ``` не трогаем.
 */
export function hardBreaks(md) {
  return String(md || '').split(/(```[\s\S]*?```)/).map((part, i) => (i % 2 ? part : part.replace(/([^\n])[ \t]*\n(?=[^\n])/g, '$1  \n'))).join('');
}

/** Описание в Markdown (трекер хранит его так): текст сообщения и ссылка на него в Band. */
export function buildDescription(text, src) {
  const parts = [String(text || '').trim()];
  if (src && src.link) parts.push(`[Сообщение в Band](${src.link})${src.author ? ` — ${src.author}` : ''}`);
  return parts.filter(Boolean).join('\n\n');
}

// ---------- Типы задач ----------
async function loadBoards() {
  const { boards: saved = [], hiddenPresets = [], boardPath = '' } = await chrome.storage.local.get(['boards', 'hiddenPresets', 'boardPath']);
  boards = [...saved, ...PRESETS.filter((p) => !hiddenPresets.includes(p.path) && !saved.some((b) => b.path === p.path))];
  els.board.innerHTML = '';
  boards.forEach((b) => els.board.add(new Option(b.name, b.path)));
  if (!boards.length) els.board.add(new Option('Нет типов — см. «Типы задач» вверху', ''));
  els.board.value = boards.some((b) => b.path === boardPath) ? boardPath : (boards[0] ? boards[0].path : '');
  renderBoardList();
}
const currentBoard = () => boards.find((b) => b.path === els.board.value) || null;

function renderBoardList() {
  els.boardList.innerHTML = '';
  boards.forEach((b) => {
    const li = document.createElement('li');
    const name = document.createElement('span');
    name.className = 'bname';
    const t = document.createElement('b'); t.textContent = b.name;
    const s = document.createElement('small');
    s.textContent = [b.space, b.project, b.preset ? 'встроенный' : `запомнен ${new Date(b.at || 0).toLocaleDateString('ru-RU')}`].filter(Boolean).join(' · ');
    name.append(t, s);
    const ren = document.createElement('button');
    ren.className = 'btn small'; ren.type = 'button'; ren.textContent = '✏️'; ren.title = 'Переименовать'; ren.dataset.ren = b.path;
    const del = document.createElement('button');
    del.className = 'btn small'; del.type = 'button'; del.textContent = '✕'; del.title = 'Убрать'; del.dataset.del = b.path;
    li.append(name, ren, del);
    els.boardList.append(li);
  });
}

els.boardList.addEventListener('click', async (e) => {
  const btn = e.target.closest('button');
  if (!btn) return;
  const { boards: saved = [], hiddenPresets = [] } = await chrome.storage.local.get(['boards', 'hiddenPresets']);
  if (btn.dataset.ren) {
    const b = boards.find((x) => x.path === btn.dataset.ren);
    const name = (prompt('Как назвать этот тип задачи?', b.name) || '').trim();
    if (!name) return;
    const rest = saved.filter((x) => x.path !== b.path);
    await chrome.storage.local.set({ boards: [...rest, { ...b, name, preset: false, at: b.at || Date.now() }] });
  } else if (btn.dataset.del) {
    const b = boards.find((x) => x.path === btn.dataset.del);
    if (!confirm(`Убрать «${b.name}»? Вернуть: создать задачу этого типа в tracker.wb.ru.`)) return;
    const patch = { boards: saved.filter((x) => x.path !== b.path) };
    if (PRESETS.some((p) => p.path === b.path)) patch.hiddenPresets = [...new Set([...hiddenPresets, b.path])];
    await chrome.storage.local.set(patch);
  }
  await loadBoards();
});

els.board.addEventListener('change', () => chrome.storage.local.set({ boardPath: els.board.value }));
function showBoards(on) {
  els.boards.hidden = !on;
  document.body.classList.toggle('boards-mode', on);
  els.boardsBtn.setAttribute('aria-pressed', String(on));
  window.scrollTo(0, 0);
}
els.boardsBtn.addEventListener('click', () => showBoards(els.boards.hidden));
els.boardsClose.addEventListener('click', () => showBoards(false));
els.helpBtn.addEventListener('click', () => {
  showBoards(false);
  els.intro.hidden = !els.intro.hidden;
  window.scrollTo(0, 0);
});

// ---------- Форма ----------
/** Открыть форму: пустую или из черновика Band { selection, text, link, author }. */
function openForm(draft = null) {
  created = null;
  source = draft && (draft.link || draft.author) ? { link: draft.link || '', author: draft.author || '' } : null;
  const text = draft ? (draft.selection || draft.text || '') : '';
  els.title.value = titleFromText(text);
  els.desc.value = buildDescription(text, source);
  els.from.hidden = !source;
  els.from.textContent = '';
  if (source) {
    els.from.append(`Из Band${source.author ? ` · ${source.author}` : ''} · `);
    const a = document.createElement('a');
    a.target = '_blank'; a.rel = 'noopener';
    a.href = source.link || '#';
    a.textContent = source.link ? 'сообщение ↗' : 'канал';
    els.from.append(a);
  }
  showBoards(false);
  els.form.hidden = false;
  els.done.hidden = true;
  setMsg(draft && !text ? 'Текст сообщения не нашёлся — впишите название сами. Если вкладку Band открыли до установки расширения, обновите её (F5).' : '');
  els.intro.hidden = !!source || Boolean(draft && text); // пришли из Band — подсказка уже не нужна
  (els.title.value ? els.desc : els.title).focus();
}

/** Свежий черновик из Band — забрать и стереть, чтобы не открылся второй раз. */
async function pickDraft() {
  const { trackerDraft: d } = await chrome.storage.local.get('trackerDraft');
  if (!d) return;
  await chrome.storage.local.remove('trackerDraft');
  if (Date.now() - (d.at || 0) < DRAFT_TTL) openForm(d);
}

// ---------- Вкладка трекера ----------
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
async function createIssue(board, title, description) {
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

// ---------- Кнопки ----------
els.form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const board = currentBoard();
  const title = els.title.value.trim();
  if (!board) return setMsg('Нет типа задачи: создайте одну задачу нужного типа в tracker.wb.ru — он появится в списке', 'err');
  if (!title) return setMsg('Впишите название задачи', 'err');
  els.submit.disabled = true;
  els.submit.textContent = 'Создаю…';
  setMsg('');
  try {
    const key = await createIssue(board, title, hardBreaks(els.desc.value.trim()));
    created = { key, url: issueUrl(board, key) };
    els.doneKey.textContent = key;
    els.doneKey.href = created.url;
    els.doneTitle.textContent = title;
    els.form.hidden = true;
    els.done.hidden = false;
    els.intro.hidden = true;
  } catch (err) {
    setMsg(err.message, 'err');
  } finally {
    els.submit.disabled = false;
    els.submit.textContent = 'Создать в трекере';
  }
});

els.copyLink.addEventListener('click', async () => {
  if (!created) return;
  await navigator.clipboard.writeText(created.url);
  toast('✓ Ссылка скопирована');
});
els.copyKey.addEventListener('click', async () => {
  if (!created) return;
  await navigator.clipboard.writeText(created.key);
  toast('✓ Номер скопирован');
});
els.again.addEventListener('click', () => openForm());
els.clear.addEventListener('click', () => openForm());

chrome.storage.onChanged.addListener(async (changes) => {
  if (changes.trackerDraft && changes.trackerDraft.newValue) pickDraft();
  if (changes.boardsAdded && changes.boardsAdded.newValue) {
    await loadBoards();
    const b = boards.find((x) => x.path === changes.boardsAdded.newValue.path);
    if (b) { els.board.value = b.path; chrome.storage.local.set({ boardPath: b.path }); toast(`📌 Запомнила тип: ${b.name}`); }
  }
});

// ---------- Старт ----------
await loadBoards();
openForm();
pickDraft(); // панель открылась по правому клику в Band — черновик уже ждёт
