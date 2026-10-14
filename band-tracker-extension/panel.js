// Боковая панель «Band → Трекер»: форма новой задачи и список досок.
// Трекер пускает запросы только со своей страницы (вход + csrf_token из cookie), поэтому задачу создаём
// во вкладке tracker.wb.ru — от имени того, кто в ней вошёл. Нет открытой вкладки — откроем в фоне.

import { PRESETS, issueUrl, createPlan, logCreated, journalFor, fieldsFor } from './tracker.js';
import { parseOrder, subtaskTitle, orderSummary } from './parse.js';

const DRAFT_TTL = 10 * 60 * 1000; // черновик из Band старше 10 минут уже не подхватываем

const $ = (s) => document.querySelector(s);
const els = {
  intro: $('#intro'), helpBtn: $('#helpBtn'), form: $('#form'), from: $('#from'), board: $('#board'), title: $('#title'), desc: $('#desc'),
  submit: $('#submit'), clear: $('#clear'), done: $('#done'), doneKey: $('#doneKey'), doneTitle: $('#doneTitle'),
  copyLink: $('#copyLink'), copyKey: $('#copyKey'), again: $('#again'),
  boards: $('#boards'), boardsBtn: $('#boardsBtn'), boardList: $('#boardList'), boardsClose: $('#boardsClose'),
  msg: $('#msg'), toast: $('#toast'),
  exists: $('#exists'), meta: $('#meta'), subsBox: $('#subsBox'), subs: $('#subs'), replyRow: $('#replyRow'), reply: $('#reply'),
  doneSubs: $('#doneSubs'), doneReply: $('#doneReply'), diagText: $('#diagText'), diagCopy: $('#diagCopy'), diagLists: $('#diagLists'),
};

let boards = [];     // типы задач (адрес «пространство/проект/тип»): запомненные + встроенные
let source = null;   // { link, author, postId, tabId } — сообщение Band, из которого пришли
let order = null;    // разобранный заказ (продукт, языки, ссылки)
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

/** Описание в Markdown (трекер хранит его так): что из заказа, текст сообщения и ссылка на него в Band. */
export function buildDescription(text, src, ord) {
  const parts = [ord ? orderSummary(ord) : '', String(text || '').trim()];
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
els.boardsBtn.addEventListener('click', () => { showBoards(els.boards.hidden); renderDiag(); });

// ---------- Диагностика: что расширение видело в трекере ----------
const ago = (t) => { const m = Math.round((Date.now() - t) / 60000); return m < 1 ? 'только что' : m < 60 ? `${m} мин назад` : new Date(t).toLocaleString('ru-RU'); };
async function renderDiag() {
  const { trackerSeen = [], trackerLists = [], trackerApis = [] } = await chrome.storage.local.get(['trackerSeen', 'trackerLists', 'trackerApis']);
  const last = trackerSeen[0];
  els.diagLists.textContent = `Собрано: списков — ${trackerLists.length} (${trackerLists.slice(0, 6).map((l) => `${l.count} шт.: ${l.items.slice(0, 2).map((x) => x.label).join(', ')}…`).join('; ') || 'пока нет'}), адресов трекера — ${trackerApis.length}.`;
  els.diagText.textContent = !last
    ? 'В трекере ещё не видела ни одного создания задачи. Обновите вкладку tracker.wb.ru (F5) после установки расширения и создайте задачу там.'
    : last.ok ? `Последняя задача в трекере — ${ago(last.at)}: запомнила тип.`
      : `Последний раз в трекере (${ago(last.at)}) не запомнила: ${last.why}. Нажмите «Скопировать для разработчика» и пришлите.`;
}
els.diagCopy.addEventListener('click', async () => {
  const { trackerSeen = [], boards: saved = [], trackerLists = [], trackerApis = [], fieldOptions = {} } = await chrome.storage.local.get(['trackerSeen', 'boards', 'trackerLists', 'trackerApis', 'fieldOptions']);
  const data = {
    version: chrome.runtime.getManifest().version, seen: trackerSeen,
    boards: saved.map((b) => ({ name: b.name, path: b.path, space: b.space, project: b.project, ws: b.ws, fields: b.fields })),
    fieldOptions, lists: trackerLists, apis: trackerApis,
  };
  await navigator.clipboard.writeText(JSON.stringify(data, null, 1));
  toast('✓ Скопировано — вставьте в чат разработчику');
});
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
  source = draft && (draft.link || draft.author) ? { link: draft.link || '', author: draft.author || '', postId: draft.postId || '', tabId: draft.tabId || 0 } : null;
  const text = draft ? (draft.selection || draft.text || '') : '';
  order = parseOrder(text);
  // Заказ с полями («Продукт:», «Языки:») — название по сути; иначе как раньше (первая строка / сообщение бота)
  els.title.value = (draft && draft.title) || (order.product ? order.title : titleFromText(text)) || order.title;
  els.desc.value = buildDescription(text, source, order);
  renderOrderExtras();
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

/** Подзадачи по языкам заказа, «ответить в тред», «уже есть». */
async function renderOrderExtras() {
  const langs = order ? order.languages : [];
  const countries = [...new Set(langs.map((l) => l.country))];
  els.meta.hidden = !(order && (order.product || countries.length));
  els.meta.innerHTML = '';
  const metaItem = (label, value) => { const sp = document.createElement('span'); const b = document.createElement('b'); b.textContent = value; sp.append(`${label} `, b); els.meta.append(sp); };
  const board = currentBoard();
  const f = board && order ? await fieldsFor(board, roleOf(order)) : { picked: {}, missing: {} };
  const showField = (label, ok, miss) => { if (ok.length || miss.length) metaItem(label, ok.join(', ') + (miss.length ? `${ok.length ? ', ' : ''}${miss.join(', ')} — выберите в задаче сами` : '')); };
  showField('БЗ', f.picked.bz ? [f.picked.bz] : [], f.missing.bz ? [f.missing.bz] : []);
  showField('Страны', f.picked.countries || [], f.missing.countries || []);
  if (order && order.due) metaItem('Дедлайн', `${order.due.slice(8, 10)}.${order.due.slice(5, 7)}.${order.due.slice(0, 4)}`);
  const { subLangs = [], replyInThread = true } = await chrome.storage.local.get(['subLangs', 'replyInThread']);
  els.subsBox.hidden = !langs.length;
  els.subs.innerHTML = '';
  langs.forEach((l) => {
    const lab = document.createElement('label'); lab.className = 'lang';
    const cb = document.createElement('input'); cb.type = 'checkbox'; cb.value = l.code; cb.checked = subLangs.includes(l.code);
    cb.addEventListener('change', updateSubmit);
    lab.append(cb, `${l.code} · ${l.name}`);
    els.subs.append(lab);
  });
  els.replyRow.hidden = !(source && source.postId && source.tabId);
  els.reply.checked = replyInThread;
  updateSubmit();
  // уже есть тикет по этому сообщению/треду — предупредить (журнал + ответы в треде)
  els.exists.hidden = true;
  if (!source || !source.postId) return;
  const keys = new Set();
  const j = await journalFor(source.postId);
  if (j) [j.main, ...(j.subs || [])].forEach((x) => keys.add(x.key));
  if (source.tabId) {
    try {
      const r = await chrome.tabs.sendMessage(source.tabId, { type: 'bandThread', postId: source.postId });
      if (r && r.ok) r.keys.forEach((k) => keys.add(k));
    } catch { /* вкладку Band закрыли — обойдёмся журналом */ }
  }
  if (keys.size) {
    els.exists.textContent = `По этому треду уже есть: ${[...keys].join(', ')}. Создать ещё — только если это другая задача.`;
    els.exists.hidden = false;
  }
}
const roleOf = (o) => ({ countries: [...new Set(o.languages.map((l) => l.country))], bz: o.product, due: o.due });
const checkedLangs = () => [...els.subs.querySelectorAll('input:checked')].map((i) => i.value);
function updateSubmit() {
  const n = checkedLangs().length;
  els.submit.textContent = n ? `Создать ${n + 1} тикета` : 'Создать в трекере';
}

/** Свежий черновик из Band — забрать и стереть, чтобы не открылся второй раз. */
async function pickDraft() {
  const { trackerDraft: d } = await chrome.storage.local.get('trackerDraft');
  if (!d) return;
  await chrome.storage.local.remove('trackerDraft');
  if (Date.now() - (d.at || 0) < DRAFT_TTL) openForm(d);
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
    const langs = (order ? order.languages : []).filter((l) => checkedLangs().includes(l.code));
    const one = (l) => buildDescription(order.text, source, { ...order, languages: [l] });
    const res = await createPlan({
      board, title, description: hardBreaks(els.desc.value.trim()),
      fields: (await fieldsFor(board, roleOf(order))).fields,
      subs: await Promise.all(langs.map(async (l) => ({
        lang: l.code, title: subtaskTitle(order, l), description: hardBreaks(one(l)),
        fields: (await fieldsFor(board, roleOf({ ...order, languages: [l] }))).fields,
      }))),
    });
    created = res.main;
    await chrome.storage.local.set({ subLangs: checkedLangs(), replyInThread: els.reply.checked });
    await logCreated({ postId: source ? source.postId : '', link: source ? source.link : '', title, main: res.main, subs: res.subs, board: board.name });
    els.doneKey.textContent = res.main.key;
    els.doneKey.href = res.main.url;
    els.doneTitle.textContent = title;
    els.doneSubs.innerHTML = '';
    res.subs.forEach((x) => {
      const li = document.createElement('li'); const a = document.createElement('a');
      a.href = x.url; a.target = '_blank'; a.rel = 'noopener'; a.textContent = x.key;
      li.append(a, ` — ${x.lang}`); els.doneSubs.append(li);
    });
    els.doneSubs.hidden = !res.subs.length;
    els.doneReply.hidden = true;
    if (source && source.postId && source.tabId && els.reply.checked) {
      const text = `Завели в трекере: [${res.main.key}](${res.main.url})` + (res.subs.length ? ` — основная задача; подзадачи: ${res.subs.map((x) => `[${x.key}](${x.url}) (${x.lang})`).join(', ')}` : '');
      let r = null;
      try { r = await chrome.tabs.sendMessage(source.tabId, { type: 'bandReply', postId: source.postId, message: text }); } catch { /* вкладку закрыли */ }
      els.doneReply.textContent = r && r.ok ? '✓ Ответ со ссылками отправлен в тред' : `В тред ответить не вышло${r && r.error ? ` (${r.error})` : ''} — скопируйте ссылку и вставьте сами`;
      els.doneReply.hidden = false;
    }
    els.form.hidden = true;
    els.done.hidden = false;
    els.intro.hidden = true;
  } catch (err) {
    setMsg(err.message, 'err');
  } finally {
    els.submit.disabled = false;
    updateSubmit();
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
