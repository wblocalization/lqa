// Вкладка «Мои задачи»: открытые задачи менеджера, статус меняется прямо здесь.
import { settings, saveSettings, isConfigured, api, esc } from './core.js';
import { openEditor } from './edit.js';

const $ = (s) => document.querySelector(s);
const els = {
  manager: $('#mManager'), refresh: $('#mRefresh'), summary: $('#mSummary'),
  list: $('#mList'), msg: $('#mMsg'), badge: $('#mineBadge'),
  filter: $('#mFilter'), searchForm: $('#mSearchForm'), search: $('#mSearch'), found: $('#mFound'),
};

let data = null;
let viewing = '';
// Какие задачи показывать: open / done / cancelled / all — запоминается
let filter = 'open';
try { filter = localStorage.getItem('mineFilter') || 'open'; } catch { /* по умолчанию открытые */ }
const CLOSED = ['Отдано', 'Отменено'];
const FILTER_TEXT = { open: 'Открытых', done: 'Отданных', cancelled: 'Отменённых', all: 'Всего' }; // можно посмотреть задачи коллеги, не меняя «Кто вы» в настройках

function setMsg(text, kind = 'info') {
  els.msg.textContent = text;
  els.msg.className = `status ${kind}`;
}

/** Цифра просрочек — на вкладке и на иконке расширения. */
export function showOverdue(n) {
  els.badge.hidden = !n;
  els.badge.textContent = n ? String(n) : '';
  chrome.action.setBadgeText({ text: n ? String(n) : '' });
  chrome.action.setBadgeBackgroundColor({ color: '#C0262D' });
}

export async function loadMine() {
  if (!isConfigured()) {
    els.summary.textContent = 'Заполните настройки (⚙️), чтобы видеть свои задачи.';
    return;
  }
  els.summary.textContent = 'Загружаю…';
  setMsg('');
  try {
    const r = await api({ action: 'myTasks', manager: viewing || settings.manager || '', filter });
    if (!r.ok) throw new Error(r.error);
    data = r;
    render();
  } catch (e) {
    els.summary.textContent = '';
    setMsg(`Не получилось загрузить задачи: ${e.message}`, 'err');
  }
}

function render() {
  els.manager.innerHTML = `<option value="">— выберите себя —</option>` +
    data.managers.map((m) => `<option value="${esc(m)}"${m === data.manager ? ' selected' : ''}>${esc(m)}</option>`).join('');
  if (!data.manager || data.manager === settings.manager) showOverdue(data.overdue);
  if (!data.manager) {
    els.summary.textContent = 'Выберите себя в списке выше (или в ⚙️ «Кто вы») — запомню.';
    els.list.innerHTML = '';
    return;
  }
  renderFilter();
  const n = data.tasks.length;
  const more = data.more ? ` (показаны последние ${n} из ${n + data.more})` : '';
  els.summary.textContent = n
    ? `${FILTER_TEXT[filter]} задач: ${n + (data.more || 0)}${more}${filter === 'open' && data.overdue ? ` · просрочено: ${data.overdue}` : ''}`
    : (filter === 'open' ? 'Открытых задач нет 🎉' : 'Таких задач нет');
  els.list.innerHTML = data.tasks.map((t, i) => {
    const closed = CLOSED.includes(t.status);
    const due = closed ? `${esc(t.status.toLowerCase())}${t.date ? ' · от ' + esc(t.date) : ''}`
      : t.overdue ? `<span class="late">просрочено · ${esc(t.due)}</span>`
      : t.dueToday ? `<span class="today">сегодня</span>`
      : t.due ? `срок ${esc(t.due)}` : (t.deadline ? esc(t.deadline) : 'срок не указан');
    const title = t.link ? `<a href="${esc(t.link)}" target="_blank" rel="noopener">${esc(t.title)} 🔗</a>` : esc(t.title);
    return `<div class="mine-item${t.overdue ? ' overdue' : ''}">
      <div class="title">${title}</div>
      <div class="meta">${esc(t.id || '—')}${t.ticket ? ' · ' + esc(t.ticket) : ''} · ${due}</div>
      <div class="row-actions">
        <select data-i="${i}" aria-label="Статус">
          <option value="">— статус —</option>
          ${data.statuses.map((s) => `<option value="${esc(s)}"${s === t.status ? ' selected' : ''}>${esc(s)}</option>`).join('')}
        </select>
        <button class="btn ghost small" type="button" data-edit="${i}">Изменить</button>
      </div>
    </div>`;
  }).join('');
}

// ---------- Правка ----------
/** После сохранения — обновить списки и сказать, что получилось. */
async function afterSave(text) {
  await loadMine();
  if (els.search.value.trim()) await runSearch();
  setMsg(text, 'ok');
}

els.list.addEventListener('click', (e) => {
  const b = e.target.closest('button[data-edit]');
  if (!b) return;
  const t = data.tasks[Number(b.dataset.edit)];
  openEditor({ row: t.row, id: t.id, origSubject: t.subject }, afterSave);
});

// Поиск любой задачи — не только своих открытых
let found = [];
async function runSearch() {
  const q = els.search.value.trim();
  if (!q) { els.found.hidden = true; return; }
  els.found.hidden = false;
  els.found.innerHTML = '<p class="muted">Ищу…</p>';
  try {
    const r = await api({ action: 'searchTasks', query: q });
    if (!r.ok) throw new Error(r.error);
    found = r.tasks;
    els.found.innerHTML = found.length
      ? found.map((t, i) => `<div class="mine-item">
          <div class="title">${esc(t.title)}</div>
          <div class="row-actions"><span class="meta" style="flex:1">${esc(t.id || '—')}${t.ticket ? ' · ' + esc(t.ticket) : ''} · ${esc(t.date)}${t.status ? ' · ' + esc(t.status) : ''}</span>
          <button class="btn ghost small" type="button" data-found="${i}">Изменить</button></div>
        </div>`).join('')
      : '<p class="muted">Ничего не нашлось.</p>';
  } catch (e) {
    els.found.innerHTML = `<p class="status err">Не получилось найти: ${esc(e.message)}</p>`;
  }
}
els.searchForm.addEventListener('submit', (e) => { e.preventDefault(); runSearch(); });
els.search.addEventListener('search', () => { if (!els.search.value) els.found.hidden = true; }); // крестик в поле
els.found.addEventListener('click', (e) => {
  const b = e.target.closest('button[data-found]');
  if (!b) return;
  const t = found[Number(b.dataset.found)];
  openEditor({ row: t.row, id: t.id }, afterSave);
});

// Статус меняется сразу: задача, которая больше не подходит под фильтр, пропадает из списка,
// а сохранение идёт в фоне. Не получилось — задача возвращается на место и видно ошибку.
function matches(f, status) {
  if (f === 'done') return status === 'Отдано';
  if (f === 'cancelled') return status === 'Отменено';
  if (f === 'all') return true;
  return !CLOSED.includes(status);
}

/** Счётчики в фильтре: задача перешла из одного статуса в другой. */
function recount(from, to) {
  if (!data.counts) return;
  ['open', 'done', 'cancelled'].forEach((f) => {
    if (matches(f, from)) data.counts[f]--;
    if (matches(f, to)) data.counts[f]++;
  });
}

function renderFilter() {
  const c = data.counts || {};
  els.filter.querySelectorAll('button').forEach((b) => {
    const f = b.dataset.f;
    b.setAttribute('aria-pressed', String(f === filter));
    b.innerHTML = `${b.dataset.label || (b.dataset.label = b.textContent)}${c[f] != null ? ` <span class="n">${c[f]}</span>` : ''}`;
  });
}

els.filter.addEventListener('click', (e) => {
  const b = e.target.closest('button[data-f]');
  if (!b || b.dataset.f === filter) return;
  filter = b.dataset.f;
  try { localStorage.setItem('mineFilter', filter); } catch { /* не страшно */ }
  els.list.innerHTML = '';
  renderFilter();
  loadMine();
});
els.list.addEventListener('change', async (e) => {
  const sel = e.target.closest('select[data-i]');
  if (!sel) return;
  const i = Number(sel.dataset.i);
  const t = data.tasks[i];
  const prev = t.status;
  const status = sel.value;
  const gone = !matches(filter, status);
  const wasLate = t.overdue;
  t.status = status;
  recount(prev, status);
  if (CLOSED.includes(status) && t.overdue) { t.overdue = false; data.overdue--; }
  if (gone) data.tasks.splice(i, 1);
  render();
  setMsg(`${t.id || 'Задача'}: «${status || 'пусто'}» — сохраняю…`);
  try {
    const r = await api({ action: 'setStatus', row: t.row, id: t.id, origSubject: t.subject, status });
    if (!r.ok) throw new Error(r.error);
    setMsg(`${t.id || 'Задача'}: статус «${status || 'пусто'}» ✓`, 'ok');
  } catch (err) {
    recount(status, prev);
    t.status = prev;
    if (wasLate && !t.overdue) { t.overdue = true; data.overdue++; }
    if (gone) data.tasks.splice(Math.min(i, data.tasks.length), 0, t);
    render();
    setMsg(`${t.id || 'Задача'}: не сохранилось — ${err.message}`, 'err');
  }
});


els.manager.addEventListener('change', async () => {
  if (!settings.manager) await saveSettings({ manager: els.manager.value, user: els.manager.value });
  else viewing = els.manager.value === settings.manager ? '' : els.manager.value;
  loadMine();
});
els.refresh.addEventListener('click', loadMine);
