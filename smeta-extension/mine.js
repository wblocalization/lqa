// Вкладка «Мои задачи»: открытые задачи менеджера, статус меняется прямо здесь.
import { settings, saveSettings, isConfigured, api, esc } from './core.js';
import { openEditor } from './edit.js';

const $ = (s) => document.querySelector(s);
const els = {
  manager: $('#mManager'), refresh: $('#mRefresh'), summary: $('#mSummary'),
  list: $('#mList'), msg: $('#mMsg'), badge: $('#mineBadge'),
  searchForm: $('#mSearchForm'), search: $('#mSearch'), found: $('#mFound'),
};

let data = null;
let viewing = ''; // можно посмотреть задачи коллеги, не меняя «Кто вы» в настройках

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
    const r = await api({ action: 'myTasks', manager: viewing || settings.manager || '' });
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
  const n = data.tasks.length;
  els.summary.textContent = n
    ? `Открытых задач: ${n}${data.overdue ? ` · просрочено: ${data.overdue}` : ''}`
    : 'Открытых задач нет 🎉';
  els.list.innerHTML = data.tasks.map((t, i) => {
    const due = t.overdue ? `<span class="late">просрочено · ${esc(t.due)}</span>`
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

els.list.addEventListener('change', async (e) => {
  const sel = e.target.closest('select[data-i]');
  if (!sel) return;
  const t = data.tasks[Number(sel.dataset.i)];
  sel.disabled = true;
  setMsg('Сохраняю…');
  try {
    const r = await api({ action: 'setStatus', row: t.row, id: t.id, origSubject: t.subject, status: sel.value });
    if (!r.ok) throw new Error(r.error);
    await loadMine(); // закрытые пропадут из списка, строки могли сдвинуться
    setMsg(`${t.id || 'Задача'}: статус «${sel.value || 'пусто'}» ✓`, 'ok');
  } catch (err) {
    setMsg(`Не сохранилось: ${err.message}`, 'err');
    sel.disabled = false;
  }
});

els.manager.addEventListener('change', async () => {
  if (!settings.manager) await saveSettings({ manager: els.manager.value, user: els.manager.value });
  else viewing = els.manager.value === settings.manager ? '' : els.manager.value;
  loadMine();
});
els.refresh.addEventListener('click', loadMine);
