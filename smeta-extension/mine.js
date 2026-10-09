// Вкладка «Мои задачи»: открытые задачи менеджера, статус меняется прямо здесь.
import { settings, saveSettings, isConfigured, api, esc } from './core.js';
import { openEditor } from './edit.js';
import { openMail } from './outlook.js';


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

export async function loadMine({ keepMsg = false } = {}) {
  if (!isConfigured()) {
    els.summary.textContent = 'Заполните настройки (⚙️), чтобы видеть свои задачи.';
    return;
  }
  // Сначала — список с прошлого раза (мгновенно), потом тихо обновляем из таблицы
  const manager = viewing || settings.manager || '';
  const key = `mine:${manager}:${filter}`;
  const seq = ++loadSeq;
  let cached = null;
  try { cached = (await chrome.storage.local.get(key))[key]; } catch { /* нет — просто ждём таблицу */ }
  if (cached && seq === loadSeq) {
    data = notDeleted(cached);
    render();
    els.summary.textContent += ' · обновляю…';
  } else {
    els.summary.textContent = 'Загружаю…';
  }
  if (!keepMsg) setMsg('');
  try {
    const r = await api({ action: 'myTasks', manager, filter });
    if (!r.ok) throw new Error(r.error);
    if (seq !== loadSeq) return; // пока грузили, переключили фильтр или менеджера
    data = notDeleted(r);
    render();
  } catch (e) {
    if (seq !== loadSeq) return;
    if (cached) els.summary.textContent = els.summary.textContent.replace(' · обновляю…', '');
    else els.summary.textContent = '';
    setMsg(`Не получилось ${cached ? 'обновить' : 'загрузить'} задачи: ${e.message}`, 'err');
  }
}

let loadSeq = 0;

/** Запоминаем, что показали, — в следующий раз список появится сразу. */
function saveCache() {
  if (!data || !data.manager) return;
  const key = `mine:${data.manager}:${data.filter || 'open'}`;
  chrome.storage.local.set({ [key]: data }).catch(() => {});
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
  saveCache();
  const n = data.tasks.length;
  // Только что отданная карточка ещё на экране, но в счёт открытых не входит
  const count = data.tasks.filter((t) => matches(filter, t.status)).length;
  const more = data.more ? ` (показаны последние ${n} из ${n + data.more})` : '';
  els.summary.textContent = count
    ? `${FILTER_TEXT[filter]} задач: ${count + (data.more || 0)}${more}${filter === 'open' && data.overdue ? ` · просрочено: ${data.overdue}` : ''}`
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
      <div class="meta-row"><div class="meta">${esc(t.id || '—')}${t.ticket ? ' · ' + esc(t.ticket) : ''} · ${due}</div>
        ${t.id ? `<button class="link-btn mail-btn" type="button" data-mail="${i}" title="Найти письма по этой задаче в Outlook">✉️ Почта</button>` : ''}</div>
      <div class="row-actions">
        <select data-i="${i}" class="status-sel" data-status="${esc(t.status)}" aria-label="Статус">
          <option value="">— статус —</option>
          ${data.statuses.map((s) => `<option value="${esc(s)}"${s === t.status ? ' selected' : ''}>${esc(s)}</option>`).join('')}
        </select>
        <button class="btn ghost small" type="button" data-edit="${i}">Изменить</button>
      </div>
      ${t.status === 'Отдано' ? deliveryRow(t, i) : ''}
    </div>`;
  }).join('');
}

/** «Статус отдачи» — под «Отдано»: вовремя ли отчитались перед заказчиком. */
const DELIVERY_FALLBACK = ['Отдано в срок', 'Сообщили о просрочке', 'Не сообщили о просрочке'];
function deliveryRow(t, i) {
  const list = (data.deliveryStatuses && data.deliveryStatuses.length) ? data.deliveryStatuses : DELIVERY_FALLBACK;
  return `<label class="delivery-row"><span>Статус отдачи</span>
    <select data-d="${i}" class="delivery-sel" data-delivery="${esc(t.delivery || '')}" aria-label="Статус отдачи">
      <option value="">— выберите —</option>
      ${list.map((s) => `<option value="${esc(s)}"${s === t.delivery ? ' selected' : ''}>${esc(s)}</option>`).join('')}
    </select></label>`;
}

// ---------- Правка ----------
/** После сохранения — обновить списки и сказать, что получилось. */
async function afterSave(text) {
  setMsg(text, 'ok');
  await loadMine({ keepMsg: true });
  if (els.search.value.trim()) await runSearch();
}

// Удалённые в этой сессии — чтобы список с прошлого раза (кэш) не показал их снова
const deleted = new Set();
const refKey = (t) => `${t.id || ''}|${t.row}`;
const notDeleted = (d) => (d && d.tasks ? { ...d, tasks: d.tasks.filter((t) => !deleted.has(refKey(t)) && !deleted.has(`${t.id || ''}|*`)) } : d);

/** Задача пропадает из списка сразу, таблица удаляет строку в фоне. Не получилось — список вернётся как был. */
async function deleteTask(ref) {
  const label = ref.id || 'Задача';
  if (ref.id) deleted.add(`${ref.id}|*`);
  else deleted.add(refKey(ref));
  const i = data && data.tasks ? data.tasks.findIndex((t) => (ref.id ? t.id === ref.id : t.row === ref.row)) : -1;
  if (i !== -1) {
    const [t] = data.tasks.splice(i, 1);
    if (data.counts) ['open', 'done', 'cancelled', 'all'].forEach((f) => { if (matches(f, t.status) && data.counts[f]) data.counts[f]--; });
    if (t.overdue) data.overdue--;
    render();
  }
  if (!els.found.hidden && found.length) { found = notDeleted({ tasks: found }).tasks; renderFound(); }
  setMsg(`${label}: удаляю…`);
  try {
    const r = await api({ action: 'deleteTask', ...ref });
    if (!r.ok && /Неизвестное действие/.test(r.error || '')) {
      throw new Error('веб-приложение ещё старое. В Apps Script: «Развернуть → Управление развёртываниями → ✏️ → Версия: новая → Развернуть»');
    }
    if (!r.ok) throw new Error(r.error);
    setMsg(`${label}: удалена`, 'ok');
  } catch (err) {
    deleted.delete(`${ref.id}|*`);
    deleted.delete(refKey(ref));
    setMsg(`${label}: не удалилась — ${err.message}`, 'err');
  }
  // Строки в таблице сдвинулись — берём свежий список (сообщение не трогаем)
  await loadMine({ keepMsg: true });
  if (els.search.value.trim()) await runSearch();
  if (!ref.id) deleted.delete(refKey(ref)); // без номера помним по строке, а строки сдвинулись — дальше не нужно
}

els.list.addEventListener('click', (e) => {
  const mail = e.target.closest('button[data-mail]');
  if (mail) return openMail(data.tasks[Number(mail.dataset.mail)].id);
  const b = e.target.closest('button[data-edit]');
  if (!b) return;
  const t = data.tasks[Number(b.dataset.edit)];
  openEditor({ row: t.row, id: t.id, origSubject: t.subject }, afterSave, deleteTask);
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
    found = notDeleted(r).tasks;
    renderFound();
  } catch (e) {
    els.found.innerHTML = `<p class="status err">Не получилось найти: ${esc(e.message)}</p>`;
  }
}
function renderFound() {
  els.found.innerHTML = found.length
    ? found.map((t, i) => `<div class="mine-item">
        <div class="title">${esc(t.title)}</div>
        ${t.id ? `<div class="meta-row"><span></span><button class="link-btn mail-btn" type="button" data-mail-found="${i}" title="Найти письма по этой задаче в Outlook">✉️ Почта</button></div>` : ''}
        <div class="row-actions"><span class="meta" style="flex:1">${esc(t.id || '—')}${t.ticket ? ' · ' + esc(t.ticket) : ''} · ${esc(t.date)}${t.status ? ' · ' + esc(t.status) : ''}</span>
        <button class="btn ghost small" type="button" data-found="${i}">Изменить</button></div>
      </div>`).join('')
    : '<p class="muted">Ничего не нашлось.</p>';
}
els.searchForm.addEventListener('submit', (e) => { e.preventDefault(); runSearch(); });
els.search.addEventListener('search', () => { if (!els.search.value) els.found.hidden = true; }); // крестик в поле
els.found.addEventListener('click', (e) => {
  const mail = e.target.closest('button[data-mail-found]');
  if (mail) return openMail(found[Number(mail.dataset.mailFound)].id);
  const b = e.target.closest('button[data-found]');
  if (!b) return;
  const t = found[Number(b.dataset.found)];
  openEditor({ row: t.row, id: t.id }, afterSave, deleteTask);
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
  if (!data || !data.counts) return;
  ['open', 'done', 'cancelled'].forEach((f) => {
    if (matches(f, from)) data.counts[f]--;
    if (matches(f, to)) data.counts[f]++;
  });
}

function renderFilter() {
  const c = (data && data.counts) || {}; // список ещё не загрузился — фильтр без цифр
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
// Статус отдачи — сразу в таблицу
els.list.addEventListener('change', async (e) => {
  const sel = e.target.closest('select[data-d]');
  if (!sel) return;
  const t = data.tasks[Number(sel.dataset.d)];
  const prev = t.delivery || '';
  t.delivery = sel.value;
  sel.dataset.delivery = sel.value;
  setMsg(`${t.id || 'Задача'}: «${sel.value || 'без статуса отдачи'}» — сохраняю…`);
  try {
    const r = await api({ action: 'setDelivery', row: t.row, id: t.id, origSubject: t.subject, value: sel.value });
    if (!r.ok) throw new Error(/Неизвестное действие/.test(r.error || '') ? 'веб-приложение ещё старое: нужна новая версия развёртывания' : r.error);
    saveCache();
    setMsg(`${t.id || 'Задача'}: статус отдачи «${sel.value || '—'}» ✓`, 'ok');
  } catch (err) {
    t.delivery = prev;
    sel.value = prev;
    sel.dataset.delivery = prev;
    setMsg(`${t.id || 'Задача'}: не сохранилось — ${err.message}`, 'err');
  }
});

els.list.addEventListener('change', async (e) => {
  const sel = e.target.closest('select[data-i]');
  if (!sel) return;
  const i = Number(sel.dataset.i);
  const t = data.tasks[i];
  const prev = t.status;
  const status = sel.value;
  // «Отдано» — карточка остаётся до обновления, чтобы сразу выбрать статус отдачи
  const gone = !matches(filter, status) && status !== 'Отдано';
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
