// Вкладка «Мои»: открытые задачи менеджера, статус меняется прямо здесь.
import { settings, saveSettings, isConfigured, api, esc } from './core.js';

const $ = (s) => document.querySelector(s);
const els = {
  manager: $('#mManager'), refresh: $('#mRefresh'), summary: $('#mSummary'),
  list: $('#mList'), msg: $('#mMsg'), badge: $('#mineBadge'),
};

let data = null;

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
    const r = await api({ action: 'myTasks', manager: settings.manager || '' });
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
  showOverdue(data.overdue);
  if (!data.manager) {
    els.summary.textContent = 'Выберите себя в списке выше — запомню.';
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
      <select data-i="${i}" aria-label="Статус">
        <option value="">— статус —</option>
        ${data.statuses.map((s) => `<option value="${esc(s)}"${s === t.status ? ' selected' : ''}>${esc(s)}</option>`).join('')}
      </select>
    </div>`;
  }).join('');
}

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
  await saveSettings({ manager: els.manager.value });
  loadMine();
});
els.refresh.addEventListener('click', loadMine);
