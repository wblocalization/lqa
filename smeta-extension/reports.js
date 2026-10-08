// Вкладка «Отчёты»: те же отчёты, что в таблице («📈 Отчёты»), и сверка с подрядчиками.
// Excel собирает таблица и присылает сюда — файл сразу ложится в «Загрузки».
import { settings, isConfigured, call, esc, saveXlsx, copyRich, money, toast } from './core.js';

const $ = (s) => document.querySelector(s);
const els = {
  menu: $('#rMenu'), view: $('#rView'), back: $('#rBack'), title: $('#rTitle'), form: $('#rForm'),
  run: $('#rRun'), out: $('#rOut'), actions: $('#rActions'), copy: $('#rCopy'), excel: $('#rExcel'), msg: $('#rMsg'),
};

let lists = null;   // менеджеры, месяцы, тикеты… — один раз
let current = null; // открытый отчёт
let result = null;  // что сейчас на экране — для «Скопировать»

function setMsg(text, kind = 'info') {
  els.msg.textContent = text;
  els.msg.className = `status ${kind}`;
}

export async function initReports() {
  if (lists || !isConfigured()) {
    if (!isConfigured()) setMsg('Заполните настройки (⚙️), чтобы смотреть отчёты.', 'err');
    return;
  }
  // Списки — сразу из памяти, свежие — в фоне (месяцы и тикеты меняются редко)
  const key = `reportLists:${settings.url}`;
  try { lists = (await chrome.storage.local.get(key))[key] || null; } catch { /* нет — ждём таблицу */ }
  const fresh = call({ action: 'reportLists' }).then((r) => {
    lists = r;
    chrome.storage.local.set({ [key]: r }).catch(() => {});
    if (els.msg.textContent === 'Загружаю списки…') setMsg('');
  });
  if (lists) { fresh.catch(() => {}); return; }
  setMsg('Загружаю списки…');
  try {
    await fresh;
  } catch (e) {
    setMsg(`Не загрузилось: ${e.message}`, 'err');
  }
}

/** Вкладку открыли заново — показываем список отчётов. */
export function showMenu() {
  els.view.hidden = true;
  els.menu.hidden = false;
  current = null;
}

// ---------- Поля фильтров ----------
const opt = (v, label, sel) => `<option value="${esc(v)}"${v === sel ? ' selected' : ''}>${esc(label)}</option>`;
const select = (id, label, items, sel, all) =>
  `<label>${label}<select id="${id}">${all != null ? opt('', all, sel) : ''}${items.map((x) => (Array.isArray(x) ? opt(x[0], x[1], sel) : opt(x, x, sel))).join('')}</select></label>`;
const monthSel = (allLabel) => select('rfMonth', 'Месяц', lists.months.map((m) => [m.value, m.label]), lists.current, allLabel);
const bySel = () => select('rfBy', 'Считать по', [['date', 'дате поступления'], ['due', 'дате закрытия (срок сдачи)']], 'date');
const val = (id) => ($(`#${id}`) || {}).value || '';

// ---------- Отчёты ----------
const REPORTS = {
  manager: {
    title: 'По менеджеру',
    form: () => `${select('rfManager', 'Менеджер', lists.managers, settings.manager)}<div class="two">${monthSel('Все месяцы')}${bySel()}</div>`,
    run: () => call({ action: 'managerReport', manager: val('rfManager'), month: val('rfMonth'), by: val('rfBy') }),
    render: renderManager, copy: true, excel: () => ({ kind: 'manager', manager: val('rfManager'), month: val('rfMonth'), by: val('rfBy') }),
  },
  tracker: {
    title: 'Для трекера / Band',
    form: () => `<label>Тикет<input id="rfTicket" type="text" list="rfTickets" placeholder="LOCAL-1234" autocomplete="off"></label>
      <datalist id="rfTickets">${lists.tickets.map((t) => `<option value="${esc(t)}">`).join('')}</datalist>
      <p class="field-note">Задачи тикета со ссылками на Band — скопируйте и вставьте в Band.</p>`,
    run: () => {
      const t = val('rfTicket').trim();
      if (!t) throw new Error('Впишите или выберите тикет');
      return call({ action: 'trackerReport', ticket: /^\d+$/.test(t) ? `LOCAL-${t}` : t });
    },
    render: renderTracker, copy: true,
  },
  custom: {
    title: 'Кастомный',
    form: () => `<div class="two">${select('rfYear', 'Год', lists.years.map(String), String(new Date().getFullYear()), 'Всё время')}
      ${select('rfMon', 'Месяц', ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'].map((m, i) => [String(i + 1), m]), '', 'Весь год')}</div>${bySel()}`,
    run: () => call({ action: 'customReport', year: val('rfYear'), month: val('rfYear') ? val('rfMon') : '', by: val('rfBy') }),
    render: renderCustom, excel: () => ({ kind: 'custom', year: val('rfYear'), month: val('rfYear') ? val('rfMon') : '', by: val('rfBy') }),
  },
  money: {
    title: 'Сверка с подрядчиками',
    form: () => `<p class="field-note">В Excel — все колонки таблицы в том же порядке: № задачи, языки, ссылка на смету, сумма с НДС…
      Рядом — «Итоги» (подрядчики, менеджеры, языки) и «Проверить».</p>
      <div class="two">${monthSel('Всё время')}${bySel()}</div>
      <div class="two">${select('rfContractor', 'Подрядчик', lists.contractors, '', 'Все')}${select('rfMgr', 'Менеджер', lists.managers, '', 'Все')}</div>
      ${select('rfLang', 'Язык', lists.langs, '', 'Все языки')}`,
    opts: () => ({ month: val('rfMonth'), by: val('rfBy'), contractor: val('rfContractor'), manager: val('rfMgr'), lang: val('rfLang') }),
    run: () => call({ action: 'moneyPreview', opts: REPORTS.money.opts() }),
    render: renderMoney, live: true, excel: () => ({ kind: 'money', opts: REPORTS.money.opts() }),
  },
};

function open(name) {
  const r = REPORTS[name];
  if (!lists) return setMsg('Отчёты ещё грузятся — секунду…');
  current = r;
  result = null;
  els.menu.hidden = true;
  els.view.hidden = false;
  els.title.textContent = r.title;
  els.form.innerHTML = r.form();
  els.out.innerHTML = '';
  els.actions.hidden = true;
  els.copy.hidden = !r.copy;
  els.excel.hidden = !r.excel;
  els.run.hidden = !!r.live;
  setMsg('');
  // Сверка пересчитывается сама при смене фильтра
  if (r.live) {
    els.form.querySelectorAll('select').forEach((s) => s.addEventListener('change', run));
    run();
  }
}

let seq = 0;
async function run() {
  const r = current;
  const my = ++seq;
  els.run.disabled = true;
  els.actions.hidden = true;
  els.out.innerHTML = '<p class="muted center">Считаю…</p>';
  setMsg('');
  try {
    const d = await r.run();
    if (my !== seq) return;
    result = d;
    els.out.innerHTML = r.render(d);
    els.actions.hidden = !(r.copy || r.excel) || isEmpty(d);
  } catch (e) {
    if (my !== seq) return;
    els.out.innerHTML = '';
    setMsg(`Не получилось: ${e.message}`, 'err');
  } finally {
    els.run.disabled = false;
  }
}

const isEmpty = (d) => (d.groups && !d.groups.length) || d.count === 0 || d.totalTasks === 0 || d.lines === 0;
const tiles = (list) => `<div class="tiles">${list.map(([n, l, bad]) => `<div class="tile${bad ? ' bad' : ''}"><b>${esc(n)}</b><span>${esc(l)}</span></div>`).join('')}</div>`;
const kv = (rows, fmt) => (rows.length
  ? `<div class="kv">${rows.map((r) => `<div><span>${esc(r[0])}</span><b>${esc(fmt ? fmt(r) : r[1])}</b></div>`).join('')}</div>`
  : '<p class="muted">Нет данных</p>');
const link = (text, url) => (url ? `<a href="${esc(url)}" target="_blank" rel="noopener">${esc(text)}</a>` : esc(text));

function renderManager(d) {
  if (!d.groups.length) return '<p class="muted center">За этот период задач нет.</p>';
  return tiles([[d.taskCount, 'тикетов'], [d.lineCount, 'задач'], [d.grandTotal, 'SP'], [money(d.grandMoney), '₽ с НДС']]) +
    d.groups.map((g) => `<div class="r-group"><h3>${esc(g.ticket || '(без тикета)')}</h3>${g.lines.map((l) =>
      `<div class="r-line"><span>${link(l.subject, l.link)}${l.link2.split('\n').filter(Boolean).map((u, i) => ` <a class="small" href="${esc(u)}" target="_blank" rel="noopener">(ссылка ${i + 2})</a>`).join('')}</span><b>${l.sp} SP</b></div>`).join('')}
      <div class="r-line sum"><span>Итого по тикету</span><b>${g.ticketTotal} SP</b></div></div>`).join('') +
    `<div class="r-total">Итого по ${esc(d.manager)}${d.monthLabel ? ` за ${esc(d.monthLabel)}` : ''}: ${d.grandTotal} SP</div>`;
}

function renderTracker(d) {
  if (!d.count) return '<p class="muted center">По этому тикету задач нет.</p>';
  return `<p class="muted">Задач: ${d.count}</p><pre class="r-pre">${esc(d.text)}</pre>`;
}

function renderCustom(d) {
  if (!d.totalTasks) return '<p class="muted center">За этот период задач нет.</p>';
  return `<p class="muted">Период: <b>${esc(d.periodLabel)}</b></p>` +
    tiles([[d.totalTasks, 'задач'], [d.totalSp, 'SP'], [d.avgSp, 'SP на задачу'], [money(d.totalMoney), '₽ с НДС']]) +
    `<h3>Продукты</h3>${kv(d.topProducts)}<h3>Языки</h3>${kv(d.languages)}` +
    `<h3>Статус отдачи</h3>${kv(d.deliveryStatus, (r) => `${r[1]} (${r[2]}%)`)}` +
    `<h3>Топ-5 менеджеров по SP</h3>${kv(d.topManagers, (r) => `${r[2]} SP · ${r[1]} задач`)}` +
    `<h3>Топ-5 подрядчиков</h3>${kv(d.topContractors)}` +
    `<h3>Самые трудоёмкие задачи</h3>${d.topTasksBySP.length ? d.topTasksBySP.map((t) => `<div class="r-line"><span>${link(t.subject, t.link)}</span><b>${t.sp} SP</b></div>`).join('') : '<p class="muted">Нет задач с SP</p>'}`;
}

function renderMoney(d) {
  if (!d.lines) return '<p class="muted center">По этим фильтрам задач нет.</p>';
  return tiles([[d.orders, 'заказов'], [d.lines, 'строк'], [money(d.total), '₽ с НДС']]) +
    (d.issues
      ? `<p class="r-note warn">⚠️ Проверить: ${d.issues} — сумма без сметы, отдано без суммы и т. п. Список будет на листе «Проверить».</p>`
      : '<p class="r-note ok">✓ Всё сходится: у всех сумм есть сметы и наоборот.</p>');
}

// ---------- Кнопки ----------
els.menu.addEventListener('click', (e) => {
  const b = e.target.closest('[data-report]');
  if (b) open(b.dataset.report);
});
els.back.addEventListener('click', showMenu);
els.run.addEventListener('click', run);

els.copy.addEventListener('click', async () => {
  if (!result) return;
  try {
    await copyResult();
    toast('✓ Скопировано — можно вставлять');
  } catch {
    setMsg('Не получилось скопировать — разрешите доступ к буферу обмена', 'err');
  }
});

async function copyResult() {
  if (current === REPORTS.tracker) {
    // В Band ссылки вида [текст](адрес) становятся кликабельными
    await navigator.clipboard.writeText(result.text);
  } else {
    const lines = result.groups.flatMap((g) => g.lines);
    const html = lines.map((l) => `<div>${l.link ? `<a href="${esc(l.link)}">${esc(l.subject)}</a>` : esc(l.subject)} — ${l.sp} SP</div>`).join('') +
      `<div><b>Итого по ${esc(result.manager)}: ${result.grandTotal} SP</b></div>`;
    await copyRich(html, result.copyText);
  }
}

els.excel.addEventListener('click', async () => {
  const label = els.excel.textContent;
  els.excel.disabled = true;
  els.excel.textContent = 'Собираю файл…';
  setMsg('Таблица собирает Excel — обычно 5–20 секунд');
  try {
    const r = await call({ action: 'exportXlsx', ...current.excel() });
    saveXlsx(r.name, r.b64);
    setMsg('');
    toast('✓ Файл в «Загрузках»');
  } catch (e) {
    setMsg(`Не получилось: ${e.message}`, 'err');
  } finally {
    els.excel.disabled = false;
    els.excel.textContent = label;
  }
});
