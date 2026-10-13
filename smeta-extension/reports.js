// Вкладка «Отчёты»: те же отчёты, что в таблице («📈 Отчёты»), и сверка с подрядчиками.
// Excel собирает таблица и присылает сюда — файл сразу ложится в «Загрузки».
import { settings, isConfigured, call, esc, saveXlsx, copyRich, money, toast } from './core.js';
import { parseAct, repairIdBreaks, linesFromItems } from './parser.js';
import { readPdfTexts } from './smeta.js';

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
    title: 'По менеджеру — для Band',
    // «Кто вы» не выбран — человека выбирают здесь
    form: () => `${select('rfManager', 'Чей отчёт', lists.managers, settings.manager, settings.manager ? null : 'Выберите…')}
      <div class="two">${monthSel('Все месяцы')}${bySel()}</div>
      <p class="field-note">Все тикеты за месяц: задачи со ссылками на Band внутри и SP — «Скопировать для Band» и вставить.</p>`,
    run: () => {
      if (!val('rfManager')) throw new Error('Выберите, чей отчёт');
      return call({ action: 'managerReport', manager: val('rfManager'), month: val('rfMonth'), by: val('rfBy') });
    },
    render: renderManager, copy: true, excel: () => ({ kind: 'manager', manager: val('rfManager'), month: val('rfMonth'), by: val('rfBy') }),
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
      ${select('rfLang', 'Язык', lists.langs, '', 'Все языки')}
      <div id="actBox" class="act-box">
        <b>Акт или счёт подрядчика за месяц</b>
        <span>Загрузите PDF — строки с одним номером сложатся (по сумме с НДС) и сравнятся с таблицей.
          Выберите выше месяц и подрядчика — тогда покажем и задачи, которых в акте нет.</span>
        <button id="actPick" class="btn ghost" type="button">📄 Загрузить акт подрядчика (PDF)</button>
        <input id="actFile" type="file" accept="application/pdf,.pdf" hidden>
        <div id="actOut"></div>
      </div>`,
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
  els.copy.textContent = '📋 Скопировать всё для Band';
  els.excel.hidden = !r.excel;
  els.run.hidden = !!r.live;
  setMsg('');
  // Сверка пересчитывается сама при смене фильтра
  if (r.live) {
    els.form.querySelectorAll('select').forEach((s) => s.addEventListener('change', () => { run(); if (act) compareAct(); }));
    run();
  }
  if (name === 'money') initAct();
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
    d.groups.map((g, gi) => `<div class="r-group"><h3>${esc(g.ticket || '(без тикета)')}</h3>${g.lines.map((l) =>
      `<div class="r-line"><span>${link(bandLabel(l.subject), l.link)}${l.link2.split('\n').filter(Boolean).map((u, i) => ` <a class="small" href="${esc(u)}" target="_blank" rel="noopener">(ссылка ${i + 2})</a>`).join('')}</span><b>${l.sp} SP</b></div>`).join('')}
      <div class="r-line sum"><span>Итого по тикету</span><b>${g.ticketTotal} SP</b></div>
      <button class="btn ghost small r-copy-one" type="button" data-group="${gi}">📋 Скопировать задачи ${esc(g.ticket || 'без тикета')}</button></div>`).join('') +
    `<div class="r-total">Итого по ${esc(d.manager)}${d.monthLabel ? ` за ${esc(d.monthLabel)}` : ''}: ${d.grandTotal} SP</div>`;
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

// «Скопировать» под тикетом — только его задачи: вставляются каждый в свой тикет трекера
els.out.addEventListener('click', async (e) => {
  const b = e.target.closest('button[data-group]');
  if (!b || !result) return;
  const g = result.groups[Number(b.dataset.group)];
  const label = b.textContent;
  try {
    const one = bandTicket(g);
    await copyRich(one.html, one.text);
    toast(`✓ ${g.ticket || 'Задачи без тикета'} — скопировано`);
    b.textContent = '✓ Скопировано';
    setTimeout(() => { b.textContent = label; }, 1500);
  } catch {
    setMsg('Не получилось скопировать — разрешите доступ к буферу обмена', 'err');
  }
});
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

/** «[LIT-26-2237][LogrusIT][kk][Магазинка] Новые строки» → «LIT-26-2237 · Новые строки» — без скобок, чтобы ссылка в Band не сломалась. */
function bandLabel(subject) {
  const id = (String(subject).match(/^\s*\[([^\]]+)\]/) || [])[1] || '';
  const short = String(subject).replace(/^(\s*\[[^\]]*\])+\s*/, '').trim() || String(subject);
  return id ? `${id} · ${short}` : short;
}

/**
 * Отчёт для Band: по тикетам, ссылка спрятана в названии задачи, в конце — SP.
 * text — разметка Band ([название](ссылка)), html — то же со ссылками для почты и документов.
 */
/** Задачи одного тикета для Band: ссылка внутри названия, в конце — SP, внизу — итог по тикету. */
function bandTicket(g) {
  const text = [], html = ['<ol>'];
  g.lines.forEach((l, i) => {
    const label = bandLabel(l.subject);
    const extra = l.link2.split('\n').filter(Boolean);
    text.push(`${i + 1}. ${l.link ? `[${label}](${l.link})` : label}${extra.map((u, k) => ` ([ссылка ${k + 2}](${u}))`).join('')} — ${l.sp} SP`);
    html.push(`<li>${l.link ? `<a href="${esc(l.link)}">${esc(label)}</a>` : esc(label)}${extra.map((u, k) => ` (<a href="${esc(u)}">ссылка ${k + 2}</a>)`).join('')} — ${l.sp} SP</li>`);
  });
  text.push(`Итого: ${g.ticketTotal} SP`);
  html.push(`</ol><p>Итого: ${g.ticketTotal} SP</p>`);
  return { text: text.join('\n'), html: html.join('') };
}

/**
 * Весь отчёт для Band: по тикетам с заголовками.
 * text — разметка Band ([название](ссылка)), html — то же со ссылками для почты, документов и трекеров с оформлением.
 */
function bandReport(d) {
  const text = [], html = [];
  d.groups.forEach((g) => {
    const one = bandTicket(g);
    text.push(`**${g.ticket || 'Без тикета'}**`, one.text, '');
    html.push(`<p><b>${esc(g.ticket || 'Без тикета')}</b></p>${one.html}`);
  });
  const total = `Итого${d.monthLabel ? ` за ${d.monthLabel}` : ''}: ${d.grandTotal} SP`;
  text.push(`**${total}**`);
  html.push(`<p><b>${esc(total)}</b></p>`);
  return { text: text.join('\n'), html: html.join('') };
}

async function copyResult() {
  const r = bandReport(result);
  await copyRich(r.html, r.text);
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

// ---------- Сверка с актом подрядчика ----------
let act = null;      // { name, items, total, unread } — разобранный PDF
let actResult = null; // строки для таблицы и CSV

function initAct() {
  act = null;
  actResult = null;
  const box = $('#actBox'), file = $('#actFile');
  $('#actPick').addEventListener('click', () => file.click());
  file.addEventListener('change', () => { if (file.files[0]) loadAct(file.files[0]); file.value = ''; });
  ['dragenter', 'dragover'].forEach((t) => box.addEventListener(t, (e) => { e.preventDefault(); box.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((t) => box.addEventListener(t, () => box.classList.remove('over')));
  box.addEventListener('drop', (e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) loadAct(f); });
}

async function loadAct(file) {
  const out = $('#actOut');
  if (!/\.pdf$/i.test(file.name) && file.type !== 'application/pdf') { out.innerHTML = '<p class="r-note warn">Это не PDF</p>'; return; }
  out.innerHTML = '<p class="muted center">Читаю акт…</p>';
  try {
    const pdf = await readPdfTexts(file);
    const r = parseAct(linesFromItems(repairIdBreaks(pdf.items)));
    if (!r.items.length) {
      out.innerHTML = `<p class="r-note warn">В «${esc(file.name)}» не нашлось строк с номером задачи вида LIT-26-2240 и суммой.
        ${pdf.texts.length ? 'Возможно, номера в акте записаны иначе — пришлите разработчику скриншот одной строки.' : 'В PDF нет текста — похоже, это скан (картинка). Попросите у подрядчика PDF из 1С/Excel.'}</p>`;
      return;
    }
    act = { name: file.name, ...r };
    await compareAct();
  } catch (e) {
    out.innerHTML = `<p class="r-note warn">Не получилось прочитать PDF: ${esc(e.message)}</p>`;
  }
}

const key = (id) => String(id).trim().toUpperCase().replace(/-0+(\d)/g, '-$1');
const near = (a, b) => Math.abs(a - b) < 0.01;

async function compareAct() {
  const out = $('#actOut');
  if (!out || !act) return;
  out.innerHTML = '<p class="muted center">Сравниваю с таблицей…</p>';
  let d;
  try {
    d = await call({ action: 'actCompare', ids: act.items.map((g) => g.id), opts: REPORTS.money.opts() });
  } catch (e) {
    out.innerHTML = `<p class="r-note warn">Не получилось: ${esc(e.message)}</p>`;
    return;
  }
  const rows = act.items.map((g) => {
    const t = d.found[key(g.id)];
    const table = t && t.hasTotal ? t.sum : null;
    const kind = !t ? 'none' : table == null ? 'nosum' : near(table, g.gross) ? 'ok' : 'diff';
    return { id: g.id, act: g.gross, table, diff: table == null ? null : Math.round((g.gross - table) * 1000) / 1000, kind, t, g };
  });
  const extra = (d.extra || []).map((t) => ({ id: t.id, act: null, table: t.hasTotal ? t.sum : null, diff: null, kind: 'extra', t }));
  const order = { diff: 0, none: 1, nosum: 2, extra: 3, ok: 4 };
  actResult = rows.concat(extra).sort((a, b) => order[a.kind] - order[b.kind]);
  const bad = actResult.filter((r) => r.kind !== 'ok').length;
  const tableSum = rows.reduce((a, r) => a + (r.table || 0), 0);
  const LABEL = { ok: '✅ сходится', diff: '⚠️ разница', none: '❓ номера нет в таблице', nosum: '❓ в таблице нет суммы', extra: '➕ в таблице есть, в акте нет' };
  const line = (r) => `<details class="act-row ${r.kind}"><summary>
      <span class="act-id">${esc(r.id)}</span>
      <span class="act-st">${LABEL[r.kind]}${r.kind === 'diff' ? ` ${r.diff > 0 ? '+' : ''}${money(r.diff)} ₽` : ''}</span>
      <span class="act-sum">${r.act != null ? `акт ${money(r.act)}` : ''}${r.act != null && r.table != null ? ' · ' : ''}${r.table != null ? `таблица ${money(r.table)}` : ''}</span></summary>
      ${r.t ? `<p class="act-subj">${esc(r.t.subject)} · ${esc(r.t.contractor)} · ${esc(r.t.manager)} · строк в таблице: ${r.t.lines}${r.t.bad ? ` · в сумме не число: «${esc(r.t.bad)}»` : ''}</p>` : ''}
      ${r.g ? r.g.lines.map((l) => `<p class="act-line">${esc(l.text)} → <b>${money(l.gross)}</b></p>`).join('') : ''}
    </details>`;
  out.innerHTML = `<p class="act-file">📄 ${esc(act.name)}</p>` +
    tiles([[money(act.total), '₽ в акте'], [money(tableSum), '₽ в таблице'], [bad, bad ? 'проверить' : 'всё ок', bad > 0]]) +
    (bad ? '' : '<p class="r-note ok">✓ Акт полностью сходится с таблицей.</p>') +
    actResult.map(line).join('') +
    (act.unread.length ? `<p class="r-note warn">Эти строки акта не разобрал — проверьте вручную:<br>${act.unread.map(esc).join('<br>')}</p>` : '') +
    '<button id="actCsv" class="btn ghost" type="button">📥 Скачать сверку (Excel)</button>';
  $('#actCsv').addEventListener('click', downloadAct);
}

/** CSV для Excel: «;» и запятая в копейках — открывается двойным кликом. */
function downloadAct() {
  const num = (x) => (x == null ? '' : String(x).replace('.', ','));
  const cell = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const KIND = { ok: 'сходится', diff: 'разница', none: 'нет в таблице', nosum: 'в таблице нет суммы', extra: 'нет в акте' };
  const head = ['Номер', 'Итог', 'В акте, с НДС', 'В таблице', 'Разница (акт − таблица)', 'Тема', 'Подрядчик', 'Менеджер', 'Строк в акте', 'Строк в таблице'];
  const lines = actResult.map((r) => [r.id, KIND[r.kind], num(r.act), num(r.table), num(r.diff), r.t ? r.t.subject : '', r.t ? r.t.contractor : '', r.t ? r.t.manager : '', r.g ? r.g.lines.length : '', r.t ? r.t.lines : '']);
  const csv = '﻿' + [head, ...lines].map((l) => l.map(cell).join(';')).join('\r\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  a.download = `Сверка с актом — ${act.name.replace(/\.pdf$/i, '')}.csv`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  toast('✓ Файл в «Загрузках»');
}
