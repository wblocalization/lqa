// Правка задачи из расширения: то же, что «Поиск и правка» в таблице.
import { api, esc, getLists, autoDeadline } from './core.js';
import { makeLinkList } from './links.js';

const $ = (s) => document.querySelector(s);
const els = {
  view: $('#mineView'), form: $('#editForm'), back: $('#eBack'), title: $('#eTitle'),
  subject: $('#eSubject'), link: $('#eLink'), link2: $('#eLink2'), ticketNum: $('#eTicketNum'), date: $('#eDate'),
  product: $('#eProduct'), customer: $('#eCustomer'), langs: $('#eLangs'), deadline: $('#eDeadline'),
  exactDeadline: $('#eExactDeadline'), status: $('#eStatus'), deliveryStatus: $('#eDeliveryStatus'),
  contractor: $('#eContractor'), manager: $('#eManager'), estimateLink: $('#eEstimateLink'),
  total: $('#eTotal'), sp: $('#eSp'), comment: $('#eComment'), complaints: $('#eComplaints'), save: $('#eSave'), del: $('#eDelete'), msg: $('#eMsg'), top: $('#eTop'),
};
autoDeadline(els.deadline, els.exactDeadline, els.date);

let lists = null;    // справочники из таблицы — один раз
let current = null;  // задача, как её отдала таблица (row, id, origSubject — чтобы найти строку при сохранении)
let onDone = () => {};
let onDelete = () => {};

function setMsg(text, kind = 'info') {
  els.msg.textContent = text;
  els.msg.className = `status ${kind}`;
}

/** В списке может не оказаться значения из старой задачи — тогда добавляем его, чтобы не потерять. */
function fillSelect(el, items, value) {
  const all = value && !items.includes(value) ? items.concat([value]) : items;
  el.innerHTML = '<option value="">—</option>' +
    all.map((v) => `<option value="${esc(v)}"${v === value ? ' selected' : ''}>${esc(v)}</option>`).join('');
}

function renderLangs(checked) {
  const groups = [['Основные', lists.languages.regular], ['ШТАТ', lists.languages.shtat], ['Редкие', lists.languages.rare]];
  const known = groups.flatMap((g) => g[1]);
  const other = checked.filter((l) => !known.includes(l));
  if (other.length) groups.push(['Другие', other]);
  els.langs.innerHTML = groups.filter((g) => g[1].length).map(([title, items]) =>
    `<div class="lang-group"><span class="lang-title">${title}</span><div class="chips">` +
    items.map((l) => `<label class="chip"><input type="checkbox" value="${esc(l)}"${checked.includes(l) ? ' checked' : ''}>` +
      `<span>${esc(l.replace(/^ШТАТ /, ''))}</span></label>`).join('') + '</div></div>').join('');
}

makeLinkList(els.link2);

function close() {
  els.form.hidden = true;
  els.view.hidden = false;
  current = null;
  window.scrollTo(0, 0); // форма длинная — список и сообщение наверху
}
els.back.addEventListener('click', close);

/**
 * Открыть задачу на правку. ref — { row, id, origSubject } из списка «Мои задачи» или из поиска.
 * done() вызывается после сохранения — обновить список, del(ref) — после подтверждения удаления.
 */
export async function openEditor(ref, done, del) {
  onDone = done || (() => {});
  onDelete = del || (() => {});
  els.view.hidden = true;
  els.form.hidden = false;
  els.form.reset();
  els.link2.value = '';
  els.title.textContent = ref.id || 'Задача';
  els.save.disabled = true;
  els.del.disabled = true;
  els.langs.innerHTML = '';
  // «Загружаю задачу…» — сверху, где смотрят, а не под кнопкой
  els.top.hidden = false;
  els.top.className = 'status';
  els.top.textContent = 'Загружаю задачу…';
  setMsg('');
  window.scrollTo(0, 0);
  try {
    // Справочники — из памяти, задача — одним запросом, параллельно
    const [l, r] = await Promise.all([
      lists ? lists : getLists((fresh) => { lists = fresh; }),
      api({ action: 'getTask', row: ref.row, id: ref.id || '', origSubject: ref.origSubject }),
    ]);
    lists = l;
    if (!r.ok) throw new Error(r.error);
    fill(r.task);
    els.save.disabled = false;
    els.del.disabled = false;
    els.top.hidden = true;
  } catch (e) {
    els.top.className = 'status err';
    els.top.textContent = `Не получилось открыть: ${e.message}`;
  }
}

function fill(t) {
  current = t;
  els.title.textContent = [t.id || 'без номера', t.contractor].filter(Boolean).join(' · ');
  els.subject.value = t.subject;
  els.link.value = t.link;
  els.link2.value = t.link2;
  els.ticketNum.value = String(t.ticket || '').replace(/^LOCAL-/i, '');
  els.date.value = t.date;
  els.customer.value = t.customer;
  els.exactDeadline.value = t.exactDeadline;
  els.estimateLink.value = t.estimateLink;
  els.total.value = t.total;
  els.sp.value = t.sp;
  els.comment.value = t.comment;
  els.complaints.value = t.complaints || '';
  fillSelect(els.product, lists.products, t.product);
  fillSelect(els.deadline, lists.deadlines, t.deadline);
  fillSelect(els.status, lists.statuses, t.status);
  els.status.dataset.status = els.status.value;
  fillSelect(els.deliveryStatus, lists.deliveryStatuses, t.deliveryStatus);
  els.deliveryStatus.dataset.delivery = els.deliveryStatus.value;
  fillSelect(els.contractor, lists.contractors, t.contractor);
  fillSelect(els.manager, lists.managers, t.manager);
  renderLangs(String(t.languages || '').split(',').map((s) => s.trim()).filter(Boolean));
}

/** Языки в прежнем порядке, новые — в конце, как в таблице. */
function keepOrder(before, checked) {
  return before.filter((l) => checked.includes(l)).concat(checked.filter((l) => !before.includes(l)));
}

els.form.addEventListener('submit', async (e) => {
  e.preventDefault();
  if (!current) return;
  if (!els.subject.value.trim()) return setMsg('Тема не может быть пустой', 'err');
  const ticketNum = els.ticketNum.value.trim().replace(/^local-?\s*/i, '');
  if (ticketNum && !/^\d+$/.test(ticketNum)) return setMsg('В тикете — только цифры, LOCAL- подставится сам', 'err');
  const task = {
    row: current.row, id: current.id, origSubject: current.origSubject,
    subject: els.subject.value.trim(), link: els.link.value.trim(),
    link2: els.link2.value.split(/[\s,;]+/).filter(Boolean).join('\n'),
    ticket: ticketNum ? `LOCAL-${ticketNum}` : '', date: els.date.value, product: els.product.value,
    customer: els.customer.value.trim(), deadline: els.deadline.value, exactDeadline: els.exactDeadline.value,
    status: els.status.value, deliveryStatus: els.deliveryStatus.value, contractor: els.contractor.value,
    manager: els.manager.value, estimateLink: els.estimateLink.value.trim(), total: els.total.value,
    sp: els.sp.value, comment: els.comment.value.trim(), complaints: els.complaints.value.trim(),
    languages: keepOrder(String(current.languages || '').split(',').map((x) => x.trim()).filter(Boolean),
      [...els.langs.querySelectorAll('input:checked')].map((cb) => cb.value)).join(', '),
  };
  els.save.disabled = true;
  setMsg('Сохраняю…');
  try {
    const r = await api({ action: 'saveTask', task });
    if (!r.ok) throw new Error(r.error);
    close();
    onDone(`${task.id || 'Задача'}: сохранено ✓`);
  } catch (err) {
    setMsg(`Не сохранилось: ${err.message}`, 'err');
  } finally {
    els.save.disabled = false;
  }
});

// Удаление — как в «Поиске и правке» таблицы: с подтверждением, отменить нельзя.
// После «ОК» окно сразу закрывается, задача пропадает из списка, а таблица удаляет строку в фоне (см. mine.js).
els.del.addEventListener('click', () => {
  if (!current) return;
  const name = current.id ? `№ ${current.id}` : `«${current.subject}»`;
  if (!confirm(`Точно удалить задачу ${name}?\n\nСтрока пропадёт из таблицы, отменить нельзя.`)) return;
  const ref = { row: current.row, id: current.id || '', origSubject: current.origSubject };
  close();
  onDelete(ref);
});
