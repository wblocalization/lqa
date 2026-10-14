// Правка задачи из расширения: то же, что «Поиск и правка» в таблице.
import { api, esc, getLists, autoDeadline, ask, sumField, langGroupsHtml, bindLangCounts, updateLangCounts } from './core.js';
import { makeLinkList } from './links.js';
import { openLetterFor } from './task.js';

const $ = (s) => document.querySelector(s);
const els = {
  view: $('#mineView'), form: $('#editForm'), back: $('#eBack'), title: $('#eTitle'), letter: $('#eLetter'),
  subject: $('#eSubject'), link: $('#eLink'), link2: $('#eLink2'), ticketNum: $('#eTicketNum'), date: $('#eDate'),
  product: $('#eProduct'), customer: $('#eCustomer'), langs: $('#eLangs'), deadline: $('#eDeadline'),
  exactDeadline: $('#eExactDeadline'), status: $('#eStatus'), deliveryStatus: $('#eDeliveryStatus'),
  contractor: $('#eContractor'), manager: $('#eManager'), estimateLink: $('#eEstimateLink'),
  total: $('#eTotal'), sp: $('#eSp'), comment: $('#eComment'), complaints: $('#eComplaints'), save: $('#eSave'), del: $('#eDelete'), msg: $('#eMsg'), top: $('#eTop'),
};
// «Итого»: «5263+5336» или текст письма с суммами → в таблицу формулой
const totalSum = sumField(els.total);
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
  const known = ['regular', 'shtat', 'rare'].flatMap((g) => lists.languages[g] || []);
  els.langs.innerHTML = langGroupsHtml(lists.languages, checked, checked.filter((l) => !known.includes(l)));
  bindLangCounts(els.langs);
  updateLangCounts(els.langs);
}

makeLinkList(els.link2);

// Где был список, когда открыли задачу: после «Назад», сохранения или удаления возвращаемся туда же,
// а не наверх — удобно идти по списку дальше (например, удалять тестовые задачи подряд)
let returnY = 0;
function close() {
  els.form.hidden = true;
  els.view.hidden = false;
  current = null;
  requestAnimationFrame(() => window.scrollTo(0, returnY));
}
els.back.addEventListener('click', close);
// 📨 Письмо подрядчику по этой задаче (как после «Добавить»); несохранённые правки формы в письмо не попадут
els.letter.addEventListener('click', () => {
  if (!current) return;
  const t = { ...current };
  close();
  openLetterFor(t);
});

/**
 * Открыть задачу на правку. ref — { row, id, origSubject } из списка «Мои задачи» или из поиска.
 * done() вызывается после сохранения — обновить список, del(ref) — после подтверждения удаления.
 */
export async function openEditor(ref, done, del) {
  onDone = done || (() => {});
  onDelete = del || (() => {});
  returnY = window.scrollY;
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
  els.total.value = t.totalFormula || t.total;
  totalSum.show();
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
  let total;
  try { total = totalSum.read(); } catch (err) { return setMsg(err.message, 'err'); }
  const ticketNum = els.ticketNum.value.trim().replace(/^local-?\s*/i, '');
  if (ticketNum && !/^\d+$/.test(ticketNum)) return setMsg('В тикете — только цифры, LOCAL- подставится сам', 'err');
  const task = {
    row: current.row, id: current.id, origSubject: current.origSubject,
    subject: els.subject.value.trim(), link: els.link.value.trim(),
    link2: els.link2.value.split(/[\s,;]+/).filter(Boolean).join('\n'),
    ticket: ticketNum ? `LOCAL-${ticketNum}` : '', date: els.date.value, product: els.product.value,
    customer: els.customer.value.trim(), deadline: els.deadline.value, exactDeadline: els.exactDeadline.value,
    status: els.status.value, deliveryStatus: els.deliveryStatus.value, contractor: els.contractor.value,
    manager: els.manager.value, estimateLink: els.estimateLink.value.trim(), total,
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
els.del.addEventListener('click', async () => {
  if (!current) return;
  const name = current.id ? `№ ${current.id}` : `«${current.subject}»`;
  if (!await ask({ title: `Удалить задачу ${name}?`, text: 'Строка пропадёт из таблицы, отменить нельзя.', ok: 'Удалить', danger: true })) return;
  const ref = { row: current.row, id: current.id || '', origSubject: current.origSubject };
  close();
  onDelete(ref);
});
