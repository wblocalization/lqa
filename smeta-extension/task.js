// Вкладка «Задача»: то же, что окно «➕ Добавить задачу» в таблице, только без захода в таблицу.
import { settings, saveSettings, isConfigured, api, esc, readClipboard } from './core.js';

const $ = (s) => document.querySelector(s);
const els = {
  form: $('#taskForm'), loading: $('#taskLoading'), done: $('#taskDone'), doneId: $('#taskDoneId'),
  doneSubject: $('#taskDoneSubject'), copyDone: $('#taskCopyDone'), again: $('#taskAgain'),
  contractor: $('#tContractor'), ticketNum: $('#tTicketNum'), template: $('#tTemplate'),
  subject: $('#tSubject'), link: $('#tLink'), pasteLink: $('#tPasteLink'),
  product: $('#tProduct'), date: $('#tDate'), deadline: $('#tDeadline'), exactDeadline: $('#tExactDeadline'),
  customer: $('#tCustomer'), langs: $('#tLangs'), manager: $('#tManager'), status: $('#tStatus'),
  link2: $('#tLink2'), deliveryStatus: $('#tDeliveryStatus'), estimateLink: $('#tEstimateLink'),
  total: $('#tTotal'), sp: $('#tSp'), comment: $('#tComment'),
  preview: $('#tPreview'), copyPreview: $('#tCopyPreview'), submit: $('#tSubmit'), msg: $('#tMsg'),
};

let lists = null;       // справочники из таблицы
let nextId = '';        // номер, который получит задача (подсказка)
let lastSubject = '';   // тема добавленной задачи — для «Скопировать тему»

// Те же быстрые шаблоны, что в окне «➕ Добавить задачу».
const TEMPLATES = {
  push: {
    label: 'Пуш-рассылка (маркетинг)',
    product: 'Магазинка Пуши и коммуникации: Маркетинг',
    langs: ['Армянский', 'Грузинский', 'Казахский', 'Кыргызский', 'Таджикский', 'Узбекский'],
  },
  web: {
    label: 'Новые строки для веба',
    product: 'Магазинка', contractor: 'LogrusIT', customer: '@coy.elena3', manager: 'Анастасия Лисовая', ticketNum: '1114',
    langs: ['Амхарский', 'Армянский', 'Грузинский', 'Казахский', 'Кыргызский', 'Таджикский', 'Узбекский', 'ШТАТ английский'],
    subject: 'Новые строчки для веба от {date}',
  },
  app: {
    label: 'Новые строки для приложения',
    product: 'Магазинка', contractor: 'LogrusIT', customer: '@syrcov.evgeniy, @arslanov.anton', manager: 'Анастасия Лисовая', ticketNum: '1116',
    langs: ['Азербайджанский', 'Армянский', 'Грузинский', 'Казахский', 'Кыргызский', 'Таджикский', 'Узбекский', 'ШТАТ английский'],
    subject: 'Новые строчки для приложения от {date}',
  },
  seller: {
    label: 'Портал продавца (поставки/аналитика/FBS)',
    product: 'Портал продавца',
    langs: ['ШТАТ английский', 'ШТАТ грузинский', 'ШТАТ китайский', 'ШТАТ узбекский', 'ШТАТ армянский', 'Казахский', 'Таджикский'],
  },
};

// ---------- Сообщения ----------
function setMsg(text, kind = 'info') {
  els.msg.textContent = text;
  els.msg.className = `status ${kind}`;
}

// ---------- Загрузка справочников ----------
export async function initTaskTab() {
  if (lists) return;
  if (!isConfigured()) {
    els.loading.hidden = false;
    els.loading.textContent = 'Заполните настройки (⚙️), чтобы добавлять задачи.';
    return;
  }
  els.loading.hidden = false;
  els.loading.textContent = 'Загружаю справочники из таблицы…';
  try {
    const r = await api({ action: 'taskForm' });
    if (!r.ok) throw new Error(r.error);
    lists = r.lists;
    buildForm();
    els.loading.hidden = true;
    els.form.hidden = false;
  } catch (e) {
    els.loading.textContent = `Не получилось загрузить справочники: ${e.message}`;
  }
}

/** Настройки поменялись — перечитать справочники. */
export function onSettingsSaved() {
  lists = null;
  els.form.hidden = true;
  initTaskTab();
}

function fillSelect(el, items, placeholder = '—') {
  el.innerHTML = `<option value="">${esc(placeholder)}</option>` +
    items.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
}

function buildForm() {
  fillSelect(els.contractor, lists.contractors, 'Выберите…');
  fillSelect(els.product, lists.products);
  fillSelect(els.deadline, lists.deadlines);
  fillSelect(els.status, lists.statuses);
  fillSelect(els.deliveryStatus, lists.deliveryStatuses);
  fillSelect(els.manager, lists.managers);
  els.template.innerHTML = '<option value="">— без шаблона —</option>' +
    Object.entries(TEMPLATES).map(([k, t]) => `<option value="${k}">${esc(t.label)}</option>`).join('');

  const titles = { regular: 'Основные', shtat: 'ШТАТ', rare: 'Редкие' };
  els.langs.innerHTML = ['regular', 'shtat', 'rare']
    .filter((g) => (lists.languages[g] || []).length)
    .map((g) => `<div class="lang-group"><span class="lang-title">${titles[g]}</span><div class="chips">` +
      lists.languages[g].map((l) =>
        `<label class="chip"><input type="checkbox" value="${esc(l)}"><span>${esc(l.replace(/^ШТАТ /, ''))}</span></label>`).join('') +
      '</div></div>')
    .join('');

  resetForm();
}

function resetForm() {
  els.form.reset();
  els.date.value = today();
  if (settings.manager && lists.managers.includes(settings.manager)) els.manager.value = settings.manager;
  nextId = '';
  updatePreview();
}

function today() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

// ---------- Шаблоны ----------
els.template.addEventListener('change', () => {
  const t = TEMPLATES[els.template.value];
  if (!t) return;
  els.product.value = t.product || '';
  checkedLangInputs(false).forEach((cb) => { cb.checked = t.langs.includes(cb.value); });
  if (t.contractor) { els.contractor.value = t.contractor; refreshNextId(); }
  if (t.ticketNum) els.ticketNum.value = t.ticketNum;
  if (t.manager) els.manager.value = t.manager;
  if (t.customer) els.customer.value = t.customer;
  if (t.subject) {
    const [, m, d] = els.date.value.split('-');
    els.subject.value = t.subject.replace('{date}', d && m ? `${d}.${m}` : '');
  }
  updatePreview();
});

// ---------- Номер и превью темы ----------
let idSeq = 0;
async function refreshNextId() {
  const contractor = els.contractor.value;
  const seq = ++idSeq;
  nextId = '';
  updatePreview();
  if (!contractor) return;
  try {
    const r = await api({ action: 'previewTaskId', contractor });
    if (seq !== idSeq) return;
    if (!r.ok) throw new Error(r.error);
    nextId = r.id;
  } catch (e) {
    if (seq === idSeq) setMsg(`Не получилось узнать номер: ${e.message}`, 'err');
  }
  updatePreview();
}

function checkedLangInputs(onlyChecked = true) {
  return [...els.langs.querySelectorAll(`input[type=checkbox]${onlyChecked ? ':checked' : ''}`)];
}

/** Тема так же, как её собирает таблица: [номер][подрядчик][коды языков][продукт] тема. */
function buildSubject(id) {
  const contractor = els.contractor.value;
  const product = els.product.value;
  const codes = checkedLangInputs().map((cb) => lists.langCodes[cb.value]).filter(Boolean);
  const prefix = `[${id}]` + (contractor ? `[${contractor}]` : '') + codes.map((c) => `[${c}]`).join('') +
    (product ? `[${product}]` : '');
  const subject = els.subject.value.trim();
  return subject ? `${prefix} ${subject}` : prefix;
}

function updatePreview() {
  if (!lists) return;
  if (!els.contractor.value) { els.preview.textContent = 'Выберите подрядчика — появится номер и тема'; return; }
  els.preview.textContent = buildSubject(nextId || '…');
}

els.contractor.addEventListener('change', refreshNextId);
[els.subject, els.product].forEach((el) => el.addEventListener('input', updatePreview));
els.product.addEventListener('change', updatePreview);
els.langs.addEventListener('change', updatePreview);

async function copyText(text, okMsg) {
  try {
    await navigator.clipboard.writeText(text);
    setMsg(okMsg, 'ok');
  } catch {
    setMsg('Не получилось скопировать', 'err');
  }
}
els.copyPreview.addEventListener('click', () => copyText(els.preview.textContent, 'Тема скопирована'));

els.pasteLink.addEventListener('click', async () => {
  try {
    els.link.value = await readClipboard();
  } catch {
    setMsg('Нет доступа к буферу — вставьте ссылку вручную (Ctrl+V)', 'err');
  }
});

// ---------- Добавление ----------
els.form.addEventListener('submit', async (e) => {
  e.preventDefault();
  if (!els.contractor.value) return setMsg('Выберите подрядчика', 'err');
  if (!els.subject.value.trim()) return setMsg('Впишите тему', 'err');
  const ticketNum = els.ticketNum.value.trim();
  if (ticketNum && !/^\d+$/.test(ticketNum)) return setMsg('В тикете — только цифры, LOCAL- подставится сам', 'err');

  const task = {
    ticket: ticketNum ? `LOCAL-${ticketNum}` : '',
    contractor: els.contractor.value, subject: els.subject.value.trim(),
    link: els.link.value.trim(), link2: els.link2.value.trim(),
    date: els.date.value, product: els.product.value, customer: els.customer.value.trim(),
    deadline: els.deadline.value, exactDeadline: els.exactDeadline.value,
    status: els.status.value, deliveryStatus: els.deliveryStatus.value,
    estimateLink: els.estimateLink.value.trim(), total: els.total.value, sp: els.sp.value,
    manager: els.manager.value, comment: els.comment.value.trim(),
    languages: checkedLangInputs().map((cb) => cb.value),
  };

  els.submit.disabled = true;
  setMsg('Добавляю…');
  try {
    const r = await api({ action: 'addTask', task });
    if (!r.ok) throw new Error(r.error);
    if (task.manager) await saveSettings({ manager: task.manager }); // в следующий раз менеджер подставится сам
    lastSubject = buildSubject(r.id);
    els.doneId.textContent = r.id;
    els.doneSubject.textContent = lastSubject;
    els.form.hidden = true;
    els.done.hidden = false;
    setMsg('');
  } catch (err) {
    setMsg(`Не добавилось: ${err.message}`, 'err');
  } finally {
    els.submit.disabled = false;
  }
});

els.copyDone.addEventListener('click', () => copyText(lastSubject, 'Тема скопирована — можно вставлять в письмо'));
els.again.addEventListener('click', () => {
  els.done.hidden = true;
  els.form.hidden = false;
  resetForm();
  setMsg('');
});
