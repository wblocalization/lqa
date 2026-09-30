// Вкладка «Задача»: то же, что окно «➕ Добавить задачу» в таблице, только без захода в таблицу.
import { settings, saveSettings, isConfigured, api, esc, readClipboard } from './core.js';

const $ = (s) => document.querySelector(s);
const els = {
  form: $('#taskForm'), loading: $('#taskLoading'), done: $('#taskDone'), doneId: $('#taskDoneId'),
  doneSubject: $('#taskDoneSubject'), copyDone: $('#taskCopyDone'), again: $('#taskAgain'),
  contractor: $('#tContractor'), ticketNum: $('#tTicketNum'), repeat: $('#tRepeat'),
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

let recent = [];        // последние задачи — для «Повторить задачу»

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

  const titles = { regular: 'Основные', shtat: 'ШТАТ', rare: 'Редкие' };
  els.langs.innerHTML = ['regular', 'shtat', 'rare']
    .filter((g) => (lists.languages[g] || []).length)
    .map((g) => `<div class="lang-group"><span class="lang-title">${titles[g]}</span><div class="chips">` +
      lists.languages[g].map((l) =>
        `<label class="chip"><input type="checkbox" value="${esc(l)}"><span>${esc(l.replace(/^ШТАТ /, ''))}</span></label>`).join('') +
      '</div></div>')
    .join('');

  resetForm();
  loadRecent();
}

/** Последние задачи менеджера для «Повторить задачу». */
async function loadRecent() {
  try {
    const r = await api({ action: 'recentTasks', manager: settings.manager || lists.currentManager || '' });
    if (!r.ok) throw new Error(r.error);
    recent = r.tasks || [];
  } catch {
    recent = [];
  }
  els.repeat.innerHTML = '<option value="">— новая задача с нуля —</option>' +
    recent.map((t, i) => `<option value="${i}">${esc(t.title)} · ${esc(t.date)}${t.ticket ? ' · ' + esc(t.ticket) : ''}</option>`).join('');
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

// ---------- Повторить задачу ----------
/** «Новые строчки от 22.09» → «Новые строчки от <сегодня>». */
function withToday(title) {
  const [, m, d] = els.date.value.split('-');
  // \b не работает с русскими буквами — границу слова задаём явно
  return d && m ? title.replace(/(^|\s)(от\s+)\d{1,2}\.\d{1,2}(\.\d{2,4})?/i, `$1$2${d}.${m}`) : title;
}

els.repeat.addEventListener('change', () => {
  const t = recent[els.repeat.value];
  if (!t) return;
  els.contractor.value = t.contractor;
  refreshNextId();
  els.ticketNum.value = t.ticket.replace(/^LOCAL-/i, '');
  els.product.value = t.product;
  els.customer.value = t.customer;
  els.manager.value = t.manager;
  els.deadline.value = t.deadline;
  els.subject.value = withToday(t.title);
  els.link.value = '';
  checkedLangInputs(false).forEach((cb) => { cb.checked = t.languages.includes(cb.value); });
  setMsg(`Заполнено по ${t.id || 'прошлой задаче'}. Вставьте новую ссылку на Band и проверьте тему.`);
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
els.copyPreview.addEventListener('click', async () => {
  await copyText(els.preview.textContent, 'Тема скопирована');
  els.copyPreview.textContent = 'Скопировано ✓';
  setTimeout(() => { els.copyPreview.textContent = 'Скопировать'; }, 1500);
});

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
  setMsg('Проверяю, нет ли такой задачи…');
  try {
    const d = await api({ action: 'checkDuplicates', task });
    if (d.ok && d.duplicates && d.duplicates.length && !confirm('Похожая задача уже есть:\n\n' +
        d.duplicates.map((x) => `${x.id || '—'} · ${x.title} · ${x.date} · ${x.manager} (${x.why})`).join('\n') +
        '\n\nВсё равно добавить?')) {
      setMsg('Не добавлено — похожая задача уже есть.', 'err');
      return;
    }
    setMsg('Добавляю…');
    const r = await api({ action: 'addTask', task });
    if (!r.ok) throw new Error(r.error);
    if (task.manager) await saveSettings({ manager: task.manager }); // в следующий раз менеджер подставится сам
    loadRecent();
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
