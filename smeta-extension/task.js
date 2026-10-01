// Вкладка «Задача»: то же, что окно «➕ Добавить задачу» в таблице, только без захода в таблицу.
import { settings, isConfigured, api, esc, readClipboard } from './core.js';

const $ = (s) => document.querySelector(s);
const els = {
  form: $('#taskForm'), loading: $('#taskLoading'), done: $('#taskDone'), doneId: $('#taskDoneId'),
  doneSubject: $('#taskDoneSubject'), copyDone: $('#taskCopyDone'), again: $('#taskAgain'),
  contractor: $('#tContractor'), ticketNum: $('#tTicketNum'),
  template: $('#tTemplate'), saveTpl: $('#tSaveTpl'), delTpl: $('#tDelTpl'), exportTpl: $('#tExportTpl'),
  importTpl: $('#tImportTpl'), importFile: $('#tImportFile'),
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
  loadTemplates();
}

function resetForm() {
  const tpl = els.template.value;
  els.form.reset();
  els.template.value = tpl; // выбранный шаблон остаётся выбранным, но поля — с нуля
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
// Свои у каждого (хранятся в Chrome). Файлом можно поделиться: тот же формат понимает окно «Новая задача» в таблице.
const TPL_FILE_TYPE = 'wb-task-templates';
let templates = [];

async function loadTemplates() {
  templates = (await chrome.storage.local.get('templates')).templates || [];
  renderTemplates();
}

async function storeTemplates(selectName = '') {
  await chrome.storage.local.set({ templates });
  renderTemplates(selectName);
}

function renderTemplates(selectName = '') {
  templates.sort((a, b) => a.name.localeCompare(b.name, 'ru'));
  els.template.innerHTML = '<option value="">— без шаблона —</option>' +
    templates.map((t, i) => `<option value="${i}">${esc(t.name)}</option>`).join('');
  const i = templates.findIndex((t) => t.name === selectName);
  els.template.value = i === -1 ? '' : String(i);
  els.delTpl.hidden = !els.template.value;
}

/** Только известные поля и только строки: файл мог прийти от кого угодно. */
function cleanTemplate(t) {
  const str = (v) => (typeof v === 'string' ? v.trim() : '');
  const out = { name: str(t && t.name) };
  ['contractor', 'ticket', 'subject', 'product', 'customer', 'deadline', 'comment'].forEach((k) => { out[k] = str(t[k]); });
  out.languages = Array.isArray(t.languages) ? t.languages.filter((l) => typeof l === 'string') : [];
  return out;
}

/** «Новые строчки от 22.09» → «Новые строчки от <сегодня>». */
function withToday(title) {
  const [, m, d] = els.date.value.split('-');
  // \b не работает с русскими буквами — границу слова задаём явно
  return d && m ? title.replace(/(^|\s)(от\s+)\d{1,2}\.\d{1,2}(\.\d{2,4})?/i, `$1$2${d}.${m}`) : title;
}

els.template.addEventListener('change', () => {
  els.delTpl.hidden = !els.template.value;
  const t = templates[els.template.value];
  if (!t) return;
  els.contractor.value = t.contractor;
  refreshNextId();
  els.ticketNum.value = t.ticket.replace(/^LOCAL-/i, '');
  els.product.value = t.product;
  els.customer.value = t.customer;
  els.deadline.value = t.deadline;
  els.comment.value = t.comment;
  els.subject.value = withToday(t.subject);
  els.link.value = '';
  checkedLangInputs(false).forEach((cb) => { cb.checked = t.languages.includes(cb.value); });
  setMsg(`Заполнено по шаблону «${t.name}». Вставьте ссылку на Band и проверьте тему.`);
  updatePreview();
});

els.saveTpl.addEventListener('click', async () => {
  if (!els.contractor.value && !els.subject.value.trim()) return setMsg('Заполните хотя бы подрядчика или тему — их и запомнит шаблон', 'err');
  const current = templates[els.template.value];
  const name = (prompt('Название шаблона:', current ? current.name : els.subject.value.trim()) || '').trim();
  if (!name) return;
  const existing = templates.findIndex((t) => t.name === name);
  if (existing !== -1 && !confirm(`Шаблон «${name}» уже есть. Заменить?`)) return;
  const ticketNum = els.ticketNum.value.trim();
  const tpl = cleanTemplate({
    name, contractor: els.contractor.value, ticket: ticketNum ? `LOCAL-${ticketNum}` : '', subject: els.subject.value,
    product: els.product.value, customer: els.customer.value, deadline: els.deadline.value, comment: els.comment.value,
    languages: checkedLangInputs().map((cb) => cb.value),
  });
  if (existing !== -1) templates[existing] = tpl; else templates.push(tpl);
  await storeTemplates(name);
  setMsg(`Шаблон «${name}» сохранён`, 'ok');
});

els.delTpl.addEventListener('click', async () => {
  const t = templates[els.template.value];
  if (!t || !confirm(`Удалить шаблон «${t.name}»?`)) return;
  templates.splice(Number(els.template.value), 1);
  await storeTemplates();
  setMsg(`Шаблон «${t.name}» удалён`);
});

els.exportTpl.addEventListener('click', () => {
  if (!templates.length) return setMsg('Шаблонов пока нет: заполните форму и нажмите «Сохранить» рядом с «Шаблон»', 'err');
  const blob = new Blob([JSON.stringify({ type: TPL_FILE_TYPE, version: 1, templates }, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'task-templates.json';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  setMsg(`Скачано шаблонов: ${templates.length}. Файл можно отправить коллеге — пусть нажмёт «Загрузить из файла».`, 'ok');
});

els.importTpl.addEventListener('click', () => els.importFile.click());
els.importFile.addEventListener('change', async () => {
  const file = els.importFile.files[0];
  els.importFile.value = '';
  if (!file) return;
  try {
    const data = JSON.parse(await file.text());
    const list = (Array.isArray(data) ? data : data.templates || []).map(cleanTemplate).filter((t) => t.name);
    if (!list.length) throw new Error('в файле нет шаблонов');
    let replaced = 0;
    list.forEach((t) => {
      const i = templates.findIndex((x) => x.name === t.name);
      if (i !== -1) { templates[i] = t; replaced++; } else templates.push(t);
    });
    await storeTemplates();
    setMsg(`Загружено шаблонов: ${list.length}` + (replaced ? ` (заменено с тем же названием: ${replaced})` : ''), 'ok');
  } catch (e) {
    setMsg(`Не получилось загрузить: ${e.message}`, 'err');
  }
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
