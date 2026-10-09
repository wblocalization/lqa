// Вкладка «Новая задача»: то же, что окно «➕ Добавить задачу» в таблице, только без захода в таблицу.
import { settings, isConfigured, api, esc, readClipboard, getLists, autoDeadline, setLink, toast } from './core.js';
import { makeLinkList } from './links.js';
import { showLetter, lettersForExport, importLetters, loadLetters } from './letter.js';

const $ = (s) => document.querySelector(s);
const els = {
  form: $('#taskForm'), loading: $('#taskLoading'), done: $('#taskDone'), doneId: $('#taskDoneId'),
  doneSubject: $('#taskDoneSubject'), doneCopied: $('#taskDoneCopied'), copyDone: $('#taskCopyDone'), again: $('#taskAgain'),
  contractor: $('#tContractor'), ticketNum: $('#tTicketNum'),
  template: $('#tTemplate'), saveTpl: $('#tSaveTpl'), delTpl: $('#tDelTpl'), updTpl: $('#tUpdTpl'), renTpl: $('#tRenTpl'), exportTpl: $('#tExportTpl'),
  importTpl: $('#tImportTpl'), importFile: $('#tImportFile'),
  subject: $('#tSubject'), link: $('#tLink'), pasteLink: $('#tPasteLink'),
  product: $('#tProduct'), date: $('#tDate'), deadline: $('#tDeadline'), exactDeadline: $('#tExactDeadline'),
  customer: $('#tCustomer'), langs: $('#tLangs'), manager: $('#tManager'), status: $('#tStatus'),
  link2: $('#tLink2'), deliveryStatus: $('#tDeliveryStatus'), estimateLink: $('#tEstimateLink'),
  total: $('#tTotal'), sp: $('#tSp'), comment: $('#tComment'), complaints: $('#tComplaints'),
  preview: $('#tPreview'), copyPreview: $('#tCopyPreview'), submit: $('#tSubmit'), msg: $('#tMsg'),
  tplChips: $('#tTplChips'), suggest: $('#tSuggest'), suggestText: $('#tSuggestText'), suggestApply: $('#tSuggestApply'),
  suggestClose: $('#tSuggestClose'), clearTpl: $('#tClearTpl'), tplIoMsg: $('#tTplIoMsg'), openRow: $('#taskOpenRow'),
};
autoDeadline(els.deadline, els.exactDeadline, els.date);
loadLetters().catch(() => {}); // для «Скачать файлом» в настройках

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
  if (lists) { refreshAutoDate(); return; }
  if (!isConfigured()) {
    els.loading.hidden = false;
    els.loading.textContent = 'Заполните настройки (⚙️), чтобы добавлять задачи.';
    return;
  }
  els.loading.hidden = false;
  els.loading.textContent = 'Загружаю справочники из таблицы…';
  try {
    // Справочники — из памяти (мгновенно); поменялись в таблице — перестроим форму, если в ней ещё ничего не начали
    lists = await getLists((fresh) => {
      lists = fresh;
      if (!els.contractor.value && !els.subject.value.trim()) buildForm();
    });
    buildForm();
    loadNextIds();
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

  const titles = { regular: 'Подрядчики', shtat: 'Штат', rare: 'Подрядчики · редкие' };
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

let fromLast = false; // форму заполнили «Как в прошлый раз» — тоже можно «Сбросить»
function resetForm() {
  fromLast = false;
  els.form.reset();
  document.querySelector('#tWeekend').hidden = true;
  els.link2.value = '';
  els.template.value = ''; // форма с нуля — шаблон снова можно выбрать кнопкой
  els.delTpl.hidden = true;
  renderChips();
  hideSuggest();
  els.date.value = autoDate = today();
  updateOtHint();
  if (settings.manager && lists.managers.includes(settings.manager)) els.manager.value = settings.manager;
  if (lists.statuses.includes('Принято')) els.status.value = 'Принято'; // новая задача — сразу «Принято»
  els.status.dataset.status = els.status.value;
  nextId = '';
  updatePreview();
}

// Дата получения — сегодня. Панель могла быть открыта со вчера: если дату не меняли руками, обновим её.
let autoDate = '';
function refreshAutoDate() {
  if (els.date.value && els.date.value !== autoDate) return; // поставили свою — не трогаем
  if (els.date.value === today()) return;
  els.date.value = autoDate = today();
  els.subject.value = withToday(els.subject.value);
  updateOtHint();
}

document.addEventListener('visibilitychange', () => { if (!document.hidden && lists) refreshAutoDate(); });

function today() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

// ---------- Шаблоны ----------
// Свои у каждого (хранятся в Chrome). Файлом можно поделиться: тот же формат понимает окно «Новая задача» в таблице.
const TPL_FILE_TYPE = 'wb-task-templates';
let templates = [];
let lastTask = null;     // последняя добавленная задача — «Как в прошлый раз»
let byContractor = {};   // последние поля по каждому подрядчику — для подсказки «как обычно»

async function loadTemplates() {
  const saved = await chrome.storage.local.get(['templates', 'lastTask', 'byContractor']);
  templates = saved.templates || [];
  lastTask = saved.lastTask ? cleanTemplate(saved.lastTask) : null;
  byContractor = saved.byContractor || {};
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
  renderChips();
}

/** Шаблоны — кнопками: один клик, и форма заполнена. Пунктиром — «Как в прошлый раз». */
function renderChips() {
  const sel = els.template.value;
  const chips = templates.map((t, i) =>
    `<button type="button" class="tpl-chip${String(i) === sel ? ' on' : ''}" data-tpl="${i}" title="${esc(t.name)}">${esc(t.name)}</button>`);
  if (lastTask) chips.push(`<button type="button" class="tpl-chip last" data-last="1" title="${esc(describe(lastTask))}">↻ Как в прошлый раз</button>`);
  els.tplChips.innerHTML = chips.length ? chips.join('')
    : '<span class="tpl-empty">Шаблонов пока нет — заполните форму и нажмите «Сохранить как шаблон»</span>';
  els.delTpl.textContent = templates[sel] ? `Удалить шаблон «${templates[sel].name}»` : 'Удалить шаблон';
  // Выбран шаблон — его можно поправить: поменяли в форме тикет, языки… → «Сохранить изменения»
  const on = Boolean(templates[sel]);
  els.delTpl.hidden = els.updTpl.hidden = els.renTpl.hidden = !on;
  els.clearTpl.hidden = !(on || fromLast);
  els.updTpl.textContent = on ? `💾 Сохранить изменения в «${templates[sel].name}»` : '💾 Сохранить изменения';
  els.saveTpl.textContent = on ? '＋ Сохранить как новый' : '＋ Сохранить как шаблон';
}

els.tplChips.addEventListener('click', (e) => {
  const b = e.target.closest('button');
  if (!b) return;
  if (b.dataset.last) {
    els.template.value = '';
    els.delTpl.hidden = true;
    applyFields(lastTask);
    fromLast = true;
    renderChips();
    setMsg('Заполнено как прошлая задача. Проверьте тему и, если есть, вставьте ссылку на Band.');
    return;
  }
  els.template.value = b.dataset.tpl;
  els.template.dispatchEvent(new Event('change'));
});

/** «LogrusIT · Магазинка · kk, ka, hy» — коротко, что подставится. */
function describe(t) {
  const codes = (t.languages || []).map((l) => (lists && lists.langCodes[l]) || l);
  return [t.contractor, t.product, codes.join(', ')].filter(Boolean).join(' · ');
}

let applying = false; // форму заполняет скрипт — подсказку «как обычно» не показываем
/** Заполнить форму из шаблона / прошлой задачи. keepSubject — не трогать тему, если её уже вписали. */
function applyFields(t, { keepSubject = false } = {}) {
  if (!t) return;
  applying = true;
  els.contractor.value = t.contractor || '';
  refreshNextId();
  els.ticketNum.value = String(t.ticket || '').replace(/^LOCAL-/i, '');
  els.product.value = t.product || '';
  els.customer.value = t.customer || '';
  els.deadline.value = t.deadline || '';
  els.comment.value = t.comment || '';
  if (!(keepSubject && els.subject.value.trim())) els.subject.value = withDate(withToday(t.subject || ''));
  els.link.value = '';
  els.link2.value = ''; // ссылки на Band — у каждой задачи свои, из прошлой не тащим
  refreshAutoDate();
  const langs = t.languages || [];
  checkedLangInputs(false).forEach((cb) => { cb.checked = langs.includes(cb.value); });
  hideSuggest();
  updatePreview();
  applying = false;
}

// ---------- Подсказка «как обычно» по подрядчику ----------
let suggestion = null;
function hideSuggest() {
  suggestion = null;
  els.suggest.hidden = true;
}

function suggestFor(contractor) {
  hideSuggest();
  if (!contractor || applying) return;
  // В форме уже что-то выбрано — не мешаем
  if (els.product.value || checkedLangInputs().length) return;
  const tpl = templates.find((t) => t.contractor === contractor);
  const memo = byContractor[contractor];
  if (tpl) {
    suggestion = { fields: tpl, tplIndex: templates.indexOf(tpl) };
    els.suggestText.textContent = `С ${contractor} есть шаблон «${tpl.name}»: ${describe(tpl)}`;
  } else if (memo) {
    suggestion = { fields: memo };
    els.suggestText.textContent = `Как в прошлый раз с ${contractor}: ${describe(memo)}`;
  } else {
    return;
  }
  els.suggest.hidden = false;
}

els.suggestApply.addEventListener('click', () => {
  if (!suggestion) return;
  const s = suggestion;
  if (s.tplIndex != null) els.template.value = String(s.tplIndex);
  els.delTpl.hidden = s.tplIndex == null;
  applyFields(s.fields, { keepSubject: true });
  renderChips();
  setMsg('Заполнено. Проверьте тему и, если есть, вставьте ссылку на Band.');
});
els.suggestClose.addEventListener('click', hideSuggest);
els.contractor.addEventListener('change', () => suggestFor(els.contractor.value));

/** Запомнить добавленную задачу: для «Как в прошлый раз» и подсказки по подрядчику. */
async function rememberTask(task) {
  const memo = cleanTemplate({ ...task, name: 'Как в прошлый раз', comment: '' });
  lastTask = memo;
  byContractor = { ...byContractor, [task.contractor]: memo };
  await chrome.storage.local.set({ lastTask, byContractor });
  renderChips();
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
  renderChips();
  const t = templates[els.template.value];
  if (!t) return;
  applyFields(t);
  setMsg(`Заполнено по шаблону «${t.name}». Проверьте тему и, если есть, вставьте ссылку на Band.`);
});

els.saveTpl.addEventListener('click', async () => {
  if (!els.contractor.value && !els.subject.value.trim()) return setMsg('Заполните хотя бы подрядчика или тему — их и запомнит шаблон', 'err');
  const current = templates[els.template.value];
  const name = (prompt('Название шаблона:', current ? current.name : els.subject.value.trim()) || '').trim();
  if (!name) return;
  const existing = templates.findIndex((t) => t.name === name);
  if (existing !== -1 && !confirm(`Шаблон «${name}» уже есть. Заменить?`)) return;
  const tpl = templateFromForm(name);
  if (existing !== -1) templates[existing] = tpl; else templates.push(tpl);
  await storeTemplates(name);
  setMsg('');
  toast(`✓ Шаблон «${name}» сохранён`);
});

/** Шаблон из того, что сейчас в форме. */
function templateFromForm(name) {
  const ticketNum = els.ticketNum.value.trim();
  return cleanTemplate({
    name, contractor: els.contractor.value, ticket: ticketNum ? `LOCAL-${ticketNum}` : '', subject: els.subject.value,
    product: els.product.value, customer: els.customer.value, deadline: els.deadline.value, comment: els.comment.value,
    languages: checkedLangInputs().map((cb) => cb.value),
  });
}

// Поправить выбранный шаблон: без вопросов, имя то же
els.updTpl.addEventListener('click', async () => {
  const i = Number(els.template.value), t = templates[i];
  if (!t) return;
  if (!els.contractor.value && !els.subject.value.trim()) return setMsg('Заполните хотя бы подрядчика или тему', 'err');
  templates[i] = templateFromForm(t.name);
  await storeTemplates(t.name);
  setMsg('');
  toast(`✓ Шаблон «${t.name}» обновлён`);
});

els.renTpl.addEventListener('click', async () => {
  const t = templates[els.template.value];
  if (!t) return;
  const name = (prompt('Новое название шаблона:', t.name) || '').trim();
  if (!name || name === t.name) return;
  if (templates.some((x) => x.name === name)) return setMsg(`Шаблон «${name}» уже есть — выберите другое название`, 'err');
  t.name = name;
  await storeTemplates(name);
  setMsg('');
  toast(`✓ Шаблон переименован в «${name}»`);
});

els.delTpl.addEventListener('click', async () => {
  const t = templates[els.template.value];
  if (!t || !confirm(`Удалить шаблон «${t.name}»?`)) return;
  templates.splice(Number(els.template.value), 1);
  await storeTemplates();
  setMsg('');
  toast(`Шаблон «${t.name}» удалён`);
});

els.exportTpl.addEventListener('click', () => {
  const letters = lettersForExport();
  if (!templates.length && !Object.keys(letters).length) return ioMsg('Шаблонов пока нет: заполните форму и нажмите «Сохранить как шаблон»', 'err');
  const blob = new Blob([JSON.stringify({ type: TPL_FILE_TYPE, version: 1, templates, letters }, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'task-templates.json';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  const nl = Object.keys(letters).length;
  ioMsg(`Скачано шаблонов: ${templates.length}${nl ? `, писем подрядчикам: ${nl}` : ''}. Файл можно отправить коллеге — пусть нажмёт «Загрузить из файла».`, 'ok');
});

/** Скачать / загрузить шаблоны — в настройках (⚙️), сообщение — там же. */
function ioMsg(text, kind = 'info') {
  els.tplIoMsg.textContent = text;
  els.tplIoMsg.className = `status ${kind}`;
}

els.importTpl.addEventListener('click', () => els.importFile.click());
els.importFile.addEventListener('change', async () => {
  const file = els.importFile.files[0];
  els.importFile.value = '';
  if (!file) return;
  try {
    const data = JSON.parse(await file.text());
    const list = (Array.isArray(data) ? data : data.templates || []).map(cleanTemplate).filter((t) => t.name);
    const nl = Array.isArray(data) ? 0 : await importLetters(data.letters);
    if (!list.length && !nl) throw new Error('в файле нет шаблонов');
    let replaced = 0;
    list.forEach((t) => {
      const i = templates.findIndex((x) => x.name === t.name);
      if (i !== -1) { templates[i] = t; replaced++; } else templates.push(t);
    });
    await storeTemplates();
    ioMsg(`Загружено шаблонов: ${list.length}` + (replaced ? ` (заменено с тем же названием: ${replaced})` : '') + (nl ? `, писем подрядчикам: ${nl}` : ''), 'ok');
  } catch (e) {
    ioMsg(`Не получилось загрузить: ${e.message}`, 'err');
  }
});

// ---------- Номер и превью темы ----------
let idSeq = 0;
// Следующие номера для всех подрядчиков — одним запросом, чтобы при выборе подрядчика номер был сразу
let nextIds = {};
async function loadNextIds() {
  try {
    const r = await api({ action: 'previewTaskIds' });
    if (!r.ok) return; // старый скрипт не умеет — номер узнаем при выборе подрядчика
    nextIds = r.ids || {};
    if (els.contractor.value && !nextId && nextIds[els.contractor.value]) { nextId = nextIds[els.contractor.value]; updatePreview(); }
  } catch { /* не страшно */ }
}

async function refreshNextId() {
  const contractor = els.contractor.value;
  const seq = ++idSeq;
  nextId = nextIds[contractor] || '';
  updatePreview();
  if (!contractor || nextId) return;
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
  const subject = subjectText();
  return subject ? `${prefix} ${subject}` : prefix;
}

/** «Перевод строчек … от» → «… от 01.10»: дата из поля «Дата» (по умолчанию сегодня). */
function withDate(s) {
  const [, m, d] = els.date.value.split('-');
  return d && m ? s.replace(/(^|\s)(от)\s*$/i, `$1$2 ${d}.${m}`) : s;
}
const subjectText = () => withDate(els.subject.value.trim());

// Подсказка под темой — с датой из поля «Дата», чтобы было видно, что именно подставится
function updateOtHint() {
  const [, m, d] = els.date.value.split('-');
  const hint = document.querySelector('#tOtHint');
  if (hint && d && m) hint.innerHTML = `Напишите в конце «от» — дата подставится сама: «Новые строки от» → <b>«Новые строки от ${d}.${m}»</b>`;
}
els.date.addEventListener('change', updateOtHint);
els.date.addEventListener('input', updateOtHint);

// Ушли из поля — «от» в конце сразу превращается в «от 01.10», чтобы было видно, что уйдёт в тему
els.subject.addEventListener('blur', () => {
  const v = subjectText();
  if (v !== els.subject.value.trim()) { els.subject.value = v; updatePreview(); }
});
// Поменяли дату — «от 22.09» в теме меняется на неё же
els.date.addEventListener('change', () => { els.subject.value = withToday(els.subject.value); updatePreview(); });

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
// До добавления номер — догадка: если коллега добавит задачу в ту же секунду, номер сдвинется
els.copyPreview.addEventListener('click', async () => {
  await copyText(els.preview.textContent, 'Тема скопирована');
  toast('Номер пока предварительный — точная тема скопируется сама после «Добавить задачу»', 'warn');
  els.copyPreview.textContent = 'Скопировано ✓';
  setTimeout(() => { els.copyPreview.textContent = 'Скопировать'; }, 1500);
});

makeLinkList(els.link2);

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
    contractor: els.contractor.value, subject: subjectText(),
    link: els.link.value.trim(), link2: els.link2.value.split(/[\s,;]+/).filter(Boolean).join('\n'),
    date: els.date.value, product: els.product.value, customer: els.customer.value.trim(),
    deadline: els.deadline.value, exactDeadline: els.exactDeadline.value,
    status: els.status.value, deliveryStatus: els.deliveryStatus.value,
    estimateLink: els.estimateLink.value.trim(), total: els.total.value, sp: els.sp.value,
    manager: els.manager.value, comment: els.comment.value.trim(), complaints: els.complaints.value.trim(),
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
    nextIds = {};
    loadNextIds(); // номера сдвинулись
    rememberTask(task).catch(() => {});
    lastSubject = buildSubject(r.id);
    els.doneId.textContent = r.id;
    els.doneSubject.textContent = lastSubject;
    // Номер теперь точный — сразу кладём тему в буфер
    document.getElementById('toast').hidden = true; // предупреждение о предварительном номере больше не нужно
    els.doneCopied.hidden = true;
    navigator.clipboard.writeText(lastSubject).then(() => { els.doneCopied.hidden = false; }).catch(() => {});
    setLink(els.openRow, r.url);
    // Письмо — подрядчику: штатные языки (переводят свои) ему не нужны
    const shtat = new Set((lists.languages && lists.languages.shtat) || []);
    showLetter({
      contractor: task.contractor, id: r.id, subject: lastSubject,
      languages: task.languages.filter((l) => !shtat.has(l) && !/^ШТАТ /i.test(l)),
      codes: task.languages.map((l) => lists.langCodes[l]).filter(Boolean), exactDeadline: task.exactDeadline,
      product: task.product, manager: task.manager, link: task.link,
    }).catch(() => {});
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

// «✕ Сбросить» — снять шаблон и очистить форму
els.clearTpl.addEventListener('click', () => {
  resetForm();
  setMsg('');
});

// ---------- Срок сдачи на выходной ----------
// Выбрали субботу или воскресенье — подсказка и кнопки «← пятница» / «понедельник →»
const weekendHint = document.querySelector('#tWeekend');
const isoOf = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
function checkWeekend() {
  const [y, m, d] = els.exactDeadline.value.split('-').map(Number);
  const date = y ? new Date(y, m - 1, d) : null;
  const wd = date ? date.getDay() : -1;
  if (wd !== 0 && wd !== 6) { weekendHint.hidden = true; return; }
  const fri = new Date(date); fri.setDate(date.getDate() - (wd === 6 ? 1 : 2));
  const mon = new Date(date); mon.setDate(date.getDate() + (wd === 6 ? 2 : 1));
  const dm = (x) => `${String(x.getDate()).padStart(2, '0')}.${String(x.getMonth() + 1).padStart(2, '0')}`;
  weekendHint.innerHTML = `${wd === 6 ? 'Это суббота' : 'Это воскресенье'}: <button type="button" class="link-btn strong" data-to="${isoOf(fri)}">← пт ${dm(fri)}</button> <button type="button" class="link-btn strong" data-to="${isoOf(mon)}">пн ${dm(mon)} →</button>`;
  weekendHint.hidden = false;
}
els.exactDeadline.addEventListener('change', checkWeekend);
els.exactDeadline.addEventListener('input', checkWeekend);
weekendHint.addEventListener('click', (e) => {
  const b = e.target.closest('button[data-to]');
  if (!b) return;
  e.preventDefault();
  els.exactDeadline.value = b.dataset.to;
  els.exactDeadline.dispatchEvent(new Event('change', { bubbles: true }));
});
