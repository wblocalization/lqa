// Письмо подрядчику — на экране «Добавлено»: кому (почты подрядчика), копия и текст по шаблону.
// Скопировать по частям (панель открыта рядом с Outlook) или открыть новое письмо в Outlook уже заполненным.
// Почты и шаблоны — свои у каждого подрядчика, хранятся в расширении (chrome.storage, ключ «letters»),
// делятся с коллегами вместе с шаблонами задач (⚙️ → «Скачать файлом»).
import { toast, ask } from './core.js';
import { OWA_ORIGIN, openSent } from './outlook.js';
import { sendMail } from './owa-send.js';
import { richEditor, htmlToRichText } from './richtext.js';
import { playSound } from './sounds.js';
import { soundPrefs } from './remind.js';

const $ = (s) => document.querySelector(s);
const els = {
  box: $('#letter'), contractor: $('#lContractor'), edit: $('#lEdit'), view: $('#lView'), empty: $('#lEmpty'),
  subject: $('#lSubject'), copySubject: $('#lCopySubject'), resetSubject: $('#lResetSubject'), to: $('#lTo'), cc: $('#lCc'),
  body: $('#lBody'), texts: $('#lTexts'), eNameRow: $('#lENameRow'), eName: $('#lEName'), eAddr: $('#lEAddr'), copyBody: $('#lCopyBody'), compose: $('#lCompose'), links: $('#lLinks'),
  sig: $('#lSig'), sigEdit: $('#lSigEdit'), sigText: $('#lSigText'), sigFetch: $('#lSigFetch'), sigSave: $('#lSigSave'),
  send: $('#lSend'), sent: $('#lSent'), sentText: $('#lSentText'), sentOpen: $('#lSentOpen'), drop: $('#lDrop'), fileInput: $('#lFileInput'), fileList: $('#lFileList'), attachOpen: $('#lAttachOpen'),
  editor: $('#lEditor'), eTo: $('#lETo'), eCc: $('#lECc'), eBody: $('#lEBody'), save: $('#lSave'), cancel: $('#lCancel'), remove: $('#lRemove'),
};

export const DEFAULT_BODY = 'Здравствуйте!\n\nВозьмите, пожалуйста, в работу.\nЯзыки: {языки}\nДедлайн: {срок}\n\nБольшое спасибо!';
// Прежние тексты по умолчанию — заменяем на новый сами (свой текст человека не трогаем)
const OLD_DEFAULTS = [
  'Добрый день!\n\nПросим взять в работу: {тема}\nЯзыки: {языки}\nСрок сдачи: {срок}\n\nСпасибо!',
  'Добрый день!\n\nВозьмите, пожалуйста, в работу.\n\nСпасибо!',
];

/** Тема письма — как сейчас в поле (её можно поправить); пусто — тема задачи. */
const subjectNow = () => els.subject.value.replace(/\s*\n\s*/g, ' ').trim() || ctx.subject;

let letters = {};   // { [подрядчик]: { to, cc, body } }
let ctx = null;     // задача, которую только что добавили

// Тексты писем — общая библиотека для всех подрядчиков: «Новые строки для веба», «Файлы для приложения»…
// Шаблон задачи может быть привязан к тексту (letterText = id) — тогда письмо сразу с ним. Без привязки — текст подрядчика.
let texts = []; // [{ id, name, body }]

export async function loadLetters() {
  const r = await chrome.storage.local.get(['letters', 'letterTexts']);
  letters = r.letters && typeof r.letters === 'object' ? r.letters : {};
  texts = Array.isArray(r.letterTexts) ? r.letterTexts.filter((t) => t && t.id && t.name) : [];
  Object.values(letters).forEach((l) => { if (l && OLD_DEFAULTS.includes(l.body)) l.body = DEFAULT_BODY; });
  return letters;
}

/** Для файла «Скачать файлом» / «Загрузить из файла» — вместе с шаблонами задач. */
export const lettersForExport = () => letters;
export const textsForExport = () => texts;
/** Тексты писем — для выбора в форме задачи. */
export async function listTexts() { await loadLetters(); return texts.map((t) => ({ id: t.id, name: t.name })); }
const saveTexts = () => chrome.storage.local.set({ letterTexts: texts });
/** Загрузить тексты из файла: с тем же названием — заменить. */
export async function importTexts(list) {
  if (!Array.isArray(list)) return 0;
  await loadLetters();
  let n = 0;
  list.forEach((t) => {
    if (!t || typeof t.name !== 'string' || typeof t.body !== 'string' || !t.name.trim()) return;
    const id = typeof t.id === 'string' && t.id ? t.id : `t${Date.now().toString(36)}${n}`;
    const i = texts.findIndex((x) => x.id === id || x.name === t.name);
    const item = { id: i !== -1 ? texts[i].id : id, name: t.name.trim(), body: t.body };
    if (i !== -1) texts[i] = item; else texts.push(item);
    n++;
  });
  await saveTexts();
  return n;
}
export async function importLetters(data) {
  if (!data || typeof data !== 'object' || Array.isArray(data)) return 0;
  await loadLetters();
  let n = 0;
  Object.keys(data).forEach((k) => {
    const v = data[k];
    if (!k || !v || typeof v !== 'object') return;
    const str = (x) => (typeof x === 'string' ? x : '');
    letters[k] = { to: str(v.to), cc: str(v.cc), body: str(v.body) };
    n++;
  });
  await chrome.storage.local.set({ letters });
  return n;
}

/** «a@x.ru, b@y.ru\nc@z.ru» → ['a@x.ru', 'b@y.ru', 'c@z.ru'] — так Outlook вставляет без сюрпризов. */
export function parseEmails(text) {
  return [...new Set(String(text || '').split(/[\s,;]+/).map((s) => s.trim().replace(/^<|>$/g, '')).filter((s) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s)))];
}

/**
 * Ссылки на материалы — в текст письма: на место {материалы}, а если его в шаблоне нет —
 * отдельным абзацем перед последним (подписью «Большое спасибо!»).
 */
export function withMaterials(text, links) {
  const list = String(links || '').split(/\s+/).map((s) => s.trim()).filter((s) => /^https?:\/\//i.test(s));
  const block = list.length ? `Материалы:\n${list.join('\n')}` : '';
  if (/\{материалы\}/i.test(text)) return text.replace(/\{материалы\}/gi, block).replace(/\n{3,}/g, '\n\n').trim();
  if (!block) return text;
  const parts = text.split(/\n\n/);
  if (parts.length < 2) return `${text}\n\n${block}`;
  parts.splice(parts.length - 1, 0, block);
  return parts.join('\n\n');
}

/** Подставить в шаблон данные задачи: {тема} {номер} {языки} {коды} {срок} {продукт} {менеджер} {ссылка}. */
function varValues(c) {
  return {
    тема: c.subject, номер: c.id, языки: (c.languages || []).map((l) => l.toLowerCase()).join(', '), коды: (c.codes || []).join(', '),
    срок: c.due, // только «Срок сдачи» из формы, без года продукт: c.product, менеджер: c.manager, ссылка: c.link,
  };
}
export function fillTemplate(tpl, c) {
  const map = varValues(c);
  return String(tpl || '').replace(/\{([а-яё]+)\}/gi, (m, k) => (k.toLowerCase() in map ? map[k.toLowerCase()] || '' : m));
}

const ddmm = (iso) => { const [, m, d] = String(iso || '').split('-'); return d ? `${d}.${m}` : ''; }; // без года: «22.10»

/** Показать блок письма для только что добавленной задачи. c: { contractor, id, subject, languages, codes, exactDeadline, product, manager, link } */
export async function showLetter(c) {
  await loadLetters();
  await loadSignature();
  if (!signature) {
    // подписи ещё нет — тихо пробуем взять из Outlook
    fetchOutlookSignature().then(async (s) => {
      if (!s || signature) return;
      signature = s;
      await chrome.storage.local.set({ signature });
      if (ctx && !els.view.hidden) renderSignature();
    }).catch(() => {});
  }
  ctx = { ...c, due: ddmm(c.exactDeadline), textId: c.textId || '' };
  els.links.value = ''; // ссылки и вложения — у каждой задачи свои
  files = [];
  renderFiles();
  els.send.disabled = false;
  els.send.textContent = '📨 Отправить';
  els.sent.hidden = true;
  els.contractor.textContent = c.contractor || 'подрядчику';
  els.box.hidden = !c.contractor;
  closeEditor();
  render();
}

let bodyEdited = false; // текст поправили руками — шаблон и ссылки больше не перезаписывают его

// Текст письма и шаблона — как в почте: жирный виден жирным; в шаблоне вместо {языки}, {срок}… сразу
// видно, что подставится из этой задачи (цветным).
const bodyEd = richEditor(els.body, { onInput: () => { bodyEdited = true; }, onSelect: () => syncBold() });
let editing = null; // { contractor, host, onDone, ctx } — чей шаблон открыт и где (под задачей или во вкладке «Настройки»)
const tplEd = richEditor(els.eBody, { valueOf: (name) => (editing && editing.ctx ? varValues(editing.ctx)[name] || '' : ''), onSelect: () => syncBold() });
// Подпись — тоже с жирным; скопированная из Outlook вставляется с жирным
const sigEd = richEditor(els.sigText, { pasteBold: true, onSelect: () => syncBold() });
const cardEd = richEditor($('#mlSigText'), { pasteBold: true, onSelect: () => syncBold() });
const boldBtns = [[$('#lBold'), bodyEd], [$('#lEBold'), tplEd], [$('#lSigBold'), sigEd], [$('#mlSigBold'), cardEd]];
function syncBold() {
  boldBtns.forEach(([btn, ed]) => {
    const on = ed.isBold();
    btn.classList.toggle('on', on);
    btn.lastChild.textContent = on ? ' Убрать жирный' : ' Выделить жирным';
  });
}

// Смайлики — кнопка 😊 рядом с «Выделить жирным»: вставляется туда, где курсор
const EMOJI = ['😊', '🙂', '😉', '😄', '🙏', '👍', '👌', '🤝', '👋', '💪', '🎉', '✨', '🔥', '❤️', '💜', '🌸',
  '✅', '☑️', '❗', '⚠️', '⏰', '📅', '📌', '📎', '📄', '📝', '📨', '🔗', '🌍', '🚀', '⭐', '💡'];
const emojiPop = document.createElement('div');
emojiPop.className = 'emoji-pop';
emojiPop.hidden = true;
emojiPop.innerHTML = EMOJI.map((e) => `<button type="button" data-e="${e}">${e}</button>`).join('');
document.body.appendChild(emojiPop);
let emojiFor = null;
document.querySelectorAll('.emoji-btn').forEach((b) => {
  b.addEventListener('mousedown', (e) => e.preventDefault()); // не терять место курсора
  b.addEventListener('click', (e) => {
    e.stopPropagation();
    const ed = b.dataset.ed === 'tpl' ? tplEd : bodyEd;
    if (!emojiPop.hidden && emojiFor === ed) { emojiPop.hidden = true; return; }
    emojiFor = ed;
    const r = b.getBoundingClientRect();
    emojiPop.style.top = `${r.bottom + window.scrollY + 6}px`;
    emojiPop.style.left = `${Math.max(8, Math.min(r.left + window.scrollX, document.documentElement.clientWidth - 268))}px`;
    emojiPop.hidden = false;
  });
});
emojiPop.addEventListener('mousedown', (e) => e.preventDefault());
emojiPop.addEventListener('click', (e) => {
  const b = e.target.closest('button[data-e]');
  if (!b || !emojiFor) return;
  emojiFor.insertText(b.dataset.e);
  emojiPop.hidden = true;
});
document.addEventListener('click', (e) => { if (!emojiPop.hidden && !emojiPop.contains(e.target)) emojiPop.hidden = true; });
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') emojiPop.hidden = true; });

boldBtns.forEach(([btn, ed]) => {
  btn.addEventListener('mousedown', (e) => e.preventDefault()); // не терять выделение
  btn.addEventListener('click', () => {
    if (!ed.bold()) return toast('Сначала выделите слово мышкой, потом нажмите «Выделить жирным»', 'warn');
    syncBold();
  });
});

function render() {
  const l = letters[ctx.contractor];
  els.empty.hidden = Boolean(l);
  els.view.hidden = !l;
  if (!l) return;
  // Тему, поправленную руками, не сбрасываем, пока это то же письмо (перерисовка после смены адресов и т. п.)
  const key = `${ctx.id || ''}|${ctx.subject}`;
  if (els.subject.dataset.key !== key) { els.subject.value = ctx.subject; els.subject.dataset.key = key; }
  els.resetSubject.hidden = subjectNow() === ctx.subject;
  requestAnimationFrame(fitSubject); // после показа панели — тогда известна высота текста
  els.to.value = parseEmails(l.to).join(', ');
  els.cc.value = parseEmails(l.cc).join(', ');
  bodyEdited = false;
  if (ctx.textId && !texts.some((t) => t.id === ctx.textId)) ctx.textId = ''; // текст удалили — берём текст подрядчика
  renderTextChips();
  fillBody();
  if (!ctx.due && /\{срок\}/i.test(currentTemplate())) toast('В задаче нет «Срока сдачи» — впишите дату в письмо сами', 'warn');
  renderSignature();
}
/** Шаблон текста письма сейчас: выбранный из библиотеки или текст подрядчика. */
function currentTemplate() {
  const t = texts.find((x) => x.id === ctx.textId);
  return t ? t.body : ((letters[ctx.contractor] || {}).body || DEFAULT_BODY);
}
function fillBody() {
  if (!bodyEdited) bodyEd.set(withMaterials(fillTemplate(currentTemplate(), ctx), els.links.value));
}
/** Переключатель текста над письмом: «Обычный» (текст подрядчика) и тексты из библиотеки. */
function renderTextChips() {
  els.texts.hidden = !texts.length;
  els.texts.innerHTML = texts.length ? '<span class="lt-lbl">Текст:</span>' + [{ id: '', name: 'Обычный' }, ...texts].map((t) =>
    `<button type="button" class="lt-chip${(ctx.textId || '') === t.id ? ' on' : ''}" data-text="${t.id}">${escHtml(t.name)}</button>`).join('') : '';
}
const escHtml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
els.texts.addEventListener('click', async (e) => {
  const b = e.target.closest('button[data-text]');
  if (!b || b.dataset.text === (ctx.textId || '')) return;
  if (bodyEdited && !await ask({ title: 'Заменить текст письма?', text: 'Вы уже правили письмо — правки пропадут.', ok: 'Заменить' })) return;
  ctx.textId = b.dataset.text;
  bodyEdited = false;
  renderTextChips();
  fillBody();
});

async function copy(text, done) {
  try { await navigator.clipboard.writeText(text); toast(done); } catch { toast('Нет доступа к буферу — выделите и скопируйте вручную', 'err'); }
}

els.copySubject.addEventListener('click', () => copy(subjectNow(), '✓ Тема скопирована'));
// Тему можно поправить прямо в письме (в таблице она остаётся прежней); ↺ — вернуть тему задачи
// Поле темы растёт под длинную тему, чтобы её было видно целиком
function fitSubject() { els.subject.style.height = 'auto'; els.subject.style.height = `${els.subject.scrollHeight + 2}px`; }
els.subject.addEventListener('input', () => { els.resetSubject.hidden = subjectNow() === ctx.subject; fitSubject(); });
els.subject.addEventListener('keydown', (e) => { if (e.key === 'Enter') e.preventDefault(); }); // тема — одна строка
els.resetSubject.addEventListener('click', () => { els.subject.value = ctx.subject; els.resetSubject.hidden = true; fitSubject(); });
const plain = (t) => String(t).replace(/\*\*([\s\S]+?)\*\*/g, '$1'); // без звёздочек «жирного» — для Outlook и буфера
els.copyBody.addEventListener('click', () => copy(plain(bodyEd.get()), '✓ Текст письма скопирован'));

// ---------- Подпись ----------
// Одна на все письма. Сначала пробуем взять из настроек Outlook; не вышло — человек вставляет её один раз сам.
let signature = null; // { html, text }
async function loadSignature() {
  const r = await chrome.storage.local.get('signature');
  signature = r.signature && (r.signature.html || r.signature.text || r.signature.image) ? r.signature : null;
}

// Цвет текста подписи — как в Outlook бывает серая подпись
const SIG_COLORS = [['', 'Чёрный'], ['#595959', 'Тёмно-серый'], ['#7F7F7F', 'Серый']];
const sigColor = () => (signature && signature.color) || '';

/** Подпись для письма: текст (или HTML из Outlook) + цвет + картинка ссылкой cid: на встроенное вложение. */
function signatureForMail() {
  if (!signature) return { html: '', inline: [] };
  const body = signature.html || (signature.text ? textToSigHtml(signature.text) : '');
  const color = sigColor();
  let html = body ? (color ? `<div style="color:${color}">${body}</div>` : body) : '';
  const inline = [];
  const img = signature.image;
  if (img && img.dataUrl) {
    const cid = `sig-${Date.now().toString(36)}@zadachi`;
    html += `<div><img src="cid:${cid}" width="${img.width || 200}" alt="" style="width:${img.width || 200}px;height:auto;border:0"></div>`;
    inline.push({ name: img.name || 'signature.png', type: img.type || 'image/png', b64: img.dataUrl.split(',')[1], contentId: cid });
  }
  return { html, inline };
}

/** Картинка из файла/буфера → { dataUrl, type, name, width }. Большие уменьшаем до 400 px в ширину (GIF до 300 КБ не трогаем — анимация). */
async function prepareSigImage(file) {
  if (!/^image\//.test(file.type)) throw new Error('Это не картинка');
  const read = (f) => new Promise((ok, no) => { const r = new FileReader(); r.onload = () => ok(r.result); r.onerror = no; r.readAsDataURL(f); });
  const src = await read(file);
  const im = await new Promise((ok, no) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => no(new Error('Картинка не открылась')); i.src = src; });
  const MAX = 400;
  if ((file.type === 'image/gif' && file.size <= 300 * 1024) || (im.naturalWidth <= MAX && file.size <= 200 * 1024)) {
    return { dataUrl: src, type: file.type, name: file.name || 'signature.png', width: Math.min(im.naturalWidth, MAX) };
  }
  const k = Math.min(1, MAX / im.naturalWidth);
  const c = document.createElement('canvas');
  c.width = Math.round(im.naturalWidth * k); c.height = Math.round(im.naturalHeight * k);
  c.getContext('2d').drawImage(im, 0, 0, c.width, c.height);
  return { dataUrl: c.toDataURL('image/png'), type: 'image/png', name: 'signature.png', width: c.width };
}

async function setSigImage(image) {
  signature = { ...(signature || { text: '' }), image: image || undefined };
  if (!image) delete signature.image;
  if (!signature.html && !signature.text && !signature.image) signature = null;
  await chrome.storage.local.set({ signature });
  renderSigExtras();
  if (ctx && !els.view.hidden && els.sigEdit.hidden) renderSignature(); // открыт редактор — не закрываем его
  toast(image ? '✓ Картинка в подписи — будет в каждом письме' : 'Картинка из подписи убрана');
}
async function setSigColor(color) {
  signature = { ...(signature || { text: '' }), color };
  await chrome.storage.local.set({ signature });
  renderSigExtras();
  if (ctx && !els.view.hidden && els.sigEdit.hidden) renderSignature(); // открыт редактор — не закрываем его
}

/** Под полем подписи (в письме и в «Настройках»): картинка и цвет. */
function renderSigExtras() {
  ['#lSigImg', '#mlSigImg'].forEach((sel) => {
    const box = document.querySelector(sel);
    if (!box) return;
    const img = signature && signature.image;
    box.innerHTML = '';
    if (img) {
      const pic = document.createElement('img'); pic.src = img.dataUrl; pic.alt = 'Картинка в подписи';
      const del = document.createElement('button'); del.type = 'button'; del.className = 'link-btn'; del.textContent = '✕ убрать картинку';
      del.addEventListener('click', () => setSigImage(null));
      box.append(pic, del);
    } else {
      const add = document.createElement('button'); add.type = 'button'; add.className = 'btn ghost small'; add.textContent = '🖼 Добавить картинку';
      const input = document.createElement('input'); input.type = 'file'; input.accept = 'image/*'; input.hidden = true;
      add.addEventListener('click', () => input.click());
      input.addEventListener('change', async () => { const f = input.files[0]; if (!f) return; try { await setSigImage(await prepareSigImage(f)); } catch (e) { toast(e.message, 'err'); } });
      const hint = document.createElement('span'); hint.className = 'hint-inline'; hint.textContent = 'или вставьте картинку в подпись (Ctrl+V)';
      box.append(add, input, hint);
    }
    const col = document.createElement('label'); col.className = 'hint-inline';
    col.append('Цвет текста: ');
    const sel2 = document.createElement('select');
    SIG_COLORS.forEach(([v, t]) => { const o = document.createElement('option'); o.value = v; o.textContent = t; if (v === sigColor()) o.selected = true; sel2.append(o); });
    sel2.addEventListener('change', () => setSigColor(sel2.value));
    col.append(sel2);
    box.append(col);
  });
}
// Ctrl+V картинки прямо в поле подписи — это картинка подписи (раньше обработчика текста)
['#lSigText', '#mlSigText'].forEach((sel) => {
  const el = document.querySelector(sel);
  if (!el) return;
  el.addEventListener('paste', async (e) => {
    const item = [...((e.clipboardData && e.clipboardData.items) || [])].find((i) => i.kind === 'file' && /^image\//.test(i.type));
    if (!item) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    try { await setSigImage(await prepareSigImage(item.getAsFile())); } catch (err) { toast(err.message, 'err'); }
  }, true);
});
// Подпись храним как текст с **жирным** (или как HTML из Outlook, пока её не правили)
const sigToText = (s) => (s ? s.text || htmlToRichText(s.html) : '');
const textToSigHtml = (text) => String(text).split('\n')
  .map((l) => l.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\*\*(.+?)\*\*/g, '<b>$1</b>') || '<br>').join('<br>');

function renderSignature() {
  els.sigEdit.hidden = true;
  els.sig.hidden = false;
  els.sig.innerHTML = '';
  if (signature) {
    const pre = document.createElement('div');
    pre.className = 'c-sig-text';
    pre.innerHTML = textToSigHtml(sigToText(signature));
    if (sigColor()) pre.style.color = sigColor();
    if (signature.image) {
      const im = document.createElement('img'); im.className = 'c-sig-img'; im.src = signature.image.dataUrl; im.alt = '';
      im.style.width = `${Math.min(signature.image.width || 200, 240)}px`;
      pre.append(im);
    }
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'link-btn'; b.textContent = 'изменить подпись';
    b.addEventListener('click', openSigEdit);
    els.sig.append(pre, b);
  } else {
    els.sig.innerHTML = '<span class="muted">Подписи пока нет — </span>';
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'link-btn strong'; b.textContent = 'добавить подпись';
    b.addEventListener('click', openSigEdit);
    els.sig.append(b);
  }
}
function openSigEdit() {
  sigEd.set(sigToText(signature));
  renderSigExtras();
  els.sig.hidden = true;
  els.sigEdit.hidden = false;
  sigEd.focus();
}
async function saveSignature(text) {
  text = text.trim();
  // текст поменяли — своя подпись, без оформления Outlook
  const keep = signature ? { image: signature.image, color: signature.color } : {};
  signature = text ? (signature && signature.html && sigToText({ html: signature.html }) === text ? signature : { text, ...keep }) : (keep.image ? { text: '', ...keep } : null);
  if (signature && !signature.image) delete signature.image;
  await chrome.storage.local.set({ signature });
  if (ctx) renderSignature();
  toast(text ? '✓ Подпись сохранена — будет в каждом письме' : 'Подпись убрана');
}
async function fetchInto(btn, ed) {
  btn.disabled = true;
  try {
    const s = await fetchOutlookSignature();
    if (!s) return toast('Не нашла подпись в Outlook — вставьте её сюда текстом', 'warn');
    signature = { ...s, color: s.color || sigColor() || undefined };
    if (s.image) await chrome.storage.local.set({ signature }); // картинку — сразу, текст — по «Сохранить»
    ed.set(sigToText(s));
    renderSigExtras();
    toast(s.image ? '✓ Подпись и картинка взяты из Outlook — нажмите «Сохранить подпись»' : '✓ Подпись взята из Outlook — нажмите «Сохранить подпись»');
  } finally {
    btn.disabled = false;
  }
}
els.sigSave.addEventListener('click', () => saveSignature(sigEd.get()));
els.sigFetch.addEventListener('click', () => fetchInto(els.sigFetch, sigEd));

// Подпись во вкладке «Настройки» — то же самое, что «изменить подпись» в письме
const card = { save: $('#mlSigSave'), fetch: $('#mlSigFetch') };
export async function renderSignatureCard() {
  await loadSignature();
  cardEd.set(sigToText(signature));
  renderSigExtras();
}
card.save.addEventListener('click', () => saveSignature(cardEd.get()));
card.fetch.addEventListener('click', () => fetchInto(card.fetch, cardEd));

/** Подпись из Outlook (нужна открытая вкладка с почтой). */
async function fetchOutlookSignature() {
  const tabs = await chrome.tabs.query({ url: `${OWA_ORIGIN}/owa/*` });
  if (!tabs.length) return null;
  try {
    const [{ result }] = await chrome.scripting.executeScript({ target: { tabId: tabs[0].id }, world: 'MAIN', func: owaSignatureInPage });
    return result && (result.html || result.text) ? result : null;
  } catch {
    return null;
  }
}

/** Выполняется на странице Outlook: ищем подпись в настройках пользователя (там же, откуда её берёт сам Outlook). */
async function owaSignatureInPage() {
  // Картинку из подписи забираем отдельно (data: или скачиваем по ссылке), а из HTML её убираем —
  // в письмо она уйдёт встроенным вложением, иначе у получателей она часто не показывается
  async function withImage(sig) {
    if (!sig || !sig.html || !/<img/i.test(sig.html)) return sig;
    const d = document.createElement('div');
    d.innerHTML = sig.html;
    const img = d.querySelector('img');
    let image = null;
    try {
      const src = img.getAttribute('src') || '';
      const blob = await (await fetch(src, { credentials: 'include' })).blob();
      const dataUrl = await new Promise((ok) => { const r = new FileReader(); r.onload = () => ok(r.result); r.readAsDataURL(blob); });
      if (/^data:image\//.test(dataUrl)) image = { dataUrl, type: blob.type, name: 'signature.' + ((blob.type.split('/')[1] || 'png').replace('jpeg', 'jpg')), width: Math.min(Number(img.getAttribute('width')) || img.naturalWidth || 200, 400) };
    } catch { /* не скачалась — оставим без картинки */ }
    d.querySelectorAll('img').forEach((x) => x.remove());
    return { html: d.innerHTML, image };
  }
  const find = (o, depth = 0) => {
    if (!o || typeof o !== 'object' || depth > 8) return null;
    for (const k of Object.keys(o)) {
      if (/^SignatureHtml$/i.test(k) && typeof o[k] === 'string' && o[k].trim()) return { html: o[k] };
      if (/^SignatureText$/i.test(k) && typeof o[k] === 'string' && o[k].trim()) return { text: o[k] };
    }
    for (const k of Object.keys(o)) { const r = find(o[k], depth + 1); if (r) return r; }
    return null;
  };
  const canary = decodeURIComponent((document.cookie.match(/(?:^|;\s*)X-OWA-CANARY=([^;]*)/i) || [])[1] || '');
  const headers = { 'Content-Type': 'application/json; charset=utf-8', 'X-OWA-CANARY': canary, 'X-Requested-With': 'XMLHttpRequest' };
  const tries = [
    () => fetch('/owa/sessiondata.ashx?appcache=true', { method: 'POST', credentials: 'include', headers }),
    () => fetch('/owa/service.svc?action=GetOwaUserConfiguration&AC=1', { method: 'POST', credentials: 'include', headers: { ...headers, Action: 'GetOwaUserConfiguration', 'X-OWA-ActionName': 'GetOwaUserConfiguration' }, body: '{}' }),
  ];
  for (const t of tries) {
    try {
      const r = await t();
      if (!r.ok) continue;
      const found = find(await r.json());
      if (found) return await withImage(found);
    } catch { /* следующий способ */ }
  }
  // открыт черновик — подпись прямо в нём
  const el = document.querySelector('#Signature');
  return el && el.textContent.trim() ? withImage({ html: el.innerHTML }) : null;
}

// ---------- Вложения ----------
// Файлы лежат только в панели, пока не вложены в письмо; кладём их через кнопку Outlook «Вложить».
const MAX_TOTAL = 30 * 1024 * 1024; // почта больше всё равно не пропустит
let files = [];

const sizeText = (n) => (n >= 1048576 ? `${(n / 1048576).toFixed(1)} МБ` : `${Math.max(1, Math.round(n / 1024))} КБ`);
function renderFiles() {
  els.fileList.innerHTML = files.map((f, i) => `<li><span class="f-name">📄 ${f.name.replace(/</g, '&lt;')}</span><span class="f-size">${sizeText(f.size)}</span><button type="button" class="icon-btn small" data-rm="${i}" title="Убрать">✕</button></li>`).join('');
  els.fileList.hidden = !files.length;
  els.attachOpen.hidden = !files.length;
}
function addFiles(list) {
  for (const f of list) {
    if (files.some((x) => x.name === f.name && x.size === f.size)) continue;
    files.push(f);
  }
  const total = files.reduce((n, f) => n + f.size, 0);
  if (total > MAX_TOTAL) toast(`Вложений на ${sizeText(total)} — почта может не пропустить. Большие архивы лучше ссылкой`, 'warn');
  renderFiles();
}
els.drop.addEventListener('click', () => els.fileInput.click());
els.fileInput.addEventListener('change', () => { addFiles(els.fileInput.files); els.fileInput.value = ''; });
els.drop.addEventListener('dragover', (e) => { e.preventDefault(); els.drop.classList.add('over'); });
els.drop.addEventListener('dragleave', () => els.drop.classList.remove('over'));
els.drop.addEventListener('drop', (e) => { e.preventDefault(); els.drop.classList.remove('over'); addFiles(e.dataTransfer.files); });
els.fileList.addEventListener('click', (e) => {
  const b = e.target.closest('button[data-rm]');
  if (!b) return;
  files.splice(Number(b.dataset.rm), 1);
  renderFiles();
});

function toBase64(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

function waitComplete(tabId, timeoutMs = 30000) {
  return new Promise((resolve) => {
    const timer = setTimeout(done, timeoutMs);
    function done() { clearTimeout(timer); chrome.tabs.onUpdated.removeListener(on); resolve(); }
    function on(id, info) { if (id === tabId && info.status === 'complete') done(); }
    chrome.tabs.onUpdated.addListener(on);
    chrome.tabs.get(tabId).then((t) => { if (t.status === 'complete') done(); }).catch(done);
  });
}

/** Вложить файлы из панели в письмо во вкладке Outlook. */
async function attachTo(tabId) {
  if (!files.length) return;
  toast(`📎 Вкладываю файлы: ${files.length}…`);
  const payload = await Promise.all(files.map(async (f) => ({ name: f.name, type: f.type, b64: toBase64(new Uint8Array(await f.arrayBuffer())) })));
  let r = null;
  try {
    [{ result: r }] = await chrome.scripting.executeScript({ target: { tabId }, world: 'MAIN', func: owaAttach, args: [payload] });
  } catch { /* нет доступа к вкладке */ }
  if (r && r.ok) {
    toast(r.unconfirmed ? '📎 Файлы отправлены в письмо — проверьте, что они появились во вложениях' : `📎 Вложено файлов: ${files.length}`, r.unconfirmed ? 'warn' : 'ok');
    files = [];
    renderFiles();
  } else {
    toast('Не нашла в письме кнопку «Вложить» — перетащите файлы в письмо вручную', 'err');
  }
}

els.attachOpen.addEventListener('click', async () => {
  // письмо, открытое в Outlook последним
  const tabs = await chrome.tabs.query({ url: `${OWA_ORIGIN}/owa/*` });
  const tab = tabs.sort((a, b) => (b.lastAccessed || 0) - (a.lastAccessed || 0))[0];
  if (!tab) return toast('Outlook не открыт — сначала откройте письмо', 'err');
  await chrome.tabs.update(tab.id, { active: true });
  await chrome.windows.update(tab.windowId, { focused: true });
  attachTo(tab.id);
});

/**
 * Явные промахи перед отправкой: «во вложении», а файлов нет; пустой дедлайн; не те адреса; незаполненное.
 * Возвращает список предупреждений (пустой — всё в порядке).
 */
function checkLetter(text, to) {
  const out = [];
  const t = plain(text);
  if (!files.length && /вложени|прикрепл|attach/i.test(t)) out.push('В тексте про вложение, а файлов нет — добавьте их в «📎 Вложения»');
  if (/(дедлайн|срок[а-я]*)\s*:\s*(\n|$)/i.test(t)) out.push('Дедлайн пустой — в задаче нет «Срока сдачи»');
  if (/языки\s*:\s*(\n|$)/i.test(t)) out.push('Языки пустые — у подрядчика в задаче только штатные или ни одного');
  if (/\{[а-яё]+\}/i.test(t)) out.push('В тексте осталась {переменная} в скобках');
  const saved = parseEmails((letters[ctx.contractor] || {}).to);
  if (saved.length && !saved.some((e) => to.includes(e))) out.push(`В «Кому» нет адресов ${ctx.contractor} из шаблона (${saved.join(', ')})`);
  if (ctx.id && !subjectNow().includes(ctx.id)) out.push('В теме нет номера задачи');
  return out;
}

// «📨 Отправить» — прямо отсюда, с рабочей почты; перед отправкой — ещё раз показать, что уйдёт, и спросить
els.send.addEventListener('click', async () => {
  const to = parseEmails(els.to.value);
  const cc = parseEmails(els.cc.value);
  if (!to.length) return toast('Впишите, кому отправить', 'err');
  const text = bodyEd.get().trim();
  if (!text) return toast('Письмо пустое', 'err');
  const sigText = plain(sigToText(signature));
  const warnings = checkLetter(text, to);
  const ok = await ask({ title: warnings.length ? 'Проверьте письмо' : 'Отправить письмо?', warnings,
    ok: warnings.length ? 'Всё равно отправить' : 'Отправить', icon: warnings.length ? '!' : '📨', text: `Кому: ${to.join(', ')}` + (cc.length ? `\nКопия: ${cc.join(', ')}` : '') +
    `\nТема: ${subjectNow()}` + (files.length ? `\nВложения: ${files.map((f) => f.name).join(', ')}` : '') +
    `\n\n${plain(text.length > 400 ? `${text.slice(0, 400)}…` : text)}` + (sigText ? `\n\n${sigText}` : (signature && signature.image ? '' : '\n\n(без подписи)')) + (signature && signature.image ? '\n🖼 + картинка в подписи' : '') });
  if (!ok) return;
  els.send.disabled = true;
  els.send.textContent = files.length ? 'Отправляю с вложениями…' : 'Отправляю…';
  try {
    const sig = signatureForMail();
    await sendMail({ to, cc, subject: subjectNow(), text, files, signatureHtml: sig.html, inline: sig.inline });
    files = [];
    renderFiles();
    els.send.textContent = '✓ Отправлено';
    soundPrefs().then((p) => playSound(p.sendSound));
    const time = new Date().toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
    els.sentText.textContent = `✓ Отправлено в ${time} → ${to.join(', ')}`;
    document.dispatchEvent(new CustomEvent('letter-sent', { detail: { id: ctx.id } })); // «Задача» уберёт напоминание «письмо — если ещё не написали»
    els.sent.hidden = false;
    toast('📨 Письмо отправлено — оно в «Отправленных»');
  } catch (e) {
    els.send.disabled = false;
    els.send.textContent = '📨 Отправить';
  els.sent.hidden = true;
    toast(`Не отправилось: ${e.message}`, 'err');
  }
});

// Проверить, что письмо ушло: папка «Отправленные» в Outlook — новое письмо в ней первым
els.sentOpen.addEventListener('click', () => openSent());

/** Новое письмо в Outlook Web App — сразу с адресами, темой и текстом. */
els.compose.addEventListener('click', async () => {
  const q = new URLSearchParams({ path: '/mail/action/compose' });
  const to = parseEmails(els.to.value).join(';');
  const ccList = parseEmails(els.cc.value);
  const cc = ccList.join(';');
  if (to) q.set('to', to);
  if (cc) q.set('cc', cc);
  q.set('subject', subjectNow());
  // Outlook (mail.rwb.ru) «Копию» из ссылки не берёт — адреса заранее в буфер (пока панель в фокусе), потом впишем сами
  if (cc) { try { await navigator.clipboard.writeText(ccList.join('; ')); } catch { /* без буфера — только автозаполнение */ } }
  q.set('body', plain(bodyEd.get() || fillTemplate(DEFAULT_BODY, ctx)));
  const tab = await chrome.tabs.create({ url: `${OWA_ORIGIN}/owa/?${q.toString().replace(/\+/g, '%20')}` });
  if (ccList.length || files.length) await waitComplete(tab.id);
  if (ccList.length) await fillCcIn(tab.id, ccList);
  if (files.length) attachTo(tab.id);
});

/** Вписать адреса в «Копию» письма во вкладке Outlook. Не вышло — адреса уже в буфере: Ctrl+V в «Копию». */
async function fillCcIn(tabId, list) {
  let r = null;
  try {
    [{ result: r }] = await chrome.scripting.executeScript({ target: { tabId }, world: 'MAIN', func: owaFillCc, args: [list] });
  } catch { /* нет доступа к вкладке */ }
  if (r && r.ok) toast(`Копия: ${list.join(', ')} — проверьте поле «Копия» в письме`);
  else toast('Outlook не принял «Копию» из ссылки. Адреса уже скопированы — нажмите в поле «Копия» и Ctrl+V', 'warn');
}

/**
 * Выполняется на странице Outlook Web App: найти поле «Копия» (Cc) в новом письме и вписать адреса.
 * Поле ищем по подписи (aria-label / placeholder); если оно свёрнуто — жмём кнопку «Копия».
 */
async function owaFillCc(list) {
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const visible = (el) => el && el.offsetParent !== null;
  const label = (el) => [el.getAttribute('aria-label'), el.getAttribute('placeholder'), el.getAttribute('title')].filter(Boolean).join(' ');
  const isCc = (t) => /^\s*(копия|cc)(?![\p{L}])/iu.test(t) && !/скрыт|bcc/i.test(t); // \b не видит русские буквы
  const findField = () => [...document.querySelectorAll('input, textarea, [contenteditable="true"], [role="textbox"], [role="combobox"]')]
    .find((el) => visible(el) && isCc(label(el)));
  const findButton = () => [...document.querySelectorAll('button, [role="button"], a, span')]
    .find((el) => visible(el) && /^\s*(копия|cc)\s*$/i.test(el.textContent || '') && !el.querySelector('input'));
  let field = null;
  for (let i = 0; i < 30 && !field; i++) {
    field = findField();
    if (!field) { const b = findButton(); if (b && i % 5 === 0) b.click(); await sleep(500); }
  }
  if (!field) return { ok: false };
  field.focus();
  field.click();
  const text = list.join('; ') + '; ';
  if (!document.execCommand('insertText', false, text)) {
    if ('value' in field) field.value = text; else field.textContent = text;
    field.dispatchEvent(new Event('input', { bubbles: true }));
  }
  field.dispatchEvent(new KeyboardEvent('keydown', { key: ';', bubbles: true }));
  field.dispatchEvent(new KeyboardEvent('keyup', { key: ';', bubbles: true }));
  await sleep(300);
  field.blur();
  return { ok: true };
}

/**
 * Выполняется на странице Outlook Web App: вложить файлы в открытое письмо.
 * Жмём «Вложить» — Outlook сам создаёт поле выбора файла и «кликает» по нему; этот клик перехватываем
 * и отдаём файлы вместо окна выбора. Нет кнопки — «перетаскиваем» файлы на текст письма.
 */
async function owaAttach(list) {
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const visible = (el) => el && el.offsetParent !== null;
  const make = () => list.map((f) => {
    const bin = atob(f.b64);
    const u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return new File([u8], f.name, { type: f.type || 'application/octet-stream' });
  });
  const put = (input) => {
    const dt = new DataTransfer();
    make().forEach((f) => dt.items.add(f));
    input.files = dt.files;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  };
  // имена файлов появились в письме — значит, вложились (длинные Outlook обрезает — сверяем начало)
  const attached = () => list.every((f) => document.body.innerText.includes(f.name.slice(0, 20)));
  const findBtn = () => [...document.querySelectorAll('button[title="Вложить"], button[aria-label="Вложить"], button[title="Attach"], button[aria-label="Attach"]')].find(visible);
  // письмо открывается не сразу — ждём кнопку до 15 секунд
  for (let i = 0; i < 60; i++) {
    const btn = findBtn();
    if (btn) {
      let done = false;
      const orig = HTMLInputElement.prototype.click;
      HTMLInputElement.prototype.click = function () {
        if (this.type === 'file' && !done) { done = true; put(this); return undefined; }
        return orig.apply(this, arguments);
      };
      try {
        btn.click();
        await sleep(400);
        if (!done) {
          // вместо окна — меню «Компьютер / OneDrive…»
          const item = [...document.querySelectorAll('[role="menuitem"], button, span, div')].find((el) => !el.children.length && visible(el) &&
            /^(Компьютер|Этот компьютер|Обзор компьютера|Обзор этого компьютера|Computer|Browse this computer)$/i.test(el.textContent.trim()));
          if (item) { item.click(); await sleep(400); }
        }
        if (!done) {
          const input = [...document.querySelectorAll('input[type="file"]')].pop();
          if (input) { put(input); done = true; }
        }
      } finally {
        HTMLInputElement.prototype.click = orig;
      }
      if (done) {
        for (let w = 0; w < 20; w++) { await sleep(300); if (attached()) return { ok: true }; }
        return { ok: true, unconfirmed: true };
      }
    }
    await sleep(250);
  }
  // запасной путь — «перетащить» файлы на текст письма
  const body = [...document.querySelectorAll('[contenteditable="true"]')].find(visible);
  if (body) {
    const dt = new DataTransfer();
    make().forEach((f) => dt.items.add(f));
    for (const type of ['dragenter', 'dragover', 'drop']) body.dispatchEvent(new DragEvent(type, { bubbles: true, cancelable: true, dataTransfer: dt }));
    return { ok: true, unconfirmed: true };
  }
  return { ok: false };
}

// ---------- Настройка для подрядчика ----------
// Открывается под добавленной задачей («✏️ Настроить») или во вкладке «Настройки» (host — куда поставить редактор).
function openEditor(contractor = ctx && ctx.contractor, { host = null, onDone = null, textId = null } = {}) {
  if (editing && editing.host && editing.host !== host) closeEditor();
  // textId: null — письмо подрядчика (адреса + обычный текст); '' — новый текст библиотеки; id — текст библиотеки
  const kind = textId === null ? 'letter' : 'text';
  editing = { contractor, host, onDone, ctx: host ? null : ctx, kind, textId };
  const l = letters[contractor] || {};
  const t = texts.find((x) => x.id === textId);
  (host || els.box).appendChild(els.editor);
  els.eAddr.hidden = kind === 'text';
  els.eNameRow.hidden = kind !== 'text';
  els.eName.value = t ? t.name : '';
  els.eTo.value = parseEmails(l.to).join('\n');
  els.eCc.value = parseEmails(l.cc).join('\n');
  tplEd.set(kind === 'text' ? (t ? t.body : DEFAULT_BODY) : (l.body || DEFAULT_BODY));
  els.remove.hidden = kind === 'text' ? !t : !letters[contractor];
  els.remove.textContent = kind === 'text' ? '🗑 Удалить этот текст' : '🗑 Удалить почты и шаблон';
  els.editor.hidden = false;
  if (!host) {
    els.view.hidden = true;
    els.empty.hidden = true;
    els.edit.hidden = true;
  }
  (kind === 'text' ? els.eName : els.eTo).focus();
}
function closeEditor() {
  els.editor.hidden = true;
  els.edit.hidden = false;
  const was = editing;
  editing = null;
  if (was && was.host) {
    els.box.appendChild(els.editor);
    if (was.onDone) was.onDone();
  } else if (was && ctx) render();
}

els.edit.addEventListener('click', () => openEditor());
els.empty.addEventListener('click', (e) => { if (e.target.closest('button')) openEditor(); });
els.cancel.addEventListener('click', () => closeEditor());
els.save.addEventListener('click', async () => {
  if (editing.kind === 'text') return saveText();
  const who = editing.contractor;
  const to = parseEmails(els.eTo.value);
  const cc = parseEmails(els.eCc.value);
  await loadLetters();
  letters[who] = { to: to.join('; '), cc: cc.join('; '), body: tplEd.get().trim() || DEFAULT_BODY };
  await chrome.storage.local.set({ letters });
  closeEditor();
  if (ctx && ctx.contractor === who && !els.box.hidden) render();
  toast(`Письмо для ${who} сохранено`);
});

// Стереть почты и шаблон подрядчика целиком
els.remove.addEventListener('click', async () => {
  if (editing.kind === 'text') return removeText();
  const who = editing.contractor;
  if (!await ask({ title: `Удалить письмо для ${who}?`, text: 'Почты и текст письма этого подрядчика сотрутся.', ok: 'Удалить', danger: true })) return;
  await loadLetters();
  delete letters[who];
  await chrome.storage.local.set({ letters });
  closeEditor();
  if (ctx && ctx.contractor === who && !els.box.hidden) render();
  toast(`🗑 Письмо для ${who} удалено`);
});

/**
 * Вкладка «Настройки»: все подрядчики — у кого письмо уже настроено и у кого ещё нет. Нажали — редактор прямо там.
 * contractors — из справочника таблицы (может быть пустым — тогда только уже настроенные).
 */
export async function renderLetterList(box, contractors = []) {
  await loadLetters();
  const names = [...new Set([...Object.keys(letters), ...contractors])].filter(Boolean);
  names.sort((a, b) => (Boolean(letters[b]) - Boolean(letters[a])) || a.localeCompare(b, 'ru'));
  if (editing && editing.host && box.contains(editing.host)) closeEditor();
  box.innerHTML = '';
  if (!names.length) { box.innerHTML = '<p class="field-note">Подрядчиков пока нет — они берутся из «Списков» таблицы.</p>'; return; }
  names.forEach((name) => {
    const l = letters[name];
    const item = document.createElement('div');
    item.className = `ml-item lt-item${l ? '' : ' unset'}`;
    const head = document.createElement('div');
    head.className = 'ml-head';
    const info = document.createElement('div');
    info.className = 'lt-info';
    const b = document.createElement('b'); b.textContent = name;
    const sub = document.createElement('span'); sub.className = 'ml-when';
    sub.textContent = l ? (parseEmails(l.to).join(', ') || 'без адреса') : 'ещё не настроено';
    info.append(b, sub);
    const btn = document.createElement('button');
    btn.type = 'button'; btn.className = 'btn ghost small';
    btn.textContent = l ? '✏️ Изменить' : 'Настроить';
    btn.addEventListener('click', () => {
      if (editing && editing.host === item) return closeEditor();
      openEditor(name, { host: item, onDone: () => renderLetterList(box, contractors) });
    });
    head.append(info, btn);
    item.append(head);
    box.append(item);
  });
}

// Вставили ссылки — текст письма обновляется сразу
els.links.addEventListener('input', () => { if (ctx && letters[ctx.contractor]) fillBody(); });

// ---------- Вставить переменную туда, где курсор ----------
document.querySelector('#lEditor .var-chips').addEventListener('mousedown', (e) => { if (e.target.closest('button')) e.preventDefault(); });
document.querySelector('#lEditor .var-chips').addEventListener('click', (e) => {
  const b = e.target.closest('button[data-var]');
  if (b) tplEd.insertVar(b.dataset.var);
});
document.querySelector('#lDefaultBody').addEventListener('click', () => {
  tplEd.set(DEFAULT_BODY);
  tplEd.focus();
});

// ---------- Тексты писем (библиотека) ----------
async function saveText() {
  const name = els.eName.value.trim();
  if (!name) { els.eName.focus(); return toast('Назовите текст — например, «Новые строки для веба»', 'warn'); }
  await loadLetters();
  if (texts.some((t) => t.name === name && t.id !== editing.textId)) return toast(`Текст «${name}» уже есть — выберите другое название`, 'warn');
  const body = tplEd.get().trim() || DEFAULT_BODY;
  const i = texts.findIndex((t) => t.id === editing.textId);
  if (i !== -1) texts[i] = { ...texts[i], name, body };
  else texts.push({ id: `t${Date.now().toString(36)}`, name, body });
  await saveTexts();
  closeEditor();
  if (ctx && !els.box.hidden && letters[ctx.contractor]) render();
  toast(`Текст «${name}» сохранён`);
}
async function removeText() {
  const t = texts.find((x) => x.id === editing.textId);
  if (!t || !await ask({ title: `Удалить текст «${t.name}»?`, text: 'Шаблоны задач с этим текстом будут брать обычный текст подрядчика.', ok: 'Удалить', danger: true })) return;
  await loadLetters();
  texts = texts.filter((x) => x.id !== t.id);
  await saveTexts();
  closeEditor();
  if (ctx && !els.box.hidden && letters[ctx.contractor]) render();
  toast(`🗑 Текст «${t.name}» удалён`);
}

/** Вкладка «Настройки» → «Тексты писем»: список, «Изменить», «＋ Новый текст» — редактор прямо там. */
export async function renderTextList(box, newBtn) {
  await loadLetters();
  if (editing && editing.host && box.contains(editing.host)) closeEditor();
  box.innerHTML = texts.length ? '' : '<p class="field-note">Пока только обычный текст у каждого подрядчика. Нажмите «＋ Новый текст».</p>';
  const again = () => renderTextList(box, newBtn);
  texts.forEach((t) => {
    const item = document.createElement('div');
    item.className = 'ml-item lt-item';
    const head = document.createElement('div');
    head.className = 'ml-head';
    const info = document.createElement('div');
    info.className = 'lt-info';
    const b = document.createElement('b'); b.textContent = t.name;
    const sub = document.createElement('span'); sub.className = 'ml-when';
    sub.textContent = t.body.replace(/\*\*/g, '').replace(/\{([а-яё]+)\}/gi, '‹$1›').replace(/\s+/g, ' ').slice(0, 90);
    info.append(b, sub);
    const btn = document.createElement('button');
    btn.type = 'button'; btn.className = 'btn ghost small'; btn.textContent = '✏️ Изменить';
    btn.addEventListener('click', () => {
      if (editing && editing.host === item) return closeEditor();
      openEditor(null, { host: item, onDone: again, textId: t.id });
    });
    head.append(info, btn);
    item.append(head);
    box.append(item);
  });
  newBtn.onclick = async () => {
    if (editing) closeEditor();
    await renderTextList(box, newBtn);
    const item = document.createElement('div');
    item.className = 'ml-item lt-item lt-newbox';
    box.append(item);
    openEditor(null, { host: item, onDone: again, textId: '' });
  };
}
