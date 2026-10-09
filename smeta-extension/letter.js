// Письмо подрядчику — на экране «Добавлено»: кому (почты подрядчика), копия и текст по шаблону.
// Скопировать по частям (панель открыта рядом с Outlook) или открыть новое письмо в Outlook уже заполненным.
// Почты и шаблоны — свои у каждого подрядчика, хранятся в расширении (chrome.storage, ключ «letters»),
// делятся с коллегами вместе с шаблонами задач (⚙️ → «Скачать файлом»).
import { toast } from './core.js';
import { OWA_ORIGIN } from './outlook.js';
import { sendMail } from './owa-send.js';

const $ = (s) => document.querySelector(s);
const els = {
  box: $('#letter'), contractor: $('#lContractor'), edit: $('#lEdit'), view: $('#lView'), empty: $('#lEmpty'),
  subject: $('#lSubject'), copySubject: $('#lCopySubject'), to: $('#lTo'), copyTo: $('#lCopyTo'), ccRow: $('#lCcRow'), cc: $('#lCc'), copyCc: $('#lCopyCc'),
  body: $('#lBody'), copyBody: $('#lCopyBody'), compose: $('#lCompose'), links: $('#lLinks'),
  send: $('#lSend'), drop: $('#lDrop'), fileInput: $('#lFileInput'), fileList: $('#lFileList'), attachOpen: $('#lAttachOpen'),
  editor: $('#lEditor'), eTo: $('#lETo'), eCc: $('#lECc'), eBody: $('#lEBody'), save: $('#lSave'), cancel: $('#lCancel'), remove: $('#lRemove'),
};

export const DEFAULT_BODY = 'Здравствуйте!\n\nВозьмите, пожалуйста, в работу.\nЯзыки: {языки}\nДедлайн: {срок}\n\nБольшое спасибо!';
// Прежние тексты по умолчанию — заменяем на новый сами (свой текст человека не трогаем)
const OLD_DEFAULTS = [
  'Добрый день!\n\nПросим взять в работу: {тема}\nЯзыки: {языки}\nСрок сдачи: {срок}\n\nСпасибо!',
  'Добрый день!\n\nВозьмите, пожалуйста, в работу.\n\nСпасибо!',
];

let letters = {};   // { [подрядчик]: { to, cc, body } }
let ctx = null;     // задача, которую только что добавили

export async function loadLetters() {
  const r = await chrome.storage.local.get('letters');
  letters = r.letters && typeof r.letters === 'object' ? r.letters : {};
  Object.values(letters).forEach((l) => { if (l && OLD_DEFAULTS.includes(l.body)) l.body = DEFAULT_BODY; });
  return letters;
}

/** Для файла «Скачать файлом» / «Загрузить из файла» — вместе с шаблонами задач. */
export const lettersForExport = () => letters;
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
export function fillTemplate(tpl, c) {
  const map = {
    тема: c.subject, номер: c.id, языки: (c.languages || []).map((l) => l.toLowerCase()).join(', '), коды: (c.codes || []).join(', '),
    срок: c.due, продукт: c.product, менеджер: c.manager, ссылка: c.link,
  };
  return String(tpl || '').replace(/\{([а-яё]+)\}/gi, (m, k) => (k.toLowerCase() in map ? map[k.toLowerCase()] || '' : m));
}

const ddmm = (iso) => { const [y, m, d] = String(iso || '').split('-'); return d ? `${d}.${m}.${y}` : ''; };

/** Показать блок письма для только что добавленной задачи. c: { contractor, id, subject, languages, codes, exactDeadline, product, manager, link } */
export async function showLetter(c) {
  await loadLetters();
  ctx = { ...c, due: ddmm(c.exactDeadline) };
  els.links.value = ''; // ссылки и вложения — у каждой задачи свои
  files = [];
  renderFiles();
  els.send.disabled = false;
  els.send.textContent = '📨 Отправить';
  els.contractor.textContent = c.contractor || 'подрядчику';
  els.box.hidden = !c.contractor;
  closeEditor();
  render();
}

function render() {
  const l = letters[ctx.contractor];
  const to = parseEmails(l && l.to);
  const cc = parseEmails(l && l.cc);
  els.empty.hidden = Boolean(l);
  els.view.hidden = !l;
  if (!l) return;
  const chips = (list) => list.length ? list.map((e) => `<span class="mail-chip">${e.replace(/</g, '&lt;')}</span>`).join('') : '<span class="muted">не указано</span>';
  els.subject.textContent = ctx.subject;
  els.to.innerHTML = chips(to);
  els.cc.innerHTML = chips(cc);
  els.ccRow.hidden = !cc.length;
  els.copyTo.disabled = !to.length;
  els.body.textContent = withMaterials(fillTemplate(l.body || DEFAULT_BODY, ctx), els.links.value);
}

async function copy(text, done) {
  try { await navigator.clipboard.writeText(text); toast(done); } catch { toast('Нет доступа к буферу — выделите и скопируйте вручную', 'err'); }
}

els.copySubject.addEventListener('click', () => copy(ctx.subject, '✓ Тема скопирована — вставьте в «Тема»'));
els.copyTo.addEventListener('click', () => copy(parseEmails(letters[ctx.contractor].to).join('; '), '✓ Адреса скопированы — вставьте в «Кому»'));
els.copyCc.addEventListener('click', () => copy(parseEmails(letters[ctx.contractor].cc).join('; '), '✓ Копия скопирована — вставьте в «Копия»'));
els.copyBody.addEventListener('click', () => copy(els.body.textContent, '✓ Текст письма скопирован'));

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

// «📨 Отправить» — прямо отсюда, с рабочей почты; перед отправкой — ещё раз показать, что уйдёт, и спросить
els.send.addEventListener('click', async () => {
  const l = letters[ctx.contractor] || {};
  const to = parseEmails(l.to);
  const cc = parseEmails(l.cc);
  if (!to.length) return toast('Нет адресов «Кому» — «✏️ Настроить»', 'err');
  const text = els.body.textContent;
  const ok = confirm(`Отправить письмо?\n\nКому: ${to.join(', ')}` + (cc.length ? `\nКопия: ${cc.join(', ')}` : '') +
    `\nТема: ${ctx.subject}` + (files.length ? `\nВложения: ${files.map((f) => f.name).join(', ')}` : '') +
    `\n\n${text.length > 400 ? `${text.slice(0, 400)}…` : text}`);
  if (!ok) return;
  els.send.disabled = true;
  els.send.textContent = files.length ? 'Отправляю с вложениями…' : 'Отправляю…';
  try {
    const r = await sendMail({ to, cc, subject: ctx.subject, text, files });
    files = [];
    renderFiles();
    els.send.textContent = '✓ Отправлено';
    toast(r.signature ? '📨 Письмо отправлено — оно в «Отправленных»' : '📨 Письмо отправлено (без подписи — её не нашла в Outlook)');
  } catch (e) {
    els.send.disabled = false;
    els.send.textContent = '📨 Отправить';
    toast(`Не отправилось: ${e.message}`, 'err');
  }
});

/** Новое письмо в Outlook Web App — сразу с адресами, темой и текстом. */
els.compose.addEventListener('click', async () => {
  const l = letters[ctx.contractor] || {};
  const q = new URLSearchParams({ path: '/mail/action/compose' });
  const to = parseEmails(l.to).join(';');
  const cc = parseEmails(l.cc).join(';');
  if (to) q.set('to', to);
  if (cc) q.set('cc', cc);
  q.set('subject', ctx.subject);
  q.set('body', els.body.textContent || fillTemplate(DEFAULT_BODY, ctx));
  const tab = await chrome.tabs.create({ url: `${OWA_ORIGIN}/owa/?${q.toString().replace(/\+/g, '%20')}` });
  if (files.length) {
    await waitComplete(tab.id);
    attachTo(tab.id);
  }
});

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
function openEditor() {
  const l = letters[ctx.contractor] || {};
  els.eTo.value = parseEmails(l.to).join('\n');
  els.eCc.value = parseEmails(l.cc).join('\n');
  els.eBody.value = l.body || DEFAULT_BODY;
  els.remove.hidden = !letters[ctx.contractor];
  els.editor.hidden = false;
  els.view.hidden = true;
  els.empty.hidden = true;
  els.edit.hidden = true;
  els.eTo.focus();
}
function closeEditor() {
  els.editor.hidden = true;
  els.edit.hidden = false;
}

els.edit.addEventListener('click', openEditor);
els.empty.addEventListener('click', (e) => { if (e.target.closest('button')) openEditor(); });
els.cancel.addEventListener('click', () => { closeEditor(); render(); });
els.save.addEventListener('click', async () => {
  const to = parseEmails(els.eTo.value);
  const cc = parseEmails(els.eCc.value);
  await loadLetters();
  letters[ctx.contractor] = { to: to.join('; '), cc: cc.join('; '), body: els.eBody.value.trim() || DEFAULT_BODY };
  await chrome.storage.local.set({ letters });
  closeEditor();
  render();
  toast(`✓ Письмо для ${ctx.contractor} сохранено`);
});

// Стереть почты и шаблон подрядчика целиком
els.remove.addEventListener('click', async () => {
  if (!confirm(`Удалить почты и шаблон письма для ${ctx.contractor}?`)) return;
  await loadLetters();
  delete letters[ctx.contractor];
  await chrome.storage.local.set({ letters });
  closeEditor();
  render();
  toast(`Почты и шаблон для ${ctx.contractor} удалены`);
});

// Вставили ссылки — текст письма обновляется сразу
els.links.addEventListener('input', () => { if (ctx && letters[ctx.contractor]) render(); });
