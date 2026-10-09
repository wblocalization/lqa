// Письмо подрядчику — на экране «Добавлено»: кому (почты подрядчика), копия и текст по шаблону.
// Скопировать по частям (панель открыта рядом с Outlook) или открыть новое письмо в Outlook уже заполненным.
// Почты и шаблоны — свои у каждого подрядчика, хранятся в расширении (chrome.storage, ключ «letters»),
// делятся с коллегами вместе с шаблонами задач (⚙️ → «Скачать файлом»).
import { toast } from './core.js';
import { OWA_ORIGIN } from './outlook.js';

const $ = (s) => document.querySelector(s);
const els = {
  box: $('#letter'), contractor: $('#lContractor'), edit: $('#lEdit'), view: $('#lView'), empty: $('#lEmpty'),
  to: $('#lTo'), copyTo: $('#lCopyTo'), ccRow: $('#lCcRow'), cc: $('#lCc'), copyCc: $('#lCopyCc'),
  body: $('#lBody'), copyBody: $('#lCopyBody'), compose: $('#lCompose'),
  editor: $('#lEditor'), eTo: $('#lETo'), eCc: $('#lECc'), eBody: $('#lEBody'), save: $('#lSave'), cancel: $('#lCancel'),
};

export const DEFAULT_BODY = 'Добрый день!\n\nПросим взять в работу: {тема}\nЯзыки: {языки}\nСрок сдачи: {срок}\n\nСпасибо!';

let letters = {};   // { [подрядчик]: { to, cc, body } }
let ctx = null;     // задача, которую только что добавили

export async function loadLetters() {
  const r = await chrome.storage.local.get('letters');
  letters = r.letters && typeof r.letters === 'object' ? r.letters : {};
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

/** Подставить в шаблон данные задачи: {тема} {номер} {языки} {коды} {срок} {продукт} {менеджер} {ссылка}. */
export function fillTemplate(tpl, c) {
  const map = {
    тема: c.subject, номер: c.id, языки: (c.languages || []).join(', '), коды: (c.codes || []).join(', '),
    срок: c.due, продукт: c.product, менеджер: c.manager, ссылка: c.link,
  };
  return String(tpl || '').replace(/\{([а-яё]+)\}/gi, (m, k) => (k.toLowerCase() in map ? map[k.toLowerCase()] || '' : m));
}

const ddmm = (iso) => { const [y, m, d] = String(iso || '').split('-'); return d ? `${d}.${m}.${y}` : ''; };

/** Показать блок письма для только что добавленной задачи. c: { contractor, id, subject, languages, codes, exactDeadline, product, manager, link } */
export async function showLetter(c) {
  await loadLetters();
  ctx = { ...c, due: ddmm(c.exactDeadline) };
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
  els.to.innerHTML = chips(to);
  els.cc.innerHTML = chips(cc);
  els.ccRow.hidden = !cc.length;
  els.copyTo.disabled = !to.length;
  els.body.textContent = fillTemplate(l.body || DEFAULT_BODY, ctx);
}

async function copy(text, done) {
  try { await navigator.clipboard.writeText(text); toast(done); } catch { toast('Нет доступа к буферу — выделите и скопируйте вручную', 'err'); }
}

els.copyTo.addEventListener('click', () => copy(parseEmails(letters[ctx.contractor].to).join('; '), '✓ Адреса скопированы — вставьте в «Кому»'));
els.copyCc.addEventListener('click', () => copy(parseEmails(letters[ctx.contractor].cc).join('; '), '✓ Копия скопирована — вставьте в «Копия»'));
els.copyBody.addEventListener('click', () => copy(els.body.textContent, '✓ Текст письма скопирован'));

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
  await chrome.tabs.create({ url: `${OWA_ORIGIN}/owa/?${q.toString().replace(/\+/g, '%20')}` });
});

// ---------- Настройка для подрядчика ----------
function openEditor() {
  const l = letters[ctx.contractor] || {};
  els.eTo.value = parseEmails(l.to).join('\n');
  els.eCc.value = parseEmails(l.cc).join('\n');
  els.eBody.value = l.body || DEFAULT_BODY;
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
