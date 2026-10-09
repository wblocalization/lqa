// Письмо подрядчику — на экране «Добавлено»: кому (почты подрядчика), копия и текст по шаблону.
// Скопировать по частям (панель открыта рядом с Outlook) или открыть новое письмо в Outlook уже заполненным.
// Почты и шаблоны — свои у каждого подрядчика, хранятся в расширении (chrome.storage, ключ «letters»),
// делятся с коллегами вместе с шаблонами задач (⚙️ → «Скачать файлом»).
import { toast } from './core.js';
import { OWA_ORIGIN } from './outlook.js';

const $ = (s) => document.querySelector(s);
const els = {
  box: $('#letter'), contractor: $('#lContractor'), edit: $('#lEdit'), view: $('#lView'), empty: $('#lEmpty'),
  subject: $('#lSubject'), copySubject: $('#lCopySubject'), to: $('#lTo'), copyTo: $('#lCopyTo'), ccRow: $('#lCcRow'), cc: $('#lCc'), copyCc: $('#lCopyCc'),
  body: $('#lBody'), copyBody: $('#lCopyBody'), compose: $('#lCompose'), links: $('#lLinks'),
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
  els.links.value = ''; // ссылки — у каждой задачи свои
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
