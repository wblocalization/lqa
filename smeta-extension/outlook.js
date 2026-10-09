// «✉️ Почта»: открыть Outlook (mail.rwb.ru/owa) и найти письма по номеру задачи.
// В старом Outlook Web App поиск не задаётся адресом, поэтому вписываем номер в поле «Поиск» на самой странице
// и жмём Enter. Номер заранее копируется — если поиск не сработал, его можно просто вставить (Ctrl+V).
import { toast } from './core.js';

export const OWA_ORIGIN = 'https://mail.rwb.ru';
const OWA_URL = `${OWA_ORIGIN}/owa/`;

function waitLoaded(tabId, timeoutMs = 30000) {
  return new Promise((resolve) => {
    const timer = setTimeout(done, timeoutMs);
    function done() {
      clearTimeout(timer);
      chrome.tabs.onUpdated.removeListener(onUpdated);
      resolve();
    }
    function onUpdated(id, info) { if (id === tabId && info.status === 'complete') done(); }
    chrome.tabs.onUpdated.addListener(onUpdated);
  });
}

/** Вкладка Outlook: уже открытая (её и показываем) или новая. */
async function owaTab() {
  const [tab] = await chrome.tabs.query({ url: `${OWA_ORIGIN}/owa/*` });
  if (tab) {
    await chrome.tabs.update(tab.id, { active: true });
    await chrome.windows.update(tab.windowId, { focused: true });
    if (tab.status !== 'complete') await waitLoaded(tab.id);
    return tab;
  }
  const fresh = await chrome.tabs.create({ url: OWA_URL, active: true });
  await waitLoaded(fresh.id);
  return fresh;
}

/** Найти переписку по номеру задачи, например «LIT-26-2217». */
export async function openMail(id) {
  if (!id) return toast('У задачи нет номера — искать нечего', 'err');
  // Пока панель в фокусе — копируем номер: пригодится, если поиск не сработает
  let copied = false;
  try { await navigator.clipboard.writeText(id); copied = true; } catch { /* без копии */ }
  let ok = false;
  try {
    const tab = await owaTab();
    const [res] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, world: 'MAIN', func: owaSearch, args: [`"${id}"`] });
    ok = Boolean(res && res.result && res.result.ok);
  } catch {
    // нет доступа к странице или Outlook не открылся — остаётся вставить вручную
  }
  if (ok) toast(`✉️ Ищу в Outlook: ${id}`);
  else toast(copied ? `Номер ${id} скопирован — вставьте в поиск Outlook (Ctrl+V) и Enter` : `Найдите в Outlook: ${id}`, 'warn');
}

/** Выполняется на странице Outlook Web App: вписать запрос в «Поиск в почте и среди людей» и нажать Enter. */
async function owaSearch(query) {
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const visible = (el) => el && el.offsetParent !== null;
  const findInput = () => document.querySelector(
    'input[aria-label^="Поиск в почте"], input[aria-label^="Search Mail"], input._is_s[role="combobox"]');
  // Поиск запущен: слева «← Выйти из поиска» (или видна кнопка «Отменить поиск»)
  const searching = () => visible(document.querySelector('button[aria-label="Отменить поиск"], button[title="Выход из поиска"], button[aria-label="Exit search"]')) ||
    [...document.querySelectorAll('span, button, a, div[role="button"]')].some((el) => !el.children.length && /^(Выйти из поиска|Exit search)$/.test(el.textContent.trim()) && visible(el));
  // Человек сам кликнул или нажал клавишу (например, открыл письмо) — больше ничего не трогаем
  let userActed = false;
  const stop = (e) => { if (e.isTrusted) userActed = true; };
  addEventListener('mousedown', stop, true);
  addEventListener('keydown', stop, true);
  try {
    // Outlook грузится не сразу — ждём поле до 10 секунд
    for (let i = 0; i < 40 && !userActed; i++) {
      let input = findInput();
      if (!visible(input)) {
        // поле прячется за кнопкой-лупой
        const open = document.querySelector('button[aria-label^="Активировать текстовое окно поиска"], button[aria-label^="Activate Search"]');
        if (visible(open)) { open.click(); await sleep(150); }
        input = findInput();
      }
      if (!visible(input)) { await sleep(250); continue; }
      // Enter; если Outlook ещё не готов и поиск не начался — ещё раз, но не больше трёх попыток
      for (let attempt = 0; attempt < 3 && !userActed; attempt++) {
        input.focus();
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, query);
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        for (const type of ['keydown', 'keypress', 'keyup']) {
          input.dispatchEvent(new KeyboardEvent(type, { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true }));
        }
        for (let w = 0; w < 10; w++) {
          await sleep(250);
          if (userActed || searching()) return { ok: true };
        }
      }
      return { ok: userActed };
    }
    return { ok: userActed };
  } finally {
    removeEventListener('mousedown', stop, true);
    removeEventListener('keydown', stop, true);
  }
}

/**
 * «Проверить в «Отправленных»»: открыть Outlook и перейти в папку «Отправленные».
 * Адрес …#path=/mail/sentitems при первой загрузке Outlook теряет (после входа открываются «Входящие»),
 * поэтому, когда страница загрузилась, нажимаем на папку слева сами.
 */
export async function openSent() {
  try {
    const tab = await owaTab();
    const [res] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, world: 'MAIN', func: owaOpenSentInPage });
    if (!(res && res.result)) toast('Откройте папку «Отправленные» слева — письмо будет в ней первым', 'warn');
  } catch {
    chrome.tabs.create({ url: `${OWA_URL}#path=/mail/sentitems` });
  }
}

/** Выполняется на странице Outlook: дождаться списка папок и нажать «Отправленные». */
async function owaOpenSentInPage() {
  const NAMES = /^(Отправленные|Sent Items)$/i;
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  for (let i = 0; i < 40; i++) { // до 20 секунд — пока Outlook догрузится
    const el = [...document.querySelectorAll('[title], [role="treeitem"] span')]
      .find((e) => NAMES.test((e.getAttribute('title') || e.textContent || '').trim()) && e.offsetParent !== null);
    if (el) {
      const target = el.closest('[role="treeitem"]') || el;
      ['mousedown', 'mouseup', 'click'].forEach((type) => target.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window })));
      return true;
    }
    await sleep(500);
  }
  location.hash = '#path=/mail/sentitems';
  return false;
}
