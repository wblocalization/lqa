// «✉️ Переписка»: открыть Outlook (mail.rwb.ru/owa) и найти письма по номеру задачи.
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
  // Outlook грузится не сразу — ждём поле до 10 секунд
  for (let i = 0; i < 40; i++) {
    let input = findInput();
    if (!visible(input)) {
      // поле прячется за кнопкой-лупой
      const open = document.querySelector('button[aria-label^="Активировать текстовое окно поиска"], button[aria-label^="Activate Search"]');
      if (visible(open)) { open.click(); await sleep(150); }
      input = findInput();
    }
    if (visible(input)) {
      // Поиск запустился — появляется «Отменить поиск» (title «Выход из поиска»). Нет — Outlook ещё не готов, жмём ещё раз
      for (let attempt = 0; attempt < 6; attempt++) {
        input.focus();
        input.click();
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, query);
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        for (const type of ['keydown', 'keypress', 'keyup']) {
          input.dispatchEvent(new KeyboardEvent(type, { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true }));
        }
        for (let w = 0; w < 6; w++) {
          await sleep(250);
          if (visible(document.querySelector('button[aria-label="Отменить поиск"], button[title="Выход из поиска"], button[aria-label="Exit search"]'))) return { ok: true };
        }
      }
      return { ok: false };
    }
    await sleep(250);
  }
  return { ok: false };
}
