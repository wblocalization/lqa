// Клик по иконке — боковая панель. Правый клик по сообщению в Band → «📮 Создать задачу в трекере»:
// панель открывается с готовой формой (текст, автор и ссылка на сообщение).
// Кнопка «Создать тикет» прямо в треде (band.js) присылает сюда сообщения: разобрать заказ, создать задачи, открыть панель.
import { parseOrder, subtaskTitle, orderSummary } from './parse.js';
import { loadBoards, createPlan, logCreated, journalFor, fieldsFor, ensureDictionaries, rememberBz } from './tracker.js';

chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(console.error);

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({
      id: 'toTracker', title: '📮 Создать задачу в трекере',
      contexts: ['selection', 'page', 'link'], documentUrlPatterns: ['https://band.wb.ru/*'],
    });
  });
});

chrome.contextMenus.onClicked.addListener((info, tab) => {
  if (info.menuItemId !== 'toTracker' || !tab) return;
  // Открыть панель можно только сразу по клику — до любых await
  chrome.sidePanel.open({ windowId: tab.windowId }).catch(console.error);
  saveDraft(info, tab);
});

async function saveDraft(info, tab) {
  let post = null;
  try {
    // band.js запомнил, по какому сообщению кликнули
    const [res] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, func: () => window.__lqaBandPost || null });
    post = res && res.result;
  } catch {
    // нет доступа к странице — обойдёмся выделенным текстом
  }
  // Вкладку Band открыли до установки расширения — band.js в ней нет (post пустой), тоже только выделенный текст
  await chrome.storage.local.set({
    trackerDraft: {
      selection: (post && post.selection) || info.selectionText || '',
      text: (post && post.text) || '',
      author: (post && post.author) || '',
      link: (post && post.link) || '',
      postId: (post && post.postId) || '',
      tabId: tab.id,
      at: Date.now(),
    },
  });
}

// ---------- Кнопка «Создать тикет» в треде Band ----------
/** Что из заказа идёт в поля трекера: страны (по языкам), БЗ (по продукту), дедлайн. */
export const roleValues = (order) => ({ countries: [...new Set(order.languages.map((l) => l.country))], bz: order.product, due: order.due });

/** Описание задачи: что из заказа (продукт, языки, страны, ссылки) + текст сообщения + ссылка на него. */
export function describe(order, post) {
  const parts = [orderSummary(order), order.text.trim()];
  if (post.link) parts.push(`[Сообщение в Band](${post.link})${post.author ? ` — ${post.author}` : ''}`);
  return hardBreaks(parts.filter(Boolean).join('\n\n'));
}
function hardBreaks(md) {
  return String(md || '').split(/(```[\s\S]*?```)/).map((part, i) => (i % 2 ? part : part.replace(/([^\n])[ \t]*\n(?=[^\n])/g, '$1  \n'))).join('');
}

chrome.runtime.onMessage.addListener((msg, sender, reply) => {
  if (!msg || !msg.type) return false;
  if (msg.type === 'openPanel') {
    // Открыть панель можно только сразу по клику — до любых await
    const opened = chrome.sidePanel.open({ tabId: sender.tab.id }).then(() => true, () => false);
    chrome.storage.local.set({ trackerDraft: { ...msg.draft, tabId: sender.tab.id, at: Date.now() } });
    opened.then((ok) => reply({ ok }));
    return true;
  }
  (async () => {
    if (msg.type === 'plan') {
      const order = parseOrder(msg.post.text);
      const { current } = await loadBoards();
      const { subLangs = [], replyInThread = true } = await chrome.storage.local.get(['subLangs', 'replyInThread']);
      if (current) await ensureDictionaries(current); // первый раз и раз в неделю — варианты БЗ и стран из трекера
      const f = current ? await fieldsFor(current, roleValues(order)) : { picked: {}, missing: {}, bzOptions: [] };
      return { order, board: current && current.name, subLangs, replyInThread, journal: await journalFor(msg.post.postId), picked: f.picked, missing: f.missing, bzOptions: f.bzOptions || [] };
    }
    if (msg.type === 'create') {
      const { boards, current } = await loadBoards();
      if (!current) throw new Error('Нет типа задачи: создайте одну задачу нужного типа в tracker.wb.ru — расширение запомнит её');
      const order = parseOrder(msg.post.text);
      const langs = order.languages.filter((l) => (msg.subLangs || []).includes(l.code));
      const plan = {
        board: current, title: (msg.title || order.title).trim(), description: describe(order, msg.post),
        fields: (await fieldsFor(current, { ...roleValues(order), bzId: msg.bzId })).fields,
        subs: await Promise.all(langs.map(async (l) => ({
          lang: l.code, title: subtaskTitle(order, l), description: describe({ ...order, languages: [l] }, msg.post),
          fields: (await fieldsFor(current, { ...roleValues({ ...order, languages: [l] }), bzId: msg.bzId })).fields, // у подзадачи — страна её языка
        }))),
      };
      if (msg.bzId && order.product) await rememberBz(order.product, msg.bzId);
      const res = await createPlan(plan);
      await chrome.storage.local.set({ subLangs: msg.subLangs || [], replyInThread: !!msg.reply });
      await logCreated({ postId: msg.post.postId, link: msg.post.link, title: plan.title, main: res.main, subs: res.subs, board: current.name });
      return res;
    }
    throw new Error('Неизвестная команда');
  })().then((r) => reply({ ok: true, ...r }), (e) => reply({ ok: false, error: e.message }));
  return true;
});
