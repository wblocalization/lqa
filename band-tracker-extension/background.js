// Клик по иконке — боковая панель. Правый клик по сообщению в Band → «📮 Создать задачу в трекере»:
// панель открывается с готовой формой (текст, автор и ссылка на сообщение).
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
      at: Date.now(),
    },
  });
}
