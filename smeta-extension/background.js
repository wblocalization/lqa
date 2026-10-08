// Клик по иконке расширения открывает боковую панель — она не закрывается,
// пока переключаешься между Outlook и ВБ Диском.
// Раз в полчаса обновляет цифру просроченных задач на иконке.
import { loadSettings, isConfigured, settings, api } from './core.js';

chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(console.error);

async function refreshBadge() {
  await loadSettings();
  if (!isConfigured() || !settings.manager) return;
  try {
    const r = await api({ action: 'myTasks', manager: settings.manager, filter: 'open' });
    if (!r.ok) return;
    // Заодно запоминаем список: вкладка «Мои задачи» покажет его сразу
    if (r.manager) await chrome.storage.local.set({ [`mine:${r.manager}:open`]: { ...r, filter: 'open' } });
    chrome.action.setBadgeText({ text: r.overdue ? String(r.overdue) : '' });
    chrome.action.setBadgeBackgroundColor({ color: '#C0262D' });
  } catch {
    // нет сети или таблица недоступна — попробуем в следующий раз
  }
}

chrome.runtime.onInstalled.addListener(() => {
  chrome.alarms.create('badge', { periodInMinutes: 30 });
  refreshBadge();
});
chrome.runtime.onStartup.addListener(refreshBadge);
chrome.alarms.onAlarm.addListener((a) => { if (a.name === 'badge') refreshBadge(); });

// ---------- Band → трекер ----------
// Правый клик по сообщению в Band → «📮 Создать задачу в трекере»: панель открывается с готовой формой.
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
  bandDraft(info, tab);
});

async function bandDraft(info, tab) {
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
