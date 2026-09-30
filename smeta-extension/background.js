// Клик по иконке расширения открывает боковую панель — она не закрывается,
// пока переключаешься между Outlook и ВБ Диском.
// Раз в полчаса обновляет цифру просроченных задач на иконке.
import { loadSettings, isConfigured, settings, api } from './core.js';

chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(console.error);

async function refreshBadge() {
  await loadSettings();
  if (!isConfigured() || !settings.manager) return;
  try {
    const r = await api({ action: 'myTasks', manager: settings.manager });
    if (!r.ok) return;
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
