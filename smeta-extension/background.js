// Клик по иконке расширения открывает боковую панель — она не закрывается,
// пока переключаешься между Outlook и ВБ Диском.
// Раз в полчаса обновляет цифру просроченных задач на иконке, а по будням около 10:00 напоминает о сроках (remind.js).
import { loadSettings, isConfigured, settings, api } from './core.js';
import { remindDue, nextRemindAt } from './remind.js';

chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch(console.error);

// Звук напоминания: фоновая часть играть не умеет — через невидимую страницу offscreen.html
async function playInBackground(sound) {
  try {
    const has = (await chrome.runtime.getContexts({ contextTypes: ['OFFSCREEN_DOCUMENT'] })).length > 0;
    if (!has) await chrome.offscreen.createDocument({ url: 'offscreen.html', reasons: ['AUDIO_PLAYBACK'], justification: 'Звук напоминания о сроках' });
    await chrome.runtime.sendMessage({ type: 'play-sound', sound });
  } catch { /* без звука — уведомление всё равно покажется */ }
}

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
    await remindDue(r, { play: playInBackground });
  } catch {
    // нет сети или таблица недоступна — попробуем в следующий раз
  }
}

chrome.runtime.onInstalled.addListener(() => {
  chrome.alarms.create('badge', { periodInMinutes: 30 });
  refreshBadge();
  scheduleRemind();
});
chrome.runtime.onStartup.addListener(() => { refreshBadge(); scheduleRemind(); });
chrome.alarms.onAlarm.addListener((a) => {
  if (a.name === 'badge') refreshBadge();
  if (a.name === 'remind') refreshBadge().then(scheduleRemind);
});

// Будильник ровно на выбранное время напоминания; поменяли время — переставляем
async function scheduleRemind() {
  const when = await nextRemindAt();
  await chrome.alarms.clear('remind');
  if (when) chrome.alarms.create('remind', { when });
}
chrome.storage.onChanged.addListener((ch) => { if (ch.remind) scheduleRemind(); });

// Нажали на напоминание — панель на «Моих задачах»
chrome.notifications.onClicked.addListener(async (id) => {
  chrome.notifications.clear(id);
  await chrome.storage.local.set({ openTab: 'mine' });
  try {
    const [w] = await chrome.windows.getAll({ windowTypes: ['normal'] });
    if (w) { await chrome.windows.update(w.id, { focused: true }); await chrome.sidePanel.open({ windowId: w.id }); }
  } catch { /* Chrome не дал открыть панель сам — откроется на «Моих задачах», когда нажмёте на значок */ }
});
