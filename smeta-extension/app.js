// Боковая панель: вкладки «Задача» / «Смета» и общие настройки.
import { settings, loadSettings, saveSettings, isConfigured } from './core.js';
import * as smeta from './smeta.js';
import * as task from './task.js';

const $ = (s) => document.querySelector(s);
const els = {
  settings: $('#settings'), setUrl: $('#setUrl'), setToken: $('#setToken'), setUser: $('#setUser'),
  openSettings: $('#openSettings'), closeSettings: $('#closeSettings'), saveSettings: $('#saveSettings'),
  tabs: [...document.querySelectorAll('.tab')],
};

// ---------- Вкладки ----------
function showTab(name) {
  els.tabs.forEach((t) => {
    const on = t.dataset.tab === name;
    t.setAttribute('aria-selected', String(on));
    $(`#${t.getAttribute('aria-controls')}`).hidden = !on;
  });
  if (name === 'task') task.initTaskTab();
  try { localStorage.setItem('tab', name); } catch { /* не страшно */ }
}
els.tabs.forEach((t) => t.addEventListener('click', () => showTab(t.dataset.tab)));

// ---------- Настройки ----------
function fillSettings() {
  els.setUrl.value = settings.url;
  els.setToken.value = settings.token;
  els.setUser.value = settings.user;
}

els.openSettings.addEventListener('click', () => {
  fillSettings();
  els.settings.hidden = !els.settings.hidden;
});
els.closeSettings.addEventListener('click', () => { els.settings.hidden = true; });
els.saveSettings.addEventListener('click', async () => {
  await saveSettings({ url: els.setUrl.value.trim(), token: els.setToken.value.trim(), user: els.setUser.value.trim() });
  els.settings.hidden = true;
  smeta.onSettingsSaved();
  task.onSettingsSaved();
});

await loadSettings();
fillSettings();
if (!isConfigured()) els.settings.hidden = false;
let startTab = 'task';
try { startTab = localStorage.getItem('tab') || 'task'; } catch { /* по умолчанию «Задача» */ }
showTab(startTab === 'smeta' ? 'smeta' : 'task');
