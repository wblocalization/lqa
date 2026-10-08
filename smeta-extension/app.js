// Боковая панель: внизу разделы «Задача» / «Мои» / «Смета» / «Отчёты» / «Письма», сверху — название и настройки.
import { settings, loadSettings, saveSettings, isConfigured, api, esc, DEFAULT_DISK_FOLDER, setLink, getLists } from './core.js';
import * as smeta from './smeta.js';
import * as task from './task.js';
import * as mine from './mine.js';
import * as reports from './reports.js';
import { loadMail } from './mail.js';

const $ = (s) => document.querySelector(s);
const els = {
  settings: $('#settings'), setUrl: $('#setUrl'), setToken: $('#setToken'), setManager: $('#setManager'), setDiskFolder: $('#setDiskFolder'), managerMsg: $('#setManagerMsg'), retryManagers: $('#retryManagers'),
  openSettings: $('#openSettings'), closeSettings: $('#closeSettings'), saveSettings: $('#saveSettings'),
  tabs: [...document.querySelectorAll('.tabbar .tab')], title: $('#pageTitle'),
};

// ---------- Вкладки ----------
function showTab(name) {
  els.tabs.forEach((t) => {
    const on = t.dataset.tab === name;
    t.setAttribute('aria-selected', String(on));
    $(`#${t.getAttribute('aria-controls')}`).hidden = !on;
  });
  const titles = { task: 'Новая задача', mine: 'Мои задачи', smeta: 'Смета', reports: 'Отчёты', mail: 'Письма' };
  els.title.textContent = titles[name];
  if (name === 'task') task.initTaskTab();
  if (name === 'mine') mine.loadMine();
  if (name === 'reports') { reports.showMenu(); reports.initReports(); }
  if (name === 'mail') loadMail();
  window.scrollTo(0, 0);
  try { localStorage.setItem('tab', name); } catch { /* не страшно */ }
}
els.tabs.forEach((t) => t.addEventListener('click', () => showTab(t.dataset.tab)));

// ---------- Настройки ----------
/**
 * Папка для смет: путь («localization/Сметы») или ссылка на папку из адресной строки ВБ Диска
 * (…/apps/files/files/123?dir=/localization/Сметы) — из ссылки берём путь после dir=.
 */
function diskFolderFromInput(value) {
  let v = String(value || '').trim();
  const m = v.match(/[?&]dir=([^&#]*)/);
  if (m) {
    try { v = decodeURIComponent(m[1].replace(/\+/g, ' ')); } catch { v = m[1]; }
  }
  return v.split('/').map((x) => x.trim()).filter(Boolean).join('/');
}
// «Кто вы» — менеджер из «Списков»: подставляется в новые задачи и во вкладку «Мои задачи».
let managersSeq = 0;
async function loadManagers() {
  const seq = ++managersSeq;
  const conf = { url: els.setUrl.value.trim(), token: els.setToken.value.trim() };
  const current = els.setManager.value || settings.manager || settings.user;
  if (!conf.url || !conf.token) {
    els.setManager.innerHTML = '<option value="">Сначала адрес и токен</option>';
    return;
  }
  els.setManager.innerHTML = '<option value="">Загружаю менеджеров…</option>';
  els.managerMsg.hidden = true;
  try {
    // «managers» — быстрый запрос; старый скрипт его не знает — тогда берём из справочников формы
    let r = await api({ action: 'managers' }, conf);
    if (!r.ok && /Неизвестное действие/.test(r.error || '')) {
      const f = await api({ action: 'taskForm' }, conf);
      r = f.ok ? { ok: true, managers: f.lists.managers } : f;
    }
    if (!r.ok) throw new Error(r.error);
    if (seq !== managersSeq) return;
    els.setManager.innerHTML = '<option value="">— выберите себя —</option>' +
      r.managers.map((m) => `<option value="${esc(m)}"${m === current ? ' selected' : ''}>${esc(m)}</option>`).join('');
  } catch (e) {
    if (seq !== managersSeq) return;
    els.setManager.innerHTML = '<option value="">Не загрузилось</option>';
    els.managerMsg.textContent = `Не загрузилось: ${e.message}`;
    els.managerMsg.hidden = false;
  }
}

function fillSettings() {
  els.setUrl.value = settings.url;
  els.setToken.value = settings.token;
  els.setDiskFolder.value = settings.diskFolder || '';
  loadManagers();
}
[els.setUrl, els.setToken].forEach((el) => el.addEventListener('change', loadManagers));
els.retryManagers.addEventListener('click', loadManagers);

els.openSettings.addEventListener('click', () => {
  if (els.settings.hidden) fillSettings();
  els.settings.hidden = !els.settings.hidden;
});
els.closeSettings.addEventListener('click', () => { els.settings.hidden = true; });
els.saveSettings.addEventListener('click', async () => {
  const patch = { url: els.setUrl.value.trim(), token: els.setToken.value.trim(),
    diskFolder: diskFolderFromInput(els.setDiskFolder.value) || DEFAULT_DISK_FOLDER };
  if (els.setManager.value) Object.assign(patch, { manager: els.setManager.value, user: els.setManager.value });
  await saveSettings(patch);
  els.settings.hidden = true;
  showTableLink();
  smeta.onSettingsSaved();
  task.onSettingsSaved();
  mine.loadMine();
});

// Статус — цветом: выбрали другой — цвет сменился сразу
document.addEventListener('change', (e) => {
  if (e.target.matches && e.target.matches('select.status-sel')) e.target.dataset.status = e.target.value;
  if (e.target.matches && e.target.matches('select.delivery-sel')) e.target.dataset.delivery = e.target.value;
});

// ---------- Тема: светлая / тёмная ----------
// Пока не нажимали — как в системе. theme.js ставит выбранную ещё до отрисовки.
const dark = () => document.documentElement.dataset.theme === 'dark' ||
  (!document.documentElement.dataset.theme && matchMedia('(prefers-color-scheme: dark)').matches);
$('#themeBtn').addEventListener('click', () => {
  const next = dark() ? 'light' : 'dark';
  document.documentElement.dataset.theme = next;
  try { localStorage.setItem('theme', next); } catch { /* на этот раз */ }
});

// ---------- Кнопка «Таблица ↗» ----------
// Адрес листа задач приходит со справочниками (старый скрипт его не присылает — тогда кнопки нет).
async function showTableLink() {
  const key = `lists:${settings.url}`;
  let lists = null;
  try { lists = (await chrome.storage.local.get(key))[key]; } catch { /* нет — без кнопки */ }
  setLink($('#openTable'), lists && lists.tableUrl);
}
chrome.storage.onChanged.addListener((changes) => { if (changes[`lists:${settings.url}`]) showTableLink(); });

// Высота шапки — чтобы тема письма прилипала точно под ней
const head = document.querySelector('.top');
const setHeadH = () => document.documentElement.style.setProperty('--head-h', head.offsetHeight + 'px');
setHeadH();
window.addEventListener('resize', setHeadH);

await loadSettings();
showTableLink();
if (!isConfigured() || !settings.manager) {
  fillSettings();
  els.settings.hidden = false;
}
let startTab = 'task';
try { startTab = localStorage.getItem('tab') || 'task'; } catch { /* по умолчанию «Новая задача» */ }
showTab(['smeta', 'mine', 'reports', 'mail'].includes(startTab) ? startTab : 'task');
// Вкладка «Новая задача» и так грузит справочники; с других — подтянем их в фоне ради кнопки «Таблица ↗»
if (isConfigured() && startTab !== 'task') getLists().catch(() => {});
