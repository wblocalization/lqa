// Общее для вкладок: настройки и связь с Apps Script таблицы.

export const settings = { url: '', token: '', user: '', manager: '' };

export async function loadSettings() {
  const saved = await chrome.storage.local.get('settings');
  Object.assign(settings, saved.settings || {});
  return settings;
}

export async function saveSettings(patch) {
  Object.assign(settings, patch);
  await chrome.storage.local.set({ settings: { ...settings } });
}

export const isConfigured = () => Boolean(settings.url && settings.token);

export async function api(payload) {
  if (!isConfigured()) throw new Error('Заполните адрес и токен в настройках');
  const res = await fetch(settings.url, {
    method: 'POST',
    body: JSON.stringify({ ...payload, token: settings.token, user: settings.user }),
  });
  const text = await res.text();
  try {
    return JSON.parse(text);
  } catch {
    throw new Error('Скрипт ответил не JSON — проверьте адрес и что доступ открыт «Всем»');
  }
}

export function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

export async function readClipboard() {
  return (await navigator.clipboard.readText()).trim();
}
