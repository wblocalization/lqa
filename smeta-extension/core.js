// Общее для вкладок: настройки и связь с Apps Script таблицы.

// Куда класть сметы на ВБ Диске: общая папка команды локализации
export const DEFAULT_DISK_FOLDER = 'localization/Сметы';
export const settings = { url: '', token: '', user: '', manager: '', diskFolder: DEFAULT_DISK_FOLDER };

export async function loadSettings() {
  const saved = await chrome.storage.local.get('settings');
  Object.assign(settings, saved.settings || {});
  // В 0.7.0 по умолчанию была папка «Сметы» в корне — правильная папка команды: localization/Сметы
  if (!settings.diskFolder || settings.diskFolder === 'Сметы') settings.diskFolder = DEFAULT_DISK_FOLDER;
  return settings;
}

export async function saveSettings(patch) {
  Object.assign(settings, patch);
  await chrome.storage.local.set({ settings: { ...settings } });
}

const API_TIMEOUT_MS = 30000;

export const isConfigured = () => Boolean(settings.url && settings.token);

/** conf — адрес и токен, если они ещё не сохранены (окно настроек). */
export async function api(payload, conf = settings) {
  if (!conf.url || !conf.token) throw new Error('Заполните адрес и токен в настройках');
  // Не ждём вечно: если таблица молчит — честно говорим об этом
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), API_TIMEOUT_MS);
  let text;
  try {
    const res = await fetch(conf.url, {
      method: 'POST',
      body: JSON.stringify({ ...payload, token: conf.token, user: settings.manager || settings.user }),
      signal: ctrl.signal,
    });
    text = await res.text();
  } catch (e) {
    if (e.name === 'AbortError') throw new Error('таблица не ответила за 30 секунд. Проверьте адрес: откройте его в браузере — должно написать «Скрипт работает»');
    throw new Error('нет связи со скриптом. Проверьте адрес и интернет');
  } finally {
    clearTimeout(timer);
  }
  try {
    return JSON.parse(text);
  } catch {
    throw new Error(explainHtmlAnswer(text));
  }
}

/** Вместо ответа пришла страница Google — достаём из неё, что случилось. */
export function explainHtmlAnswer(html) {
  const plain = String(html)
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ').trim();
  if (/accounts\.google\.com|ServiceLogin|Sign in|Войдите|Вход в аккаунт/i.test(html)) {
    return 'Google просит войти: в развёртывании «У кого есть доступ» должно быть «Все» (не «Все, у кого есть аккаунт Google»)';
  }
  if (/already been declared|уже объявлен/i.test(plain)) {
    return 'в проекте скрипта код лежит дважды (например, «ВсёВОдном» и старые «Код»/«Smeta»). Оставьте что-то одно. ' +
      'Ошибка Google: ' + plain.slice(0, 160);
  }
  if (/doPost/i.test(plain)) return 'в скрипте нет приёмника для расширения (doPost). Вставлен ли новый код целиком? Ошибка Google: ' + plain.slice(0, 160);
  if (/authoriz|авториз|разрешени/i.test(plain)) return 'скрипту нужно разрешение: в Apps Script запустите любую функцию (например, onOpen) и нажмите «Разрешить». Ошибка Google: ' + plain.slice(0, 160);
  return 'скрипт ответил ошибкой, а не данными. Ошибка Google: ' + (plain.slice(0, 220) || 'пустой ответ');
}

/**
 * Справочники таблицы (подрядчики, продукты, статусы…): сразу из памяти расширения, свежие — в фоне.
 * onFresh(lists) вызывается, если в таблице что-то поменялось. Первый раз — ждём таблицу.
 */
export async function getLists(onFresh) {
  const key = `lists:${settings.url}`;
  let cached = null;
  try { cached = (await chrome.storage.local.get(key))[key]; } catch { /* нет — ждём таблицу */ }
  const fresh = api({ action: 'taskForm' }).then((r) => {
    if (!r.ok) throw new Error(r.error);
    chrome.storage.local.set({ [key]: r.lists }).catch(() => {});
    if (cached && onFresh && JSON.stringify(cached) !== JSON.stringify(r.lists)) onFresh(r.lists);
    return r.lists;
  });
  if (cached) {
    fresh.catch(() => {}); // не обновилось — работаем с тем, что есть
    return cached;
  }
  return fresh;
}

export function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

export async function readClipboard() {
  return (await navigator.clipboard.readText()).trim();
}

// ---------- Дедлайн из срока сдачи ----------
// Сколько дней от даты получения (нет — от сегодня) до срока сдачи. Так же считает таблица (Код.gs).
const DEADLINE_STEPS = [[0, 'ASAP'], [2, '1-2 дня'], [7, 'До недели'], [29, 'Больше недели'], [Infinity, 'Месяц и больше']];

export function deadlineFor(dueIso, dateIso, options, current) {
  if (!dueIso || current === 'Холд') return '';
  const day = (s) => { const [y, m, d] = s.split('-').map(Number); return Date.UTC(y, m - 1, d) / 864e5; };
  const now = new Date();
  const todayIso = `${now.getFullYear()}-${now.getMonth() + 1}-${now.getDate()}`;
  const days = Math.round(day(dueIso) - day(dateIso || todayIso));
  const label = DEADLINE_STEPS.find(([max]) => days <= max)[1].toLowerCase();
  return options.find((v) => v.toLowerCase() === label) || '';
}

/** Поставили срок сдачи или дату получения — дедлайн выбирается сам (его можно поменять руками). */
export function autoDeadline(select, due, date) {
  const update = () => {
    const v = deadlineFor(due.value, date.value, [...select.options].map((o) => o.value), select.value);
    if (v) select.value = v;
  };
  due.addEventListener('change', update);
  date.addEventListener('change', update);
}

/** Ссылка «Открыть в таблице»: показать, если адрес есть (старый скрипт его не присылает). */
export function setLink(a, url) {
  a.href = url || '#';
  a.hidden = !url;
}

/** Плашка сверху на пару секунд: «Шаблон обновлён» и т. п. */
let toastTimer = 0;
export function toast(text, kind = 'ok') {
  const el = document.getElementById('toast');
  if (!el) return;
  clearTimeout(toastTimer);
  el.textContent = text;
  el.className = `toast ${kind}`;
  el.hidden = false;
  toastTimer = setTimeout(() => {
    el.classList.add('out');
    toastTimer = setTimeout(() => { el.hidden = true; }, 300);
  }, 2600);
}

/** Запрос к таблице: ошибка — исключением, а «старый» скрипт — понятным текстом, что сделать. */
export async function call(payload) {
  const r = await api(payload);
  if (r.ok) return r;
  if (/Неизвестное действие/.test(r.error || '')) {
    throw new Error('веб-приложение ещё старое. В Apps Script: «Развернуть → Управление развёртываниями → ✏️ → Версия: новая → Развернуть»');
  }
  throw new Error(r.error || 'таблица ответила ошибкой');
}

/** Файл Excel из ответа таблицы (base64) — сразу в «Загрузки». */
export function saveXlsx(name, b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.append(a); // без этого Chrome иногда забывает имя файла
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

/** Скопировать со ссылками: в Band, почте и Google Docs вставится кликабельным. */
export async function copyRich(html, text) {
  try {
    await navigator.clipboard.write([new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob([text], { type: 'text/plain' }),
    })]);
  } catch {
    await navigator.clipboard.writeText(text);
  }
}

/** «102 641,64» */
export function money(x) {
  return Number(x || 0).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 3 });
}
