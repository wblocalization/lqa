// Личные шаблоны — в таблице, у каждого менеджера свои (лист «Личные шаблоны», скрытый).
// Шаблоны задач, письма подрядчикам и подпись живут в расширении как раньше, но каждое изменение
// уходит в таблицу, а при запуске панели берётся оттуда. Переустановили расширение, другой компьютер —
// выберите себя в ⚙️ «Кто вы», и всё вернётся.
import { settings, isConfigured, api, toast } from './core.js';

const KEYS = ['templates', 'letters', 'signature'];
let synced = {};      // что сейчас в таблице, по ключам (JSON) — чтобы не отправлять то же самое обратно
let who = '';         // чьи шаблоны
let ready = false;    // таблица ответила (старый скрипт без «personalGet» — работаем только в браузере)
let timer = null;
let warned = false;

const hasData = (v) => (Array.isArray(v) ? v.length > 0 : v && typeof v === 'object' ? Object.keys(v).length > 0 : Boolean(v));

/**
 * Взять шаблоны из таблицы. onApplied() — если что-то поменялось и экраны надо перерисовать.
 * В таблице ещё пусто, а в браузере шаблоны есть — переносим их в таблицу.
 */
export async function cloudSync(onApplied) {
  ready = false;
  clearTimeout(timer);
  if (!isConfigured() || !settings.manager) return;
  who = settings.manager;
  let r;
  try { r = await api({ action: 'personalGet', manager: who }); } catch { return; }
  if (!r || !r.ok || who !== settings.manager) return;
  const local = await chrome.storage.local.get(KEYS);
  if (r.data && typeof r.data === 'object') {
    const patch = {};
    KEYS.forEach((k) => {
      if (!(k in r.data)) return;
      synced[k] = JSON.stringify(r.data[k]);
      if (JSON.stringify(local[k]) !== synced[k]) patch[k] = r.data[k];
    });
    ready = true;
    if (Object.keys(patch).length) {
      await chrome.storage.local.set(patch);
      if (onApplied) onApplied(Object.keys(patch));
    }
  } else {
    synced = {};
    ready = true;
    if (KEYS.some((k) => hasData(local[k]))) push();
  }
}

/** Отправить в таблицу всё, что сейчас в браузере (через секунду после последнего изменения). */
function push() {
  clearTimeout(timer);
  timer = setTimeout(async () => {
    if (!ready || who !== settings.manager) return;
    const local = await chrome.storage.local.get(KEYS);
    const data = {};
    KEYS.forEach((k) => { if (local[k] !== undefined) data[k] = local[k]; });
    try {
      const r = await api({ action: 'personalSave', manager: who, data });
      if (!r.ok) throw new Error(r.error);
      KEYS.forEach((k) => { synced[k] = JSON.stringify(data[k]); });
      warned = false;
    } catch (e) {
      if (!warned) toast(`Шаблоны сохранились только в этом браузере, в таблицу — нет: ${e.message}`, 'warn');
      warned = true;
    }
  }, 1000);
}

chrome.storage.onChanged.addListener((changes) => {
  if (!ready) return;
  if (KEYS.some((k) => k in changes && JSON.stringify(changes[k].newValue) !== synced[k])) push();
});
