// «Жесты» в панели: переключатель в «Настройках», кружок с камерой в углу и что делает каждый жест.
import { toast } from './core.js';
import { playSound } from './sounds.js';

const $ = (s) => document.querySelector(s);
const box = $('#mlGestures'), bubble = $('#gestBubble'), video = $('#gestVideo'), seen = $('#gestSeen');
const ICONS = { Swipe_Next: '👉', Swipe_Prev: '👈', Thumb_Up: '👍', Thumb_Down: '👎', Closed_Fist: '✊', ILoveYou: '🤟', F1: '☝️', F2: '✌️', F3: '3️⃣', F4: '4️⃣', F5: '🖐' };
// Пальцы — вкладки: 1 Задача · 2 Мои · 3 Смета · 4 Отчёты · 5 Настройки
const TABS = { F1: ['task', 'Задача'], F2: ['mine', 'Мои'], F3: ['smeta', 'Смета'], F4: ['reports', 'Отчёты'], F5: ['mail', 'Настройки'] };
const cheat = $('#gestCheat');
// Нажали на кружок — шпаргалка; ещё раз или мимо — спрятать
bubble.addEventListener('click', (e) => { e.stopPropagation(); cheat.hidden = !cheat.hidden; });
document.addEventListener('click', (e) => { if (!cheat.hidden && !cheat.contains(e.target)) cheat.hidden = true; });
$('#gestCheatOff').addEventListener('click', async () => {
  cheat.hidden = true;
  box.checked = false;
  await chrome.storage.local.set({ gestures: { on: false } });
  stop();
  toast('Жесты выключены, камера выключена');
});
let mod = null; // модуль с моделью грузим, только когда включили (он большой)

function confetti() {
  const layer = document.createElement('div');
  layer.className = 'confetti';
  for (let i = 0; i < 28; i++) {
    const s = document.createElement('span');
    s.textContent = ['🎉', '✨', '💜', '🌸', '⭐'][i % 5];
    s.style.left = `${Math.random() * 100}%`;
    s.style.animationDelay = `${Math.random() * 0.6}s`;
    s.style.fontSize = `${14 + Math.random() * 14}px`;
    layer.appendChild(s);
  }
  document.body.appendChild(layer);
  setTimeout(() => layer.remove(), 2600);
}

const visible = (sel) => { const el = document.querySelector(sel); return el && el.offsetParent !== null ? el : null; };

/**
 * Жест подержали — что сделать.
 * В окне «Вы уверены?» / «Отправить письмо?»: 👍 — да, 🖐 — отмена.
 * Иначе: 1–5 пальцев — вкладки, ✊ — скопировать тему, 👍 — «📨 Отправить» (спросит подтверждение — ещё 👍), 🤟 — конфетти, 👎 — кофе.
 */
function onGesture(name) {
  const dlg = document.querySelector('dialog[open]');
  bubble.classList.remove('pop'); void bubble.offsetWidth; bubble.classList.add('pop');
  if (dlg) {
    if (name === 'Thumb_Up') { const b = dlg.querySelector('[data-a="1"], .btn.primary'); if (b) b.click(); }
    if (name === 'F5') { const b = dlg.querySelector('[data-a="0"]'); if (b) b.click(); else dlg.close(); }
    return; // пока открыто окно — вкладки и остальное не трогаем
  }
  if (name === 'Swipe_Next' || name === 'Swipe_Prev') {
    // Смахнули ладонью — соседняя вкладка (по кругу)
    const tabs = [...document.querySelectorAll('.tabbar .tab')];
    const cur = tabs.findIndex((b) => b.getAttribute('aria-selected') === 'true');
    const next = tabs[(cur + (name === 'Swipe_Next' ? 1 : -1) + tabs.length) % tabs.length];
    next.click();
    seen.textContent = ICONS[name];
    toast(`${ICONS[name]} ${next.textContent.trim()}`);
    return;
  }
  if (TABS[name]) {
    const [tab, title] = TABS[name];
    const b = document.querySelector(`.tabbar .tab[data-tab="${tab}"]`);
    if (b && b.getAttribute('aria-selected') !== 'true') { b.click(); toast(`${ICONS[name]} ${title}`); }
    return;
  }
  if (name === 'Closed_Fist') {
    const b = visible('#taskCopyDone') || visible('#lCopySubject');
    if (b) { b.click(); toast('✊ Тема скопирована'); } else toast('✊ Тему копирую на экране «Добавлено» — сначала добавьте задачу', 'warn');
    return;
  }
  if (name === 'Thumb_Up') {
    const send = visible('#lSend');
    if (send && !send.disabled) { send.click(); toast('📨 Проверьте письмо и покажите 👍 ещё раз — отправлю'); } else toast('👍 Принято!');
    return;
  }
  if (name === 'ILoveYou') { playSound('harp'); confetti(); toast('🤟 Мир, дружба, локализация!'); return; }
  if (name === 'Thumb_Down') toast('👎 Понимаю. Может, кофе? ☕', 'warn');
}

async function start() {
  bubble.hidden = false;
  seen.textContent = '…';
  try {
    mod = mod || await import('./gestures.js');
    await mod.startGestures(video, { onGesture, onSeen: (n) => { seen.textContent = ICONS[n] || ''; } });
    toast('🖐 Жесты включены! Нажмите на кружок в углу — там шпаргалка');
  } catch (e) {
    bubble.hidden = true;
    box.checked = false;
    await chrome.storage.local.set({ gestures: { on: false, want: e.name === 'NotAllowedError' } });
    if (e.name === 'NotAllowedError') {
      chrome.tabs.create({ url: chrome.runtime.getURL('camera.html') });
      toast('Разрешите камеру в открывшейся вкладке — жесты включатся сами', 'warn');
    } else {
      toast(`Жесты не включились: ${e.name === 'NotFoundError' ? 'не вижу камеру' : e.message}`, 'err');
    }
  }
}
function stop() {
  if (mod) mod.stopGestures();
  bubble.hidden = true;
  cheat.hidden = true;
}

box.addEventListener('change', async () => {
  await chrome.storage.local.set({ gestures: { on: box.checked } });
  if (box.checked) start(); else { stop(); toast('Жесты выключены, камера выключена'); }
});
// Разрешили камеру во вкладке — включаем
chrome.storage.onChanged.addListener(async (ch) => {
  if (!ch.gesturesGranted) return;
  const { gestures } = await chrome.storage.local.get('gestures');
  if (gestures && gestures.want) { box.checked = true; await chrome.storage.local.set({ gestures: { on: true } }); start(); }
});
// Закрыли панель — камера выключается сама; открыли снова — включаем, если было включено
chrome.storage.local.get('gestures').then(({ gestures }) => { if (gestures && gestures.on) { box.checked = true; start(); } });
