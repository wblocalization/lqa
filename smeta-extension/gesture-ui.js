// «Жесты» в панели: переключатель в «Настройках», кружок с камерой в углу и что делает каждый жест.
import { toast } from './core.js';
import { playSound } from './sounds.js';

const $ = (s) => document.querySelector(s);
const box = $('#mlGestures'), bubble = $('#gestBubble'), video = $('#gestVideo'), seen = $('#gestSeen');
const ICONS = { Thumb_Up: '👍', Open_Palm: '✋', Victory: '✌️', ILoveYou: '🤟', Thumb_Down: '👎', Closed_Fist: '✊', Pointing_Up: '☝️' };
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

/** Жест подержали — что сделать. В окне «Вы уверены?» 👍 = да, ✋ = отмена; иначе — просто весело. */
function onGesture(name) {
  const dlg = document.querySelector('dialog[open]');
  bubble.classList.remove('pop'); void bubble.offsetWidth; bubble.classList.add('pop');
  if (dlg && name === 'Thumb_Up') { const b = dlg.querySelector('[data-a="1"], .btn.primary'); if (b) b.click(); return; }
  if (dlg && name === 'Open_Palm') { const b = dlg.querySelector('[data-a="0"]'); if (b) b.click(); else dlg.close(); return; }
  if (name === 'Victory') { playSound('harp'); confetti(); toast('✌️ Мир, дружба, локализация!'); return; }
  if (name === 'ILoveYou') { playSound('chime'); toast('🤟 Лучшая команда локализации!'); return; }
  if (name === 'Thumb_Up') toast('👍 Принято!');
  if (name === 'Open_Palm') toast('✋ Привет!');
  if (name === 'Thumb_Down') toast('👎 Понимаю. Может, кофе?', 'warn');
}

async function start() {
  bubble.hidden = false;
  seen.textContent = '…';
  try {
    mod = mod || await import('./gestures.js');
    await mod.startGestures(video, { onGesture, onSeen: (n) => { seen.textContent = ICONS[n] || ''; } });
    toast('🖐 Жесты включены: 👍 — «Да», ✋ — «Отмена», ✌️ — сюрприз');
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
