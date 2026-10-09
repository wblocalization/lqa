// 🖐 Жесты (бета, для веселья): камера ноутбука + модель Google MediaPipe прямо в браузере — видео никуда не уходит.
// Что делает каждый жест — в gesture-ui.js. Жест надо подержать ~0.7 сек, чтобы случайно не сработало.
// Включается в «Настройки» → «Жесты»; камера работает, только пока панель открыта и переключатель включён.
import { FilesetResolver, GestureRecognizer } from './vendor/mediapipe/vision_bundle.mjs';

const HOLD_MS = 700;      // сколько держать жест
const COOLDOWN_MS = 1500; // после срабатывания — пауза
const MIN_SCORE = 0.6;

let recognizer = null;
let stream = null;
let video = null;
let timer = null;
let held = { name: '', since: 0 };
let lastFire = 0;
let fired = '';

/** Загрузить модель (один раз, ~2–3 сек). */
async function loadRecognizer() {
  if (recognizer) return recognizer;
  const files = await FilesetResolver.forVisionTasks(chrome.runtime.getURL('vendor/mediapipe'));
  recognizer = await GestureRecognizer.createFromOptions(files, {
    baseOptions: { modelAssetPath: chrome.runtime.getURL('vendor/mediapipe/gesture_recognizer.task'), delegate: 'CPU' },
    runningMode: 'VIDEO', numHands: 2,
  });
  return recognizer;
}

// Жесты с большим пальцем и кулак — как их понимает модель; остальное — по числу поднятых пальцев (F1…F5).
const MODEL_GESTURES = ['Thumb_Up', 'Thumb_Down', 'Closed_Fist', 'ILoveYou'];
function handName(r, i) {
  const g = r.gestures && r.gestures[i] && r.gestures[i][0];
  const cat = g && g.score >= MIN_SCORE ? g.categoryName : '';
  if (MODEL_GESTURES.includes(cat)) return cat;
  const lm = r.landmarks && r.landmarks[i];
  if (!lm) return '';
  const n = countFingers(lm);
  return n >= 1 ? `F${n}` : '';
}
// Двумя руками сразу — праздничные жесты: ✌️✌️, 👍👍, 🙌
const DUO = { F2: 'Duo_Victory', Thumb_Up: 'Duo_Thumbs', F5: 'Duo_Palms' };
function gestureName(r) {
  const a = handName(r, 0);
  if (r.landmarks && r.landmarks.length > 1) {
    const b = handName(r, 1);
    if (a && a === b && DUO[a]) return DUO[a];
  }
  return a;
}

/** Сколько пальцев поднято — по 21 точке руки (не зависит от того, как повёрнута рука). */
export function countFingers(lm) {
  const d = (a, b) => Math.hypot(lm[a].x - lm[b].x, lm[a].y - lm[b].y);
  let n = 0;
  [[8, 6], [12, 10], [16, 14], [20, 18]].forEach(([tip, pip]) => { if (d(0, tip) > d(0, pip) * 1.12) n++; });
  // большой палец отставлен: кончик дальше от основания мизинца, чем его собственный сустав (подобрано по фото рук)
  if (d(4, 17) > d(2, 17) * 1.02) n++;
  return n;
}

/**
 * Включить. onGesture(name) — жест подержали: 'Thumb_Up' | 'Thumb_Down' | 'Closed_Fist' | 'ILoveYou' | 'F1'…'F5' (сколько пальцев)
 * или двумя руками: 'Duo_Victory' (✌️✌️) | 'Duo_Thumbs' (👍👍) | 'Duo_Palms' (🙌).
 * onSeen(name) — что видно прямо сейчас (для подсказки на экране). Бросает ошибку 'NotAllowedError', если нет доступа к камере.
 */
export async function startGestures(videoEl, { onGesture, onSeen }) {
  stopGestures();
  video = videoEl;
  stream = await navigator.mediaDevices.getUserMedia({ video: { width: 320, height: 240, facingMode: 'user' }, audio: false });
  video.srcObject = stream;
  await video.play();
  await loadRecognizer();
  let lastT = -1;
  timer = setInterval(() => {
    if (!video || video.readyState < 2 || document.hidden) return;
    const t = performance.now();
    if (t <= lastT) return;
    lastT = t;
    let name = '';
    try {
      const r = recognizer.recognizeForVideo(video, t);
      name = gestureName(r);
    } catch { return; }
    if (onSeen) onSeen(name);
    if (name !== held.name) { held = { name, since: t }; if (name !== fired) fired = ''; return; }
    if (!name || fired === name || t - held.since < HOLD_MS || t - lastFire < COOLDOWN_MS) return;
    fired = name; // тот же жест повторно — только если опустить руку и показать снова
    lastFire = t;
    onGesture(name);
  }, 100);
}

export function stopGestures() {
  clearInterval(timer);
  timer = null;
  if (stream) stream.getTracks().forEach((tr) => tr.stop());
  stream = null;
  if (video) video.srcObject = null;
  held = { name: '', since: 0 };
  fired = '';
}

export const gesturesRunning = () => Boolean(stream);
