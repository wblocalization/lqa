// 🖐 Жесты (бета, для веселья): камера ноутбука + модель Google MediaPipe прямо в браузере — видео никуда не уходит.
// 👍 — «Да» в окне подтверждения, ✋ — «Отмена», ✌️ — сюрприз. Жест надо подержать ~0.7 сек, чтобы случайно не сработало.
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
    runningMode: 'VIDEO', numHands: 1,
  });
  return recognizer;
}

/**
 * Включить. onGesture(name) — жест подержали: 'Thumb_Up' | 'Open_Palm' | 'Victory' | 'Pointing_Up' | 'Closed_Fist' | 'ILoveYou' | 'Thumb_Down'.
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
      const g = r.gestures && r.gestures[0] && r.gestures[0][0];
      if (g && g.score >= MIN_SCORE && g.categoryName !== 'None') name = g.categoryName;
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
