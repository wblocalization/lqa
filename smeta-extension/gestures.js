// 🖐 Жесты (бета, для веселья): камера ноутбука + модель Google MediaPipe прямо в браузере — видео никуда не уходит.
// Что делает каждый жест — в gesture-ui.js. Жест надо подержать ~0.7 сек, чтобы случайно не сработало.
// Включается в «Настройки» → «Жесты»; камера работает, только пока панель открыта и переключатель включён.
import { FilesetResolver, GestureRecognizer } from './vendor/mediapipe/vision_bundle.mjs';

const HOLD_MS = 700;      // сколько держать жест
const COOLDOWN_MS = 1500; // после срабатывания — пауза
const MIN_SCORE = 0.6;
// Смахивание: ладонь (3+ пальца) быстро прошла в сторону ≥ четверти кадра за полсекунды
const SWIPE_DX = 0.22, SWIPE_MAX_DY = 0.18, SWIPE_WINDOW_MS = 550, SWIPE_PAUSE_MS = 900;
const STILL = 0.07; // для «подержать жест» рука должна стоять почти на месте

let recognizer = null;
let stream = null;
let video = null;
let timer = null;
let held = { name: '', since: 0 };
let lastFire = 0;
let fired = '';
let track = [];     // где была ладонь последние полсекунды: { t, x, y }
let lastSwipe = 0;

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

// Жесты с большим пальцем и кулак — как их понимает модель; остальное — по числу поднятых пальцев (F1…F5).
const MODEL_GESTURES = ['Thumb_Up', 'Thumb_Down', 'Closed_Fist', 'ILoveYou'];
function gestureName(r) {
  const g = r.gestures && r.gestures[0] && r.gestures[0][0];
  const cat = g && g.score >= MIN_SCORE ? g.categoryName : '';
  if (MODEL_GESTURES.includes(cat)) return cat;
  const lm = r.landmarks && r.landmarks[0];
  if (!lm) return '';
  const n = countFingers(lm);
  return n >= 1 ? `F${n}` : '';
}
/** Середина ладони (основание среднего пальца) — по ней следим за движением. */
const palmAt = (r) => { const lm = r.landmarks && r.landmarks[0]; return lm ? { x: lm[9].x, y: lm[9].y } : null; };

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
 * Включить. onGesture(name) — жест подержали: 'Thumb_Up' | 'Thumb_Down' | 'Closed_Fist' | 'ILoveYou' | 'F1'…'F5' (сколько пальцев).
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
    let name = '', palm = null;
    try {
      const r = recognizer.recognizeForVideo(video, t);
      name = gestureName(r);
      palm = palmAt(r);
    } catch { return; }
    if (onSeen) onSeen(name);

    // Смахивание открытой ладонью. В кадре камеры x зеркален: рука влево (от себя) — x растёт → «дальше»
    if (palm && /^F[345]$/.test(name)) {
      track.push({ t, ...palm });
      track = track.filter((p) => t - p.t <= SWIPE_WINDOW_MS);
      const first = track[0];
      const dx = palm.x - first.x, dy = palm.y - first.y;
      if (Math.abs(dx) >= SWIPE_DX && Math.abs(dy) <= SWIPE_MAX_DY && t - lastSwipe > SWIPE_PAUSE_MS) {
        lastSwipe = t;
        track = [];
        held = { name, since: t, at: palm }; // после смахивания не считаем, что ладонь «держат»
        fired = name;
        onGesture(dx > 0 ? 'Swipe_Next' : 'Swipe_Prev');
        return;
      }
    } else track = [];

    if (name !== held.name) { held = { name, since: t, at: palm }; if (name !== fired) fired = ''; return; }
    // «Подержать» — только если рука стоит на месте (иначе это смахивание, а не жест)
    if (palm && held.at && Math.hypot(palm.x - held.at.x, palm.y - held.at.y) > STILL) { held = { name, since: t, at: palm }; return; }
    if (!name || fired === name || t - held.since < HOLD_MS || t - lastFire < COOLDOWN_MS || t - lastSwipe < SWIPE_PAUSE_MS) return;
    fired = name; // тот же жест повторно — только если опустить руку и показать снова
    lastFire = t;
    onGesture(name);
  }, 70);
}

export function stopGestures() {
  clearInterval(timer);
  timer = null;
  if (stream) stream.getTracks().forEach((tr) => tr.stop());
  stream = null;
  if (video) video.srcObject = null;
  held = { name: '', since: 0 };
  fired = '';
  track = [];
}

export const gesturesRunning = () => Boolean(stream);
