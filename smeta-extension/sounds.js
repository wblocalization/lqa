// Звуки расширения — синтезируются на лету (Web Audio), без файлов.
// Напоминание о сроках — выбранный звук; отправили письмо — «фьюх».

export const SOUNDS = {
  chime: 'Колокольчик',
  harp: 'Арфа',
  bubbles: 'Пузырьки',
  coin: 'Монетка',
  off: 'Без звука',
};

let ctx = null;
const ac = () => (ctx = ctx || new AudioContext());

/** Одна нота: частота, начало (сек от «сейчас»), длительность, громкость, тембр. */
function note(freq, at, dur, vol = 0.25, type = 'sine', slideTo = 0) {
  const a = ac();
  const t = a.currentTime + at;
  const o = a.createOscillator();
  const g = a.createGain();
  o.type = type;
  o.frequency.setValueAtTime(freq, t);
  if (slideTo) o.frequency.exponentialRampToValueAtTime(slideTo, t + dur);
  g.gain.setValueAtTime(0.0001, t);
  g.gain.exponentialRampToValueAtTime(vol, t + 0.015);
  g.gain.exponentialRampToValueAtTime(0.0001, t + dur);
  o.connect(g).connect(a.destination);
  o.start(t);
  o.stop(t + dur + 0.05);
}

const PLAY = {
  // динь-дилинь: три ноты вверх с обертоном
  chime() { [[880, 0], [1175, 0.12], [1568, 0.24]].forEach(([f, at]) => { note(f, at, 0.9, 0.22); note(f * 2, at, 0.5, 0.05); }); },
  // арпеджио, как перебор струн
  harp() { [523, 659, 784, 1047, 1319].forEach((f, i) => note(f, i * 0.07, 1.2, 0.18, 'triangle')); },
  // буль-буль
  bubbles() { [0, 0.11, 0.2].forEach((at, i) => note(300 + i * 120, at, 0.16, 0.3, 'sine', 900 + i * 250)); },
  // как монетка в игре
  coin() { note(988, 0, 0.09, 0.18, 'square'); note(1319, 0.08, 0.45, 0.16, 'square'); },
  // фьюх — письмо улетело
  whoosh() {
    const a = ac();
    const t = a.currentTime;
    const len = 0.45;
    const buf = a.createBuffer(1, Math.floor(a.sampleRate * len), a.sampleRate);
    const d = buf.getChannelData(0);
    for (let i = 0; i < d.length; i++) d[i] = Math.random() * 2 - 1;
    const src = a.createBufferSource();
    src.buffer = buf;
    const f = a.createBiquadFilter();
    f.type = 'bandpass'; f.Q.value = 1.2;
    f.frequency.setValueAtTime(400, t);
    f.frequency.exponentialRampToValueAtTime(4000, t + len);
    const g = a.createGain();
    g.gain.setValueAtTime(0.0001, t);
    g.gain.exponentialRampToValueAtTime(0.35, t + 0.12);
    g.gain.exponentialRampToValueAtTime(0.0001, t + len);
    src.connect(f).connect(g).connect(a.destination);
    src.start(t);
    note(1568, 0.3, 0.35, 0.08); // и тихое «динь» в конце
  },
};

/** Сыграть звук по имени (off или неизвестное — тишина). */
export async function playSound(name) {
  if (!PLAY[name]) return;
  try {
    if (ac().state === 'suspended') await ac().resume();
    PLAY[name]();
  } catch { /* браузер не дал звук — не страшно */ }
}
