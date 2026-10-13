// Звуки расширения — синтезируются на лету (Web Audio), без файлов.
// Напоминание о сроках — выбранный звук; отправили письмо — «фьюх».

export const SOUNDS = {
  chime: 'Колокольчик',
  harp: 'Арфа',
  bubbles: 'Пузырьки',
  coin: 'Монетка',
  whoosh: 'Фьюх',
  duck: '🦆 Кря-кря',
  squeak: '🐤 Пищалка-уточка',
  meow: '🐱 Мяу',
  cuckoo: '🐦 Ку-ку',
  boing: '🪀 Пружинка',
  laser: '🔫 Пиу-пиу',
  drum: '🥁 Ба-дум-тсс',
  horn: '🤡 Клаксон',
  trombone: '🎺 Грустный тромбон',
  off: 'Без звука',
};
/** Для выбора: спокойные и смешные отдельно. */
export const SOUND_GROUPS = [
  ['Спокойные', ['chime', 'harp', 'bubbles', 'coin', 'whoosh']],
  ['Смешные', ['duck', 'squeak', 'meow', 'cuckoo', 'boing', 'laser', 'drum', 'horn', 'trombone']],
  ['', ['off']],
];
export function soundOptions() {
  return SOUND_GROUPS.map(([g, keys]) => {
    const opts = keys.map((k) => `<option value="${k}">${SOUNDS[k]}</option>`).join('');
    return g ? `<optgroup label="${g}">${opts}</optgroup>` : opts;
  }).join('');
}

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

/**
 * Голос с «формантой»: генератор (пила/квадрат) через полосовой фильтр — так получаются кря, мяу, клаксон.
 * pitch: [[время, частота], …] от начала звука; formant: [[время, частота], …].
 */
function voice({ at = 0, dur, type = 'sawtooth', pitch, formant, q = 4, vol = 0.3, vibrato = 0 }) {
  const a = ac();
  const t = a.currentTime + at;
  const o = a.createOscillator();
  o.type = type;
  o.frequency.setValueAtTime(pitch[0][1], t);
  pitch.slice(1).forEach(([dt, f]) => o.frequency.exponentialRampToValueAtTime(f, t + dt));
  if (vibrato) {
    const lfo = a.createOscillator(), lg = a.createGain();
    lfo.frequency.value = 6; lg.gain.value = vibrato;
    lfo.connect(lg).connect(o.frequency); lfo.start(t); lfo.stop(t + dur + 0.05);
  }
  const f = a.createBiquadFilter();
  f.type = 'bandpass'; f.Q.value = q;
  f.frequency.setValueAtTime(formant[0][1], t);
  formant.slice(1).forEach(([dt, fr]) => f.frequency.linearRampToValueAtTime(fr, t + dt));
  const g = a.createGain();
  g.gain.setValueAtTime(0.0001, t);
  g.gain.exponentialRampToValueAtTime(vol, t + 0.02);
  g.gain.setValueAtTime(vol, t + dur * 0.7);
  g.gain.exponentialRampToValueAtTime(0.0001, t + dur);
  o.connect(f).connect(g).connect(a.destination);
  o.start(t); o.stop(t + dur + 0.05);
}

/** Шум (тарелка, малый барабан). */
function noise({ at = 0, dur, filter = 'highpass', freq = 5000, vol = 0.2 }) {
  const a = ac();
  const t = a.currentTime + at;
  const buf = a.createBuffer(1, Math.floor(a.sampleRate * dur), a.sampleRate);
  const d = buf.getChannelData(0);
  for (let i = 0; i < d.length; i++) d[i] = Math.random() * 2 - 1;
  const src = a.createBufferSource(); src.buffer = buf;
  const f = a.createBiquadFilter(); f.type = filter; f.frequency.value = freq;
  const g = a.createGain();
  g.gain.setValueAtTime(vol, t);
  g.gain.exponentialRampToValueAtTime(0.0001, t + dur);
  src.connect(f).connect(g).connect(a.destination);
  src.start(t);
}

const PLAY = {
  // 🦆 кря-кря: «носовой» звук с падающим тоном, два раза
  duck() {
    [0, 0.24].forEach((at) => {
      voice({ at, dur: 0.17, pitch: [[0, 560], [0.17, 330]], formant: [[0, 1300], [0.17, 900]], q: 3, vol: 0.45 });
      voice({ at, dur: 0.17, pitch: [[0, 566], [0.17, 334]], formant: [[0, 2400], [0.17, 1800]], q: 5, vol: 0.18 });
    });
  },
  // 🐤 пищалка-уточка из ванной
  squeak() { note(1900, 0, 0.12, 0.2, 'sine', 2600); note(2600, 0.12, 0.14, 0.18, 'sine', 1700); },
  // 🐱 мяу: тон вверх-вниз, форманта «и-а-у»
  meow() { voice({ dur: 0.75, pitch: [[0, 480], [0.25, 720], [0.75, 420]], formant: [[0, 900], [0.25, 1800], [0.75, 700]], q: 2.5, vol: 0.5, vibrato: 8 }); },
  // 🐦 ку-ку, ку-ку
  cuckoo() { [[659, 0], [523, 0.22], [659, 0.6], [523, 0.82]].forEach(([f, at]) => note(f, at, 0.26, 0.28, 'sine')); },
  // 🪀 пружинка: тон вверх с дрожанием
  boing() { voice({ type: 'sine', dur: 0.6, pitch: [[0, 110], [0.5, 520]], formant: [[0, 600], [0.6, 600]], q: 0.3, vol: 0.5, vibrato: 30 }); },
  // 🔫 пиу-пиу
  laser() { [0, 0.2].forEach((at) => note(1700, at, 0.17, 0.12, 'square', 180)); },
  // 🥁 ба-дум-тсс
  drum() {
    note(190, 0, 0.16, 0.6, 'sine', 90);
    note(150, 0.2, 0.22, 0.6, 'sine', 70);
    noise({ at: 0.2, dur: 0.12, filter: 'bandpass', freq: 1800, vol: 0.15 });
    noise({ at: 0.46, dur: 0.9, filter: 'highpass', freq: 7000, vol: 0.18 });
  },
  // 🤡 клаксон: бип-бип
  horn() {
    [0, 0.2].forEach((at) => {
      voice({ at, dur: 0.15, type: 'sawtooth', pitch: [[0, 440], [0.15, 430]], formant: [[0, 1200], [0.15, 1200]], q: 1.2, vol: 0.35 });
      voice({ at, dur: 0.15, type: 'sawtooth', pitch: [[0, 555], [0.15, 545]], formant: [[0, 1600], [0.15, 1600]], q: 1.2, vol: 0.25 });
    });
  },
  // 🎺 ва-ва-ва-вааа (грустный тромбон)
  trombone() {
    [[311, 0, 0.32], [294, 0.36, 0.32], [277, 0.72, 0.32], [262, 1.08, 0.9]].forEach(([f, at, dur], i) =>
      voice({ at, dur, pitch: [[0, f], [dur, f * 0.985]], formant: [[0, 700], [dur, 600]], q: 1.5, vol: 0.4, vibrato: i === 3 ? 6 : 0 }));
  },
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
