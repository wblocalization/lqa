// Невидимая страница только для звука напоминаний: фоновая часть расширения (service worker) сама играть звук не умеет.
import { playSound } from './sounds.js';

chrome.runtime.onMessage.addListener((m) => {
  if (m && m.type === 'play-sound') playSound(m.sound);
});
