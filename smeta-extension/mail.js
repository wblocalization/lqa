// Вкладка «Настройки»: своя почта для писем и включение рассылок (то же, что «⚙️ Настройки» в таблице).
import { settings, isConfigured, call, toast, ask, getLists } from './core.js';
import { renderLetterList, renderSignatureCard } from './letter.js';
import { remindOn, remindDue, soundPrefs } from './remind.js';
import { SOUNDS, playSound } from './sounds.js';

const $ = (s) => document.querySelector(s);
const els = {
  email: $('#mlEmail'), save: $('#mlSave'), who: $('#mlWho'), list: $('#mlList'), msg: $('#mlMsg'),
};

function setMsg(text, kind = 'info') {
  els.msg.textContent = text;
  els.msg.className = `status ${kind}`;
}

function render(st) {
  els.email.value = st.email || '';
  els.list.querySelectorAll('input[data-kind]').forEach((cb) => { cb.checked = !!st[cb.dataset.kind]; cb.disabled = !!st.noTriggers; });
  // Нет разрешения на расписания — переключатели неактивны, почта и «прислать мне» работают
  if (st.noTriggers) {
    setMsg('Включать письма пока нельзя: скрипту таблицы нужно разрешение Google на расписания. ' +
      `Откройте таблицу${st.runAs ? ` под аккаунтом ${st.runAs}` : ' под аккаунтом, от которого развёрнуто веб-приложение'} → ` +
      '«⚙️ Настройки → 🔑 Разрешить письма и Excel» → «Разрешить» → затем новая версия развёртывания. ' +
      `Почта и «Прислать мне сейчас» работают и так.\n\nОтвет Google: ${st.noTriggers}`, 'err');
  }
}

/** Письма подрядчикам — список и редактор; подрядчики из справочника таблицы (не загрузился — только настроенные). */
// Напоминания в браузере — свои у каждого
const remindBox = $('#mlRemind');
remindOn().then((on) => { remindBox.checked = on; });
remindBox.addEventListener('change', async () => {
  const st = (await chrome.storage.local.get('remind')).remind || {};
  await chrome.storage.local.set({ remind: { ...st, off: !remindBox.checked } });
  toast(remindBox.checked ? `Напоминания включены — в ${timeIn.value || '10:00'}` : 'Напоминания выключены');
});
$('#mlRemindTest').addEventListener('click', async (e) => {
  if (!isConfigured() || !settings.manager) return toast('Выберите себя в настройках (⚙️ → «Кто вы»)', 'err');
  e.target.disabled = true;
  try {
    const r = await call({ action: 'myTasks', manager: settings.manager, filter: 'open' });
    await remindDue(r, { force: true, play: playSound });
    toast('Уведомление показано — оно в углу экрана');
  } catch (err) {
    toast(`Не получилось: ${err.message}`, 'err');
  } finally {
    e.target.disabled = false;
  }
});

// Звуки: для напоминаний — на выбор; «фьюх» при отправке письма — вкл/выкл
const soundSel = $('#mlSound');
soundSel.innerHTML = Object.entries(SOUNDS).map(([k, v]) => `<option value="${k}">${v}</option>`).join('');
const sendSoundBox = $('#mlSendSound');
soundPrefs().then((p) => { soundSel.value = p.sound; sendSoundBox.checked = p.sendSound; });
// Время напоминания и дни — свои у каждого
const timeIn = $('#mlRemindTime'), weekdaysBox = $('#mlRemindWeekdays');
chrome.storage.local.get('remind').then(({ remind: st = {} }) => { timeIn.value = st.time || '10:00'; weekdaysBox.checked = st.weekdays !== false; });
timeIn.addEventListener('change', async () => {
  if (!timeIn.value) return;
  // новое время сегодня ещё впереди — пусть сегодня тоже напомнит
  const st = (await chrome.storage.local.get('remind')).remind || {};
  await chrome.storage.local.set({ remind: { ...st, time: timeIn.value, shown: '' } });
  toast(`Напомню в ${timeIn.value}${weekdaysBox.checked ? ' по будням' : ' каждый день'}`);
});
weekdaysBox.addEventListener('change', () => saveSound({ weekdays: weekdaysBox.checked }));
async function saveSound(patch) {
  const st = (await chrome.storage.local.get('remind')).remind || {};
  await chrome.storage.local.set({ remind: { ...st, ...patch } });
}
soundSel.addEventListener('change', () => { saveSound({ sound: soundSel.value }); playSound(soundSel.value); });
$('#mlSoundPlay').addEventListener('click', () => playSound(soundSel.value));
sendSoundBox.addEventListener('change', () => { saveSound({ sendSound: sendSoundBox.checked }); if (sendSoundBox.checked) playSound('whoosh'); });

async function loadLetterList() {
  let contractors = [];
  try { if (isConfigured()) contractors = ((await getLists()) || {}).contractors || []; } catch { /* только настроенные */ }
  await renderLetterList($('#mlLetters'), contractors);
  renderSignatureCard();
}

export async function loadMail() {
  loadLetterList();
  if (!isConfigured() || !settings.manager) {
    els.who.textContent = '';
    return setMsg('Выберите себя в настройках (⚙️ → «Кто вы»)', 'err');
  }
  els.who.textContent = settings.manager;
  setMsg('Загружаю…');
  try {
    const st = await call({ action: 'mailStatus', manager: settings.manager });
    setMsg('');
    render(st);
  } catch (e) {
    setMsg(`Не загрузилось: ${e.message}`, 'err');
  }
}

els.save.addEventListener('click', async () => {
  els.save.disabled = true;
  try {
    render(await call({ action: 'mailEmail', manager: settings.manager, email: els.email.value.trim() }));
    toast(els.email.value ? '✓ Почта сохранена' : 'Почта убрана — письма приходить не будут');
    setMsg('');
  } catch (e) {
    setMsg(e.message, 'err');
  } finally {
    els.save.disabled = false;
  }
});

els.list.addEventListener('change', async (e) => {
  const cb = e.target.closest('input[data-kind]');
  if (!cb) return;
  const on = cb.checked;
  if (!await ask(on ? { title: 'Включить для всей команды?', text: 'Письмо получат все, у кого вписана почта.', ok: 'Включить' } : { title: 'Выключить для всей команды?', ok: 'Выключить', danger: true, icon: '!' })) {
    cb.checked = !on;
    return;
  }
  cb.disabled = true;
  try {
    render(await call({ action: 'mailToggle', kind: cb.dataset.kind, on, manager: settings.manager }));
    toast(on ? '✓ Включено' : 'Выключено');
  } catch (err) {
    cb.checked = !on;
    setMsg(`Не получилось: ${err.message}`, 'err');
  } finally {
    cb.disabled = false;
  }
});

els.list.addEventListener('click', async (e) => {
  const b = e.target.closest('button[data-test]');
  if (!b) return;
  b.disabled = true;
  setMsg('Отправляю…');
  try {
    const r = await call({ action: 'mailTest', kind: b.dataset.test, manager: settings.manager });
    setMsg(r.sent ? `Отправлено на ${r.email} — проверьте почту` : 'Письма нет: по вашим задачам сейчас писать нечего', r.sent ? 'ok' : 'info');
  } catch (err) {
    setMsg(err.message, 'err');
  } finally {
    b.disabled = false;
  }
});
