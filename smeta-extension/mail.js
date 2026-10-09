// Вкладка «Письма»: своя почта для писем и включение рассылок (то же, что «⚙️ Настройки» в таблице).
import { settings, isConfigured, call, toast, ask, getLists } from './core.js';
import { renderLetterList } from './letter.js';

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
async function loadLetterList() {
  let contractors = [];
  try { if (isConfigured()) contractors = ((await getLists()) || {}).contractors || []; } catch { /* только настроенные */ }
  await renderLetterList($('#mlLetters'), contractors);
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
