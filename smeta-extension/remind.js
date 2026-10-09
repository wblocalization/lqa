// Напоминание о сроках — уведомлением Chrome (справа внизу экрана), в выбранное время (по умолчанию 10:00, по будням), раз в день.
// Нажали на уведомление — откроется панель на «Моих задачах». Включается и выключается во вкладке «Настройки».

export const REMIND_TIME = '10:00';

/** Сегодня уже пора? time — «ЧЧ:ММ». */
const reached = (now, time) => { const [h, m] = String(time || REMIND_TIME).split(':').map(Number); return now.getHours() * 60 + now.getMinutes() >= h * 60 + (m || 0); };

/** Когда следующее напоминание (для будильника в фоне): ближайшее «время» в подходящий день. */
export async function nextRemindAt() {
  const st = (await chrome.storage.local.get('remind')).remind || {};
  if (st.off) return null;
  const [h, m] = String(st.time || REMIND_TIME).split(':').map(Number);
  const d = new Date();
  d.setHours(h, m || 0, 5, 0);
  if (d <= new Date()) d.setDate(d.getDate() + 1);
  while (st.weekdays !== false && (d.getDay() === 0 || d.getDay() === 6)) d.setDate(d.getDate() + 1);
  return d.getTime();
}

/** Настройки звука: sound — для напоминаний (по умолчанию колокольчик), sendSound — «фьюх» при отправке письма. */
export async function soundPrefs() {
  const st = (await chrome.storage.local.get('remind')).remind || {};
  return { sound: st.sound || 'chime', sendSound: st.sendSound !== false };
}

/** Включены ли напоминания (по умолчанию — да). */
export async function remindOn() {
  const r = await chrome.storage.local.get('remind');
  return !(r.remind && r.remind.off);
}

const todayKey = () => new Date().toISOString().slice(0, 10);

/**
 * Показать уведомление по списку открытых задач (ответ myTasks). force — показать сразу (кнопка «Показать сейчас»),
 * иначе только по будням после 10:00 и не чаще раза в день.
 */
export async function remindDue(r, { force = false, play = null } = {}) {
  if (!r || !r.tasks) return false;
  const now = new Date();
  const st = (await chrome.storage.local.get('remind')).remind || {};
  if (!force) {
    if (st.off || st.shown === todayKey()) return false;
    if ((st.weekdays !== false && (now.getDay() === 0 || now.getDay() === 6)) || !reached(now, st.time)) return false;
  }
  const today = r.tasks.filter((t) => t.dueToday);
  const late = r.tasks.filter((t) => t.overdue);
  if (!force) await chrome.storage.local.set({ remind: { ...st, shown: todayKey() } }); // «Показать сейчас» не отменяет утреннее
  if (!today.length && !late.length && !force) return false;
  const sound = st.sound || 'chime';
  const parts = [];
  if (today.length) parts.push(`сдать сегодня: ${today.length}`);
  if (late.length) parts.push(`просрочено: ${late.length}`);
  const list = [...late, ...today].slice(0, 4).map((t) => `${t.id || '—'} · ${t.title}`).join('\n');
  chrome.notifications.create(`due-${Date.now()}`, {
    type: 'basic', iconUrl: 'icons/icon128.png', priority: 1,
    silent: sound !== 'off', // свой звук вместо системного
    title: parts.length ? `Сроки: ${parts.join(', ')}` : 'Сроки: на сегодня ничего не горит 🎉',
    message: list || 'Просроченных и сегодняшних задач нет',
  });
  if (play && sound !== 'off') play(sound);
  return true;
}
