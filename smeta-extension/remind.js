// Напоминание о сроках — уведомлением Chrome (справа внизу экрана), по будням около 10:00 и один раз в день.
// Нажали на уведомление — откроется панель на «Моих задачах». Включается и выключается во вкладке «Почта».

export const REMIND_HOUR = 10;

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
export async function remindDue(r, { force = false } = {}) {
  if (!r || !r.tasks) return false;
  const now = new Date();
  const st = (await chrome.storage.local.get('remind')).remind || {};
  if (!force) {
    if (st.off || st.shown === todayKey()) return false;
    if (now.getDay() === 0 || now.getDay() === 6 || now.getHours() < REMIND_HOUR) return false;
  }
  const today = r.tasks.filter((t) => t.dueToday);
  const late = r.tasks.filter((t) => t.overdue);
  if (!force) await chrome.storage.local.set({ remind: { ...st, shown: todayKey() } }); // «Показать сейчас» не отменяет утреннее
  if (!today.length && !late.length && !force) return false;
  const parts = [];
  if (today.length) parts.push(`сдать сегодня: ${today.length}`);
  if (late.length) parts.push(`просрочено: ${late.length}`);
  const list = [...late, ...today].slice(0, 4).map((t) => `${t.id || '—'} · ${t.title}`).join('\n');
  chrome.notifications.create(`due-${Date.now()}`, {
    type: 'basic', iconUrl: 'icons/icon128.png', priority: 1,
    title: parts.length ? `Сроки: ${parts.join(', ')}` : 'Сроки: на сегодня ничего не горит 🎉',
    message: list || 'Просроченных и сегодняшних задач нет',
  });
  return true;
}
