// Передаёт найденный тип задачи (см. tracker-hook.js) в память расширения.
// Сообщения обрабатываем строго по очереди: два запроса подряд не должны затереть записи друг друга
let queue = Promise.resolve();
window.addEventListener('message', (e) => { queue = queue.then(() => handle(e)).catch(() => {}); });

async function handle(e) {
  if (e.source !== window || e.origin !== location.origin) return;
  const d = e.data;
  // Диагностика: последние 10 запросов «создать задачу», которые видели в трекере (без названий, описаний и кук)
  if (d && d.__bandTracker === 'seen') {
    const { trackerSeen = [] } = await chrome.storage.local.get('trackerSeen');
    const { __bandTracker, ...rest } = d;
    await chrome.storage.local.set({ trackerSeen: [{ ...rest, at: Date.now() }, ...trackerSeen].slice(0, 10) });
    return;
  }
  if (!d || d.__bandTracker !== 'board' || typeof d.path !== 'string' || typeof d.key !== 'string') return;
  const project = (/^([A-Z][A-Z0-9_]*)-\d+$/.exec(d.key) || [])[1] || '';
  const ws = (/^\/n\/([^/]+)/.exec(d.page || '') || [])[1] || '';
  const n = d.names && typeof d.names === 'object' ? d.names : {};
  const str = (v) => (typeof v === 'string' ? v.trim().slice(0, 60) : '');
  const slugs = d.path.split('/').slice(-3); // пространство / проект / тип
  const type = str(n.type) || slugs[2] || '';
  const projectName = str(n.project) || slugs[1] || '';
  const { boards = [] } = await chrome.storage.local.get('boards');
  const old = boards.find((b) => b.path === d.path);
  const board = {
    path: d.path,
    fields: d.fields && typeof d.fields === 'object' ? d.fields : {},
    ws: ws || (old && old.ws) || '',
    project: project || (old && old.project) || '',
    name: (old && old.name) || [type, projectName].filter(Boolean).join(' · '),
    space: str(n.space) || (old && old.space) || '',
    at: Date.now(),
  };
  await chrome.storage.local.set({ boards: [board, ...boards.filter((b) => b.path !== d.path)], boardsAdded: { path: d.path, at: Date.now() } });
}
