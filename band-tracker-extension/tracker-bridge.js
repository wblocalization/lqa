// Передаёт найденную доску (см. tracker-hook.js) в память расширения.
window.addEventListener('message', async (e) => {
  if (e.source !== window || e.origin !== location.origin) return;
  const d = e.data;
  if (!d || d.__bandTracker !== 'board' || typeof d.path !== 'string' || typeof d.key !== 'string') return;
  const project = (/^([A-Z][A-Z0-9_]*)-\d+$/.exec(d.key) || [])[1] || '';
  const ws = (/^\/n\/([^/]+)/.exec(d.page || '') || [])[1] || '';
  const boardName = String(d.title || '').split(/\s+[-–—|]\s+/)[0].trim();
  const { boards = [] } = await chrome.storage.local.get('boards');
  const old = boards.find((b) => b.path === d.path);
  const board = {
    path: d.path,
    fields: d.fields && typeof d.fields === 'object' ? d.fields : {},
    ws: ws || (old && old.ws) || '',
    project: project || (old && old.project) || '',
    name: (old && old.name) || [boardName, project && `(${project})`].filter(Boolean).join(' ') || d.path.split('/').slice(-2).join(' / '),
    at: Date.now(),
  };
  await chrome.storage.local.set({ boards: [board, ...boards.filter((b) => b.path !== d.path)], boardsAdded: { path: d.path, at: Date.now() } });
});
