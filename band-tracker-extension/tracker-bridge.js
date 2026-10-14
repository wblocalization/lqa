// Передаёт найденный тип задачи (см. tracker-hook.js) в память расширения.
const SEED_OPTIONS = {
  'c907e2d8-c109-4147-ae67-f78fbdc3bdfb': { беларусь: { id: 'PQYSAXgqcl3w8Gxcoas1s', label: 'Беларусь' }, грузия: { id: '_3TyyGB7ox9s9H5-KNumf', label: 'Грузия' } },
  '883735ab-be95-4d44-a830-62740ae23fcb': { межнар: { id: 'kp9u0p5qrqk60pkUJfeoG', label: 'Межнар' } },
};

// Сообщения обрабатываем строго по очереди: два запроса подряд не должны затереть записи друг друга
let queue = Promise.resolve();
window.addEventListener('message', (e) => { queue = queue.then(() => handle(e)).catch(() => {}); });

async function handle(e) {
  if (e.source !== window || e.origin !== location.origin) return;
  const d = e.data;
  // Варианты полей (БЗ, страны…): из созданных руками задач и из списков, которые трекер загружает.
  if (d && d.__bandTracker === 'options' && Array.isArray(d.lists)) {
    const { fieldOptions: saved = {} } = await chrome.storage.local.get('fieldOptions');
    // стартовые варианты (как в tracker.js) — по ним узнаём списки и у тех, кто ещё не создавал задачу руками
    const fieldOptions = JSON.parse(JSON.stringify(SEED_OPTIONS));
    for (const [fid, o] of Object.entries(saved)) fieldOptions[fid] = { ...(fieldOptions[fid] || {}), ...o };
    let changed = false;
    for (const list of d.lists) {
      const ids = new Set(list.map((x) => x.id));
      for (const fid of Object.keys(fieldOptions)) {
        if (!Object.values(fieldOptions[fid]).some((o) => ids.has(o.id))) continue; // этот список — варианты этого поля
        for (const x of list) { const k = x.label.trim().toLowerCase().replace(/ё/g, 'е'); if (!fieldOptions[fid][k]) { fieldOptions[fid][k] = x; changed = true; } }
      }
    }
    if (changed) await chrome.storage.local.set({ fieldOptions });
    return;
  }
  // Диагностика: последние 10 запросов «создать задачу», которые видели в трекере (без названий, описаний и кук)
  if (d && d.__bandTracker === 'seen') {
    // варианты, выбранные в этой задаче, — сразу в словарь полей
    const { fieldOptions = {} } = await chrome.storage.local.get('fieldOptions');
    for (const [fid, vals] of Object.entries(d.fields || {})) {
      if (!Array.isArray(vals)) continue;
      for (const x of vals) {
        if (!x || typeof x !== 'object' || typeof x.id !== 'string' || typeof x.label !== 'string') continue;
        (fieldOptions[fid] = fieldOptions[fid] || {})[x.label.trim().toLowerCase().replace(/ё/g, 'е')] = { id: x.id, label: x.label };
      }
    }
    await chrome.storage.local.set({ fieldOptions });
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
