// Загрузка сметы на ВБ Диск (Nextcloud): папка с темой письма → PDF в неё → ссылка вида disk.wb.ru/f/<номер>.
// Nextcloud пускает такие запросы только со своей страницы (строгая cookie), поэтому всё делаем
// во вкладке disk.wb.ru — от имени того, кто в ней вошёл. Нет открытой вкладки — откроем в фоне.

export const DISK_ORIGIN = 'https://disk.wb.ru';

/** Имя папки из темы письма: без символов, которые нельзя в именах, и не слишком длинное. */
export function folderNameFromSubject(subject) {
  return String(subject || '')
    .replace(/[\\/:*?"<>|#%\n\r\t]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 150)
    .replace(/[.\s]+$/, '');
}

function waitLoaded(tabId, timeoutMs = 30000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(onUpdated);
      reject(new Error('ВБ Диск не открылся за 30 секунд'));
    }, timeoutMs);
    function onUpdated(id, info) {
      if (id === tabId && info.status === 'complete') {
        clearTimeout(timer);
        chrome.tabs.onUpdated.removeListener(onUpdated);
        resolve();
      }
    }
    chrome.tabs.onUpdated.addListener(onUpdated);
  });
}

async function diskTab() {
  const [tab] = await chrome.tabs.query({ url: `${DISK_ORIGIN}/*` });
  if (tab && tab.status === 'complete') return tab;
  if (tab) { await waitLoaded(tab.id); return tab; }
  const created = await chrome.tabs.create({ url: `${DISK_ORIGIN}/apps/files/`, active: false });
  await waitLoaded(created.id);
  return created;
}

function toBase64(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

/** Выполнить действие на странице ВБ Диска (там вход и requesttoken Nextcloud). */
async function inDiskPage(args) {
  const tab = await diskTab();
  let res;
  try {
    [res] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, world: 'MAIN', func: diskInPage, args });
  } catch (e) {
    if (/error page|Cannot access/i.test(e.message)) {
      throw new Error('ВБ Диск не открылся во вкладке — проверьте сеть/VPN и что disk.wb.ru открывается');
    }
    throw e;
  }
  const r = res && res.result;
  if (!r) throw new Error('ВБ Диск не ответил');
  if (r.error) throw new Error(r.error);
  return r;
}

/** Какие PDF уже лежат в папке (папки нет — пустой список). */
export async function listDiskPdfs(folders) {
  return (await inDiskPage(['list', folders, '', '', [], false])).files;
}

/**
 * file — PDF (File), folders — путь папок, например ['localization', 'Сметы', '[LIT-26-2223] Перевод…'].
 * remove — старые сметы в этой папке, которые убрать (уйдут в корзину ВБ Диска).
 * keepBoth — не затирать файл с тем же именем, а сохранить новый как «имя (2).pdf».
 * Возвращает { link, path }.
 */
export async function uploadToDisk(file, folders, { remove = [], keepBoth = false } = {}) {
  const b64 = toBase64(new Uint8Array(await file.arrayBuffer()));
  return inDiskPage(['upload', folders, file.name, b64, remove, keepBoth]);
}

/** Выполняется на странице disk.wb.ru. mode: 'list' — какие PDF в папке; 'upload' — положить PDF. */
async function diskInPage(mode, folders, fileName, b64, remove, keepBoth) {
  const OC = window.OC;
  const token = (OC && OC.requestToken) || document.head.dataset.requesttoken;
  const user = (OC && OC.getCurrentUser && OC.getCurrentUser() && OC.getCurrentUser().uid) || document.head.dataset.user;
  if (!token || !user) return { error: 'Не вижу входа на ВБ Диск — откройте disk.wb.ru, войдите и попробуйте ещё раз' };
  const enc = encodeURIComponent;
  const base = `/remote.php/dav/files/${enc(user)}`;
  const headers = { requesttoken: token };
  const dir = base + folders.map((f) => `/${enc(f)}`).join('');
  async function listPdfs() {
    const r = await fetch(`${dir}/`, { method: 'PROPFIND', headers: { ...headers, Depth: '1' }, credentials: 'include' });
    if (r.status === 404) return [];
    if (!r.ok) throw new Error(`Не получилось посмотреть папку (код ${r.status})`);
    const xml = await r.text();
    return [...xml.matchAll(/<d:href>([^<]+)<\/d:href>/g)]
      .map((m) => decodeURIComponent(m[1].replace(/\/$/, '').split('/').pop()))
      .filter((name) => /\.pdf$/i.test(name));
  }
  try {
    if (mode === 'list') return { files: await listPdfs() };

    let path = '';
    for (const name of folders) {
      path += `/${enc(name)}`;
      const r = await fetch(base + path, { method: 'MKCOL', headers, credentials: 'include' });
      if (!r.ok && r.status !== 405) return { error: `Папка «${name}» не создалась (код ${r.status})` }; // 405 — уже есть
    }
    // Старые сметы — в корзину ВБ Диска (оттуда можно восстановить)
    for (const old of remove || []) {
      if (old === fileName) continue; // файл с тем же именем просто заменится новым
      const r = await fetch(`${dir}/${enc(old)}`, { method: 'DELETE', headers, credentials: 'include' });
      if (!r.ok && r.status !== 404) return { error: `Старая смета «${old}» не удалилась (код ${r.status})` };
    }
    // Оставить обе — новый файл не должен затереть старый с тем же именем
    let name = fileName;
    if (keepBoth) {
      const have = await listPdfs();
      const dot = fileName.lastIndexOf('.');
      for (let i = 2; have.includes(name); i++) name = `${fileName.slice(0, dot)} (${i})${fileName.slice(dot)}`;
    }
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const filePath = `${dir}/${enc(name)}`;
    let r = await fetch(filePath, { method: 'PUT', headers: { ...headers, 'Content-Type': 'application/pdf' }, body: bytes, credentials: 'include' });
    if (!r.ok) return { error: `Файл не загрузился (код ${r.status})` };
    r = await fetch(filePath, {
      method: 'PROPFIND',
      headers: { ...headers, Depth: '0', 'Content-Type': 'application/xml' },
      credentials: 'include',
      body: '<?xml version="1.0"?><d:propfind xmlns:d="DAV:" xmlns:oc="http://owncloud.org/ns"><d:prop><oc:fileid/></d:prop></d:propfind>',
    });
    const m = (await r.text()).match(/<oc:fileid>(\d+)<\/oc:fileid>/);
    if (!m) return { error: 'Файл загрузился, но номер для ссылки не получен — вставьте ссылку вручную' };
    return { link: `${location.origin}/f/${m[1]}`, path: `${folders.join('/')}/${name}`, name };
  } catch (e) {
    return { error: `ВБ Диск: ${e.message}` };
  }
}
