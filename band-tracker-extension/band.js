// Band (на базе Mattermost): запоминаем, по какому сообщению кликнули правой кнопкой, —
// пункт меню «📮 Создать задачу в трекере» заберёт отсюда текст, автора и ссылку на сообщение.
// Сообщение в ленте — элемент с id «post_<id>» (в треде справа — «rhsPost_<id>»), ссылка — /<команда>/pl/<id>.
(() => {
  const POST_ID = /^(?:post|rhsPost|rhsRootPost)_([a-z0-9]{26})$/;

  const TEXT_ID = /^(?:rhsPostMessageText|postMessageText)_([a-z0-9]{26})$/;

  function findPost(el) {
    let textId = '';
    for (let n = el; n && n !== document.body; n = n.parentElement) {
      const m = n.id && POST_ID.exec(n.id);
      if (m) return { el: n, id: m[1] };
      const t = !textId && n.id && TEXT_ID.exec(n.id);
      if (t) textId = t[1];
    }
    // Рамку сообщения не нашли, но кликнули по тексту: id есть у самого текста (postMessageText_<id>)
    if (textId) {
      const box = document.getElementById(`post_${textId}`) || document.getElementById(`rhsPost_${textId}`);
      return { el: box || document.getElementById(`postMessageText_${textId}`) || document.getElementById(`rhsPostMessageText_${textId}`), id: textId };
    }
    return null;
  }

  // Имя автора у сообщений подряд не повторяется — берём из ближайшего сообщения выше с заголовком
  function authorOf(postEl) {
    for (let n = postEl; n; n = n.previousElementSibling) {
      const a = n.querySelector && n.querySelector('.post__header .user-popover, .post__header [class*="user-popover"], .post__header button');
      if (a && a.textContent.trim()) return a.textContent.trim();
    }
    return '';
  }

  // Текст сообщения → Markdown (трекер хранит описание так): ссылки под словами («тык»), жирный, списки, код.
  function toMarkdown(root) {
    let out = '';
    const kids = (el) => el.childNodes.forEach(walk);
    const block = (el, before = '', after = '\n') => { if (out && !out.endsWith('\n')) out += '\n'; out += before; kids(el); out += after; };
    function walk(node) {
      if (node.nodeType === 3) { out += node.nodeValue.replace(/\s+/g, ' '); return; }
      if (node.nodeType !== 1) return;
      const el = node;
      switch (el.tagName) {
        case 'BR': out += '\n'; return;
        case 'IMG': out += el.getAttribute('alt') || ''; return;
        case 'A': {
          const href = el.getAttribute('href') || '';
          const text = el.textContent.trim();
          // ссылка спрятана под словом — сохраняем адрес; голые адреса и @упоминания — как есть
          if (/^https?:/i.test(href) && text && text !== href && !el.classList.contains('mention-link')) out += `[${text}](${href})`;
          else out += text;
          return;
        }
        case 'STRONG': case 'B': out += '**'; kids(el); out += '**'; return;
        case 'EM': case 'I': out += '_'; kids(el); out += '_'; return;
        case 'DEL': case 'S': out += '~~'; kids(el); out += '~~'; return;
        case 'CODE': if (el.parentElement && el.parentElement.tagName === 'PRE') break; out += '`' + el.textContent + '`'; return;
        case 'PRE': block(el, '```\n', '\n```\n'); return;
        case 'LI': block(el, el.parentElement && el.parentElement.tagName === 'OL' ? `${[...el.parentElement.children].indexOf(el) + 1}. ` : '- '); return;
        case 'H1': case 'H2': case 'H3': case 'H4': case 'H5': case 'H6': block(el, '**', '**\n'); return;
        case 'BLOCKQUOTE': block(el, '> '); return;
        case 'P': case 'DIV': case 'UL': case 'OL': case 'TABLE': case 'TR': block(el); return;
        default: break;
      }
      kids(el);
    }
    walk(root);
    return out.split('\n').map((l) => l.replace(/[ \t]+$/, '').replace(/^ +/, '')).join('\n').replace(/\n{3,}/g, '\n\n').trim();
  }

  document.addEventListener('contextmenu', (e) => {
    const selection = String(window.getSelection() || '').trim();
    const post = e.target instanceof Element ? findPost(e.target) : null;
    const team = location.pathname.split('/')[1] || '';
    if (!post) {
      window.__lqaBandPost = { selection, text: '', author: '', link: '' };
      return;
    }
    const textEl = document.getElementById(`postMessageText_${post.id}`) || document.getElementById(`rhsPostMessageText_${post.id}`) ||
      post.el.querySelector('.post-message__text') || (post.el.classList.contains('post-message__text') ? post.el : null);
    window.__lqaBandPost = {
      selection,
      text: textEl ? toMarkdown(textEl) : '',
      author: authorOf(post.el),
      link: team ? `${location.origin}/${team}/pl/${post.id}` : '',
      postId: post.id,
    };
  }, true);

  // ================= Кнопка «Создать тикет» в треде =================
  // Наводите на сообщение — внизу справа кнопка. Клик — карточка: название, языки → подзадачи, «ответить в тред».
  // Всё рисуем в своём shadow DOM поверх страницы: вёрстку Band не трогаем, его стили нам не мешают.
  const postData = (post) => {
    const textEl = document.getElementById(`postMessageText_${post.id}`) || document.getElementById(`rhsPostMessageText_${post.id}`) ||
      post.el.querySelector('.post-message__text, [class*="post-message__text"], .post__body');
    const team = location.pathname.split('/')[1] || '';
    return { postId: post.id, text: textEl ? toMarkdown(textEl) : '', author: authorOf(post.el), link: team ? `${location.origin}/${team}/pl/${post.id}` : '' };
  };

  // ---- Band (Mattermost) API от имени того, кто вошёл: тред и ответ в него ----
  const csrf = () => decodeURIComponent((document.cookie.match(/(?:^|;\s*)MMCSRF=([^;]*)/) || [])[1] || '');
  async function mm(path, opts = {}) {
    const r = await fetch(`/api/v4${path}`, {
      credentials: 'include', ...opts,
      headers: { 'X-Requested-With': 'XMLHttpRequest', 'X-CSRF-Token': csrf(), 'Content-Type': 'application/json', ...(opts.headers || {}) },
    });
    if (!r.ok) throw new Error(`Band ответил ${r.status}`);
    return r.json();
  }
  const keysIn = (t) => [...String(t || '').matchAll(/tracker\.wb\.ru\/i\/[^/\s)]+\/([A-Z][A-Z0-9_]*-\d+)/g)].map((m) => m[1]);
  /** Тикеты, уже упомянутые в треде (ответ расширения со ссылками) — чтобы не создать дубль. */
  async function threadKeys(postId) {
    const p = await mm(`/posts/${postId}`);
    const t = await mm(`/posts/${p.root_id || p.id}/thread`);
    return [...new Set(Object.values(t.posts || {}).flatMap((x) => keysIn(x.message)))];
  }
  async function replyInThread(postId, message) {
    const p = await mm(`/posts/${postId}`);
    await mm('/posts', { method: 'POST', body: JSON.stringify({ channel_id: p.channel_id, root_id: p.root_id || p.id, message }) });
  }
  // панель просит то же самое
  chrome.runtime.onMessage.addListener((msg, _s, reply) => {
    if (msg && msg.type === 'ping') {
      const posts = document.querySelectorAll('[id^="post_"], [id^="rhsPost_"], [id^="rhsRootPost_"]').length;
      reply({ ok: true, posts, texts: document.querySelectorAll('[id^="postMessageText_"], [id^="rhsPostMessageText_"]').length, version: chrome.runtime.getManifest().version });
      return false;
    }
    if (msg && msg.type === 'bandThread') { threadKeys(msg.postId).then((keys) => reply({ ok: true, keys }), (e) => reply({ ok: false, error: e.message })); return true; }
    if (msg && msg.type === 'bandReply') { replyInThread(msg.postId, msg.message).then(() => reply({ ok: true }), (e) => reply({ ok: false, error: e.message })); return true; }
    return false;
  });

  const replyText = (res) => `Завели в трекере: [${res.main.key}](${res.main.url})` +
    (res.subs.length ? ` — основная задача; подзадачи: ${res.subs.map((s) => `[${s.key}](${s.url}) (${s.lang})`).join(', ')}` : '');

  // ---- интерфейс ----
  const host = document.createElement('div');
  host.style.cssText = 'position:fixed;top:0;left:0;width:0;height:0;z-index:2147483000;';
  const root = host.attachShadow({ mode: 'open' });
  root.innerHTML = `<style>
    :host { all: initial; }
    * { box-sizing: border-box; font-family: system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif; }
    .btn { position: fixed; display: inline-flex; align-items: center; gap: 6px; padding: 5px 11px; border: 0; border-radius: 8px; cursor: pointer;
      font: 700 12.5px system-ui, sans-serif; color: #fff; background: linear-gradient(90deg, #8B5CF6, #EC4899); box-shadow: 0 4px 12px rgb(139 92 246 / .35); }
    .btn.has { background: #F2EEFF; color: #6D28D9; box-shadow: 0 2px 8px rgb(0 0 0 / .12); border: 1px solid #DDD3FB; }
    .card { position: fixed; width: 340px; max-height: calc(100vh - 16px); display: flex; flex-direction: column; overflow: hidden; background: #FCFBFF; color: #2A2140; border: 1.5px solid #CDBDFB;
      border-radius: 16px; box-shadow: 0 18px 40px rgb(42 33 64 / .25); font-size: 13px; line-height: 1.45; }
    .top { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 10px 12px; border-bottom: 1px solid #E7E1F7; font-weight: 700; }
    .x { border: 0; background: none; font-size: 16px; cursor: pointer; color: #7A7194; }
    .body { padding: 10px 12px 12px; display: flex; flex-direction: column; gap: 9px; overflow: auto; min-height: 0; }
    .foot { padding: 10px 12px; border-top: 1px solid #E7E1F7; background: #FCFBFF; }
    .foot:empty { display: none; }
    .miss { color: #A79DC0; font-weight: 400; }
    .bz { display: flex; align-items: center; gap: 8px; font-size: 12px; color: #7A7194; }
    .bz select { flex: 1; min-width: 0; font: 600 13px system-ui, sans-serif; color: #2A2140; border: 1px solid #E7E1F7; border-radius: 9px; padding: 6px 8px; background: #fff; }
    .lbl { font-size: 10.5px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; color: #8B5CF6; }
    input[type=text] { width: 100%; font: 600 13px system-ui, sans-serif; color: #2A2140; border: 1px solid #E7E1F7; border-radius: 9px; padding: 7px 9px; background: #fff; }
    input[type=text]:focus { outline: none; border-color: #8B5CF6; }
    .meta { display: flex; flex-wrap: wrap; gap: 4px 12px; color: #7A7194; font-size: 12px; } .meta b { color: #2A2140; font-weight: 600; }
    .langs { display: flex; flex-wrap: wrap; gap: 6px; }
    .lang { display: inline-flex; align-items: center; gap: 5px; padding: 3px 9px; border: 1px solid #E7E1F7; border-radius: 999px; background: #fff; cursor: pointer; font-size: 12.5px; }
    .lang input { accent-color: #8B5CF6; margin: 0; }
    .chk { display: flex; gap: 7px; align-items: flex-start; font-size: 12.5px; } .chk input { accent-color: #8B5CF6; margin-top: 2px; }
    .note { font-size: 11.5px; color: #7A7194; margin: 0; }
    .row { display: flex; gap: 6px; flex-wrap: wrap; }
    .row button { flex: 1; font: 700 13px system-ui, sans-serif; border-radius: 10px; padding: 8px 10px; cursor: pointer; border: 1px solid #E7E1F7; background: #fff; color: #2A2140; }
    .row button.go { border: 0; color: #fff; background: linear-gradient(90deg, #8B5CF6, #EC4899); }
    .row button:disabled { opacity: .6; cursor: wait; }
    .ok { background: #EEFBF2; color: #15803D; border-radius: 10px; padding: 8px 10px; font-weight: 600; }
    .warn { background: #FFF4EA; color: #C2410C; border-radius: 10px; padding: 8px 10px; }
    .err { background: #FFEFF6; color: #BE185D; border-radius: 10px; padding: 8px 10px; }
    a { color: #6D28D9; font-weight: 600; }
  </style><button class="btn" id="btn" hidden></button><div class="card" id="card" hidden></div>`;
  document.documentElement.appendChild(host);
  const $ = (s) => root.getElementById(s);
  const ICON = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v3a2 2 0 0 0 0 4v3a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-3a2 2 0 0 0 0-4z"/></svg>';
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const send = (m) => new Promise((ok) => { try { chrome.runtime.sendMessage(m, (r) => ok(r || { ok: false, error: chrome.runtime.lastError && chrome.runtime.lastError.message })); } catch (e) { ok({ ok: false, error: 'Расширение обновилось — обновите страницу Band (F5)' }); } });

  let hovered = null;   // { el, id } — сообщение под мышью
  let active = null;    // сообщение, для которого открыта карточка
  let hideTimer = 0;
  const known = new Map(); // postId → номер основной задачи (из журнала/треда)

  function placeButton(post) {
    const r = post.el.getBoundingClientRect();
    const b = $('btn');
    b.innerHTML = `${ICON}${known.get(post.id) ? `Тикет ${esc(known.get(post.id))}` : 'Создать тикет'}`;
    b.classList.toggle('has', known.has(post.id));
    b.hidden = false;
    b.style.top = `${Math.max(8, Math.min(r.bottom - 34, innerHeight - 40))}px`;
    b.style.left = `${Math.max(8, r.right - b.offsetWidth - 14)}px`;
  }
  document.addEventListener('mouseover', (e) => {
    if (e.composedPath().includes(host)) { clearTimeout(hideTimer); return; }
    const post = e.target instanceof Element ? findPost(e.target) : null;
    if (post && post.el) {
      clearTimeout(hideTimer);
      hovered = post;
      placeButton(post);
    } else if (!$('card').hidden) {
      /* карточка открыта — кнопку не прячем */
    } else {
      clearTimeout(hideTimer);
      hideTimer = setTimeout(() => { $('btn').hidden = true; }, 350);
    }
  }, true);
  // Прокрутка ленты — кнопка едет за сообщением (раньше пряталась от любой прокрутки на странице)
  addEventListener('scroll', () => {
    if (!$('card').hidden || $('btn').hidden || !hovered) return;
    if (!document.contains(hovered.el)) { $('btn').hidden = true; return; }
    placeButton(hovered);
  }, true);

  $('btn').addEventListener('click', () => { if (hovered) openCard(hovered); });

  async function openCard(post) {
    active = { ...postData(post), el: post.el };
    const card = $('card');
    card.hidden = false;
    card.innerHTML = '<div class="top">Тикет по сообщению<button class="x" id="x">✕</button></div><div class="body"><p class="note">Читаю сообщение и тред…</p></div><div class="foot" id="foot"></div>';
    $('x').onclick = closeCard;
    fitCard();
    if (!active.text) return body('<div class="warn">Не вижу текста сообщения. Если вкладку Band открыли до установки расширения — обновите её (F5).</div>');
    const [plan, thread] = await Promise.all([send({ type: 'plan', post: active }), threadKeys(active.postId).catch(() => null)]);
    if (!plan.ok) return body(`<div class="err">${esc(plan.error)}</div>`);
    const existing = [...new Set([...(thread || []), ...(plan.journal ? [plan.journal.main.key, ...plan.journal.subs.map((s) => s.key)] : [])])];
    if (existing.length) known.set(active.postId, existing[0]);
    renderForm(plan, existing);
  }
  function body(html, foot = '') { $('card').querySelector('.body').innerHTML = html; $('foot').innerHTML = foot; fitCard(); }
  // Карточка всегда целиком на экране: под кнопкой, а если места мало — над ней; кнопки действий внизу видны всегда
  function fitCard() {
    const card = $('card'), b = $('btn').getBoundingClientRect();
    card.style.left = `${Math.max(8, Math.min(b.right - 340, innerWidth - 348))}px`;
    const h = card.offsetHeight;
    let top = b.bottom + 6;
    if (top + h > innerHeight - 8) top = b.top - h - 6;
    card.style.top = `${Math.max(8, Math.min(top, innerHeight - h - 8))}px`;
  }
  addEventListener('resize', () => { if (!$('card').hidden) fitCard(); });
  // «Страны Грузия, Казахстан?» — подставится / ещё не знаю код в трекере
  function fieldLine(label, ok, miss) {
    if (!ok.length && !miss.length) return '';
    return `<span>${label} <b>${ok.map(esc).join(', ')}</b>${miss.length ? `${ok.length ? ', ' : ''}<span class="miss" title="Расширение ещё не знает этот вариант — выберите в задаче сами">${miss.map(esc).join(', ')} (выберите в трекере)</span>` : ''}</span>`;
  }
  function closeCard() { $('card').hidden = true; active = null; }

  function renderForm(plan, existing) {
    const o = plan.order;
    body(`
      ${existing.length ? `<div class="warn">По этому треду уже есть: <b>${existing.map(esc).join(', ')}</b>. Создать ещё — только если это другая задача.</div>` : ''}
      <span class="lbl">Основная задача</span>
      <input type="text" id="title" value="${esc(o.title)}" aria-label="Название задачи">
      ${(plan.bzOptions || []).length ? `<label class="bz"><span>БЗ</span><select id="bz"><option value="">— не указывать —</option>${plan.bzOptions.map((x) => `<option value="${esc(x.id)}"${x.label === plan.picked.bz ? ' selected' : ''}>${esc(x.label)}</option>`).join('')}</select></label>` : ''}
      <div class="meta">${(plan.bzOptions || []).length ? '' : fieldLine('БЗ', plan.picked.bz ? [plan.picked.bz] : [], plan.missing.bz ? [plan.missing.bz] : [])}${fieldLine('Страны', plan.picked.countries || [], plan.missing.countries || [])}${o.due ? `<span>Дедлайн <b>${esc(o.due.slice(8, 10) + '.' + o.due.slice(5, 7))}</b></span>` : ''}<span>Тип <b>${esc(plan.board || '—')}</b></span></div>
      ${(plan.missing.bz || (plan.missing.countries || []).length) ? '<p class="note">Серым — варианты, которых нет в справочниках трекера (или трекер их ещё не отдал): их выберите в задаче сами.</p>' : ''}
      ${o.languages.length ? `<span class="lbl">Подзадачи по языкам</span><div class="langs">${o.languages.map((l) => `<label class="lang"><input type="checkbox" value="${l.code}"${plan.subLangs.includes(l.code) ? ' checked' : ''}>${esc(l.code)} · ${esc(l.name)}</label>`).join('')}</div>
        <p class="note">Отметьте языки штатных переводчиков — на каждый будет своя подзадача. Расширение запомнит выбор.</p>` : ''}
      <label class="chk"><input type="checkbox" id="reply"${plan.replyInThread ? ' checked' : ''}><span>Ответить в тред со ссылками на тикеты (по этому ответу видно, что тикет уже есть)</span></label>`,
      '<div class="row"><button type="button" id="more">Все поля…</button><button type="button" class="go" id="go">Создать</button></div>');
    const count = () => { const n = root.querySelectorAll('.langs input:checked').length; $('go').textContent = n ? `Создать ${n + 1} тикета` : 'Создать тикет'; };
    root.querySelectorAll('.langs input').forEach((i) => i.addEventListener('change', count));
    count();
    $('go').onclick = () => create(plan);
    $('more').onclick = async () => {
      const r = await send({ type: 'openPanel', draft: { text: active.text, author: active.author, link: active.link, postId: active.postId, title: $('title').value } });
      if (r.ok) closeCard();
      else body('<div class="warn">Откройте панель значком расширения на панели Chrome — форма уже заполнена.</div>');
    };
  }

  async function create(plan) {
    const go = $('go');
    go.disabled = true; go.textContent = 'Создаю…';
    const subLangs = [...root.querySelectorAll('.langs input:checked')].map((i) => i.value);
    const wantReply = $('reply').checked;
    const post = { postId: active.postId, text: active.text, author: active.author, link: active.link };
    const res = await send({ type: 'create', post, title: $('title').value, subLangs, reply: wantReply, bzId: $('bz') ? $('bz').value : '' });
    if (!res.ok) {
      go.disabled = false; go.textContent = 'Попробовать ещё раз';
      $('card').querySelector('.body').insertAdjacentHTML('afterbegin', `<div class="err">${esc(res.error)}</div>`);
      return;
    }
    known.set(post.postId, res.main.key);
    let replied = '';
    if (wantReply) {
      try { await replyInThread(post.postId, replyText(res)); replied = '<p class="note">✓ Ответ со ссылками отправлен в тред.</p>'; }
      catch (e) { replied = `<div class="warn">В тред ответить не вышло (${esc(e.message)}). <a href="#" id="cp">Скопировать ответ</a> и вставьте сами.</div>`; }
    }
    body(`<div class="ok">✓ Создано: ${[res.main, ...res.subs].length} ${res.subs.length ? 'тикета' : 'тикет'}</div>
      <div><a href="${esc(res.main.url)}" target="_blank" rel="noopener">${esc(res.main.key)}</a> — основная${res.subs.map((s) => `<br><a href="${esc(s.url)}" target="_blank" rel="noopener">${esc(s.key)}</a> — ${esc(s.lang)}`).join('')}</div>
      ${replied}`, '<div class="row"><button type="button" id="done">Готово</button></div>');
    $('done').onclick = closeCard;
    const cp = $('cp');
    if (cp) cp.onclick = (e) => { e.preventDefault(); navigator.clipboard.writeText(replyText(res)); cp.textContent = 'Скопировано ✓'; };
  }
  addEventListener('keydown', (e) => { if (e.key === 'Escape' && !$('card').hidden) closeCard(); });
})();
