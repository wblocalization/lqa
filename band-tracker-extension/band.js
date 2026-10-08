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
    };
  }, true);
})();
