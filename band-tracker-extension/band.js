// Band (на базе Mattermost): запоминаем, по какому сообщению кликнули правой кнопкой, —
// пункт меню «📮 Создать задачу в трекере» заберёт отсюда текст, автора и ссылку на сообщение.
// Сообщение в ленте — элемент с id «post_<id>» (в треде справа — «rhsPost_<id>»), ссылка — /<команда>/pl/<id>.
(() => {
  const POST_ID = /^(?:post|rhsPost|rhsRootPost)_([a-z0-9]{26})$/;

  function findPost(el) {
    for (let n = el; n && n !== document.body; n = n.parentElement) {
      const m = n.id && POST_ID.exec(n.id);
      if (m) return { el: n, id: m[1] };
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

  document.addEventListener('contextmenu', (e) => {
    const selection = String(window.getSelection() || '').trim();
    const post = e.target instanceof Element ? findPost(e.target) : null;
    const team = location.pathname.split('/')[1] || '';
    if (!post) {
      window.__lqaBandPost = { selection, text: '', author: '', link: '' };
      return;
    }
    const textEl = post.el.querySelector(`#postMessageText_${post.id}, #rhsPostMessageText_${post.id}, .post-message__text`);
    window.__lqaBandPost = {
      selection,
      text: textEl ? textEl.innerText.trim() : '',
      author: authorOf(post.el),
      link: team ? `${location.origin}/${team}/pl/${post.id}` : '',
    };
  }, true);
})();
