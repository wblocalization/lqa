// Работает на странице tracker.wb.ru. Когда человек сам создаёт задачу в своей доске, запоминаем,
// КУДА трекер её отправил и с какими полями, — так расширение узнаёт новые доски без настроек.
// Берём только адрес, поля доски и номер задачи; куки и токены не трогаем.
(() => {
  const CREATE = /^\/api\/gateway\/catalog\/[^?#]+$/;
  const pathOf = (url) => { try { return new URL(url, location.href).pathname; } catch { return ''; } };

  function report(url, body, answer) {
    try {
      const req = typeof body === 'string' ? JSON.parse(body) : null;
      const res = typeof answer === 'string' ? JSON.parse(answer) : answer;
      if (!req || !req.issue || typeof req.issue.title !== 'string') return;
      if (!res || !res.success || typeof res.data !== 'string') return;
      window.postMessage({
        __bandTracker: 'board', path: pathOf(url), fields: req.fields || {}, key: res.data,
        page: location.pathname, title: document.title,
      }, location.origin);
    } catch { /* не задача — не наше дело */ }
  }

  const isCreate = (method, url) => String(method || 'GET').toUpperCase() === 'POST' && CREATE.test(pathOf(url));

  const origFetch = window.fetch;
  window.fetch = function (input, init) {
    const p = origFetch.apply(this, arguments);
    try {
      const url = typeof input === 'string' ? input : (input && input.url) || String(input);
      const method = (init && init.method) || (input && input.method);
      const body = init && init.body;
      // задачи, которые создаёт само расширение, помечены — доску из них не запоминаем
      if (isCreate(method, url) && typeof body === 'string' && !(init && init.bandTrackerOwn)) {
        p.then((r) => r.clone().text()).then((t) => report(url, body, t)).catch(() => {});
      }
    } catch { /* страница работает как обычно */ }
    return p;
  };

  const { open, send } = XMLHttpRequest.prototype;
  XMLHttpRequest.prototype.open = function (method, url) {
    this.__bandTracker = { method, url };
    return open.apply(this, arguments);
  };
  XMLHttpRequest.prototype.send = function (body) {
    const m = this.__bandTracker;
    if (m && isCreate(m.method, m.url) && typeof body === 'string') {
      this.addEventListener('load', () => {
        try {
          const answer = this.responseType === 'json' ? this.response : this.responseText;
          report(m.url, body, answer);
        } catch { /* пропустим */ }
      });
    }
    return send.apply(this, arguments);
  };
})();
