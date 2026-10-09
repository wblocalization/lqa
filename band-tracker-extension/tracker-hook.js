// Работает на странице tracker.wb.ru. Когда человек сам создаёт задачу, запоминаем,
// КУДА трекер её отправил и с какими полями, — так расширение узнаёт новые типы задач без настроек.
// Адрес — /api/gateway/catalog/<пространство>/<проект>/<тип>. Берём только адрес, поля, названия и номер задачи; куки и токены не трогаем.
(() => {
  const CREATE = /^\/api\/gateway\/catalog\/[^?#]+$/;
  const pathOf = (url) => { try { return new URL(url, location.href).pathname; } catch { return ''; } };

  // Окно «Создание задачи» ещё открыто: справа «Пространство / Проект / Тип» — берём их названия
  function formValue(label) {
    for (const el of document.querySelectorAll('div, span, td, dt, label, p')) {
      if (el.children.length || el.textContent.trim() !== label) continue;
      for (let n = el; n && n !== document.body; n = n.parentElement) {
        const next = n.nextElementSibling;
        if (next && next.textContent.trim()) return next.textContent.trim().slice(0, 60);
      }
    }
    return '';
  }
  // Читаем сразу при отправке — к ответу сервера окно уже может закрыться
  const snapshot = () => ({ space: formValue('Пространство'), project: formValue('Проект'), type: formValue('Тип') });

  function report(url, body, answer, names) {
    try {
      const req = typeof body === 'string' ? JSON.parse(body) : null;
      const res = typeof answer === 'string' ? JSON.parse(answer) : answer;
      if (!req || !req.issue || typeof req.issue.title !== 'string') return;
      if (!res || !res.success || typeof res.data !== 'string') return;
      window.postMessage({
        __bandTracker: 'board', path: pathOf(url), fields: req.fields || {}, key: res.data,
        page: location.pathname, title: document.title, names: names || {},
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
      // задачи, которые создаёт само расширение, помечены — тип из них не запоминаем
      if (isCreate(method, url) && typeof body === 'string' && !(init && init.bandTrackerOwn)) {
        const names = snapshot();
        p.then((r) => r.clone().text()).then((t) => report(url, body, t, names)).catch(() => {});
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
      const names = snapshot();
      this.addEventListener('load', () => {
        try {
          const answer = this.responseType === 'json' ? this.response : this.responseText;
          report(m.url, body, answer, names);
        } catch { /* пропустим */ }
      });
    }
    return send.apply(this, arguments);
  };
})();
