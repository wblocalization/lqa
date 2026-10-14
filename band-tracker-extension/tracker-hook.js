// Работает на странице tracker.wb.ru. Когда человек сам создаёт задачу, запоминаем,
// КУДА трекер её отправил и с какими полями, — так расширение узнаёт новые типы задач без настроек.
// Адрес — /api/gateway/catalog/<пространство>/<проект>/<тип>. Берём только адрес, поля, названия и номер задачи; куки и токены не трогаем.
(() => {
  // Создание задачи: обычно POST /api/gateway/catalog/<пространство>/<проект>/<тип>; смотрим и другие POST в /api/ с полем title
  const CREATE = /^\/api\/gateway\/catalog\/[^?#]+$/;
  const API = /^\/api\//;
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

  // Номер задачи в ответе трекера: { success, data: 'LOCAL-1' } или вложенный вариант
  const keyOf = (res) => {
    const cands = [res && res.data, res && res.data && res.data.key, res && res.data && res.data.code, res && res.key, res && res.issueKey];
    return cands.find((x) => typeof x === 'string' && /^[A-Z][A-Z0-9_]*-\d+$/.test(x)) || '';
  };
  // Для диагностики («Скопировать для разработчика»): только устройство запроса — без названия, описания, кук и токенов
  const short = (v) => (typeof v === 'string' ? v.slice(0, 60) : v);
  const shape = (o) => (o && typeof o === 'object' ? Object.keys(o).slice(0, 20) : typeof o);
  function fieldsDigest(fields) {
    const out = {};
    for (const [k, v] of Object.entries(fields || {})) {
      out[k] = Array.isArray(v) ? v.slice(0, 10).map((x) => (x && typeof x === 'object' ? { id: short(x.id), label: short(x.label || x.name || x.title) } : short(x))) : short(v);
    }
    return out;
  }

  function report(url, method, status, body, answer, names) {
    let req = null, res = null, why = '';
    try { req = typeof body === 'string' ? JSON.parse(body) : null; } catch { why = 'тело не JSON'; }
    try { res = typeof answer === 'string' ? JSON.parse(answer) : answer; } catch { /* ответ не JSON */ }
    if (!req) why = why || 'тело не текст';
    const issue = req && (req.issue || req);
    const fields = req && (req.fields || (req.issue && req.issue.fields)) || {};
    const key = keyOf(res);
    if (!why && (!issue || typeof issue.title !== 'string')) why = 'нет issue.title';
    if (!why && !key) why = 'в ответе нет номера задачи';
    // что видели — в «Типах задач» → «Скопировать для разработчика»
    window.postMessage({
      __bandTracker: 'seen', path: pathOf(url), method, status, why, ok: !why,
      body: shape(req), issue: shape(issue), res: shape(res), resData: res && typeof res.data,
      fields: fieldsDigest(fields), executorId: req && req.executorId !== undefined ? short(String(req.executorId)) : undefined, names: names || {},
    }, location.origin);
    if (why) return;
    window.postMessage({
      __bandTracker: 'board', path: pathOf(url), fields, key,
      page: location.pathname, title: document.title, names: names || {},
    }, location.origin);
  }

  const isCreate = (method, url, body) => String(method || 'GET').toUpperCase() === 'POST' &&
    (CREATE.test(pathOf(url)) || (API.test(pathOf(url)) && typeof body === 'string' && /"title"\s*:/.test(body) && /"issue"|"fields"/.test(body)));

  const origFetch = window.fetch;
  window.fetch = function (input, init) {
    // тело запроса — из init или из самого Request (его читаем копией до отправки)
    const reqBody = init && init.body !== undefined ? Promise.resolve(init.body)
      : (input instanceof Request ? input.clone().text().catch(() => null) : Promise.resolve(null));
    const p = origFetch.apply(this, arguments);
    try {
      const url = typeof input === 'string' ? input : (input && input.url) || String(input);
      const method = (init && init.method) || (input && input.method) || 'GET';
      // задачи, которые создаёт само расширение, помечены — тип из них не запоминаем
      if (String(method).toUpperCase() === 'POST' && API.test(pathOf(url)) && !(init && init.bandTrackerOwn)) {
        const names = snapshot();
        Promise.all([reqBody, p.then((r) => r.clone().text().then((t) => [r.status, t]))]).then(([body, [status, t]]) => {
          if (isCreate(method, url, body)) report(url, method, status, body, t, names);
        }).catch(() => {});
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
    if (m && isCreate(m.method, m.url, body)) {
      const names = snapshot();
      this.addEventListener('load', () => {
        try {
          const answer = this.responseType === 'json' ? this.response : this.responseText;
          report(m.url, m.method, this.status, body, answer, names);
        } catch { /* пропустим */ }
      });
    }
    return send.apply(this, arguments);
  };
})();
