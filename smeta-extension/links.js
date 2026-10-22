// «Ещё ссылки на Band» — отдельное поле на каждую ссылку. Заполнили последнее — появляется следующее,
// опустевшее поле пропадает. Вставили сразу несколько ссылок — разложатся по полям.
// root.value работает как раньше: ссылки через перевод строки (так их ждёт таблица).
export function makeLinkList(root, firstNumber = 2) {
  const inputs = () => [...root.querySelectorAll('input')];
  const make = (val = '') => {
    const el = document.createElement('input');
    el.type = 'url';
    el.value = val;
    return el;
  };
  const renumber = () => inputs().forEach((el, i) => {
    el.placeholder = `ссылка ${i + firstNumber} — https://band.wb.ru/…`;
    el.setAttribute('aria-label', `Ссылка ${i + firstNumber}`);
  });
  // Пустые поля посередине убираем, в конце всегда одно пустое
  const tidy = () => {
    const all = inputs();
    all.slice(0, -1).forEach((el) => { if (!el.value.trim() && el !== document.activeElement) el.remove(); });
    const rest = inputs();
    if (!rest.length || rest[rest.length - 1].value.trim()) root.append(make());
    renumber();
  };

  root.classList.add('link-list');
  root.addEventListener('input', (e) => {
    const all = inputs();
    if (e.target === all[all.length - 1] && e.target.value.trim()) { root.append(make()); renumber(); }
  });
  root.addEventListener('focusout', () => setTimeout(tidy, 0));
  root.addEventListener('paste', (e) => {
    const parts = (e.clipboardData?.getData('text') || '').split(/[\s,;]+/).filter(Boolean);
    if (parts.length < 2 || e.target.tagName !== 'INPUT') return;
    e.preventDefault();
    e.target.value = parts.shift();
    let after = e.target;
    parts.forEach((p) => { const el = make(p); after.after(el); after = el; });
    tidy();
    root.dispatchEvent(new Event('input', { bubbles: true }));
  });

  Object.defineProperty(root, 'value', {
    configurable: true,
    get: () => inputs().map((el) => el.value.trim()).filter(Boolean).join('\n'),
    set: (v) => {
      root.replaceChildren(...String(v || '').split(/[\s,;]+/).filter(Boolean).map((u) => make(u)), make());
      renumber();
    },
  });
  root.value = '';
  return root;
}
