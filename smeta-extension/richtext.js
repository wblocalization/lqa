// Поле письма «как в почте»: жирный виден жирным, а вместо {языки}, {срок}… — цветные «таблетки»
// с тем, что подставится (например «кыргызский, казахский», «22.10»). Таблетку нельзя сломать — только удалить целиком.
// Внутри храним обычный текст: **жирный**, {переменная}, переносы строк.

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Как называется переменная, если для неё ещё нет значения. */
export const VAR_LABELS = {
  языки: 'языки', срок: 'срок', номер: 'номер', продукт: 'продукт', коды: 'коды языков',
  ссылка: 'ссылка на Band', менеджер: 'менеджер', тема: 'тема', материалы: 'материалы',
};

function chip(name, valueOf) {
  const v = valueOf ? valueOf(name) : '';
  const label = VAR_LABELS[name] || name;
  return `<span class="var-pill${v ? '' : ' empty'}" contenteditable="false" data-var="${esc(name)}" title="Подставится само: ${esc(label)}">${esc(v || label)}</span>`;
}

/** Текст (**жирный**, {переменная}) → HTML для поля. valueOf(name) — что показать на таблетке. */
export function textToRich(text, valueOf) {
  return String(text || '').split('\n').map((line) => {
    let h = esc(line);
    h = h.replace(/\{([а-яё]+)\}/gi, (m, v) => chip(v.toLowerCase(), valueOf));
    h = h.replace(/\*\*(.+?)\*\*/g, '<b>$1</b>');
    return h;
  }).join('<br>');
}

/** HTML поля → текст (**жирный**, {переменная}, переносы строк). */
export function richToText(root) {
  let out = '';
  const bold = (el) => {
    const w = el.style && el.style.fontWeight;
    return el.tagName === 'B' || el.tagName === 'STRONG' || w === 'bold' || Number(w) >= 600;
  };
  function walk(node, inBold) {
    if (node.nodeType === 3) { out += node.nodeValue.replace(/ /g, ' ').replace(/​/g, ''); return; }
    if (node.nodeType !== 1) return;
    const el = node;
    if (el.dataset && el.dataset.var) { out += `{${el.dataset.var}}`; return; }
    if (el.tagName === 'BR') { out += '\n'; return; }
    const block = el.tagName === 'DIV' || el.tagName === 'P';
    // Enter в поле Chrome оборачивает строку в <div> — это новая строка
    if (block && out && !out.endsWith('\n')) out += '\n';
    const b = !inBold && bold(el);
    if (b) out += '**';
    el.childNodes.forEach((n) => walk(n, inBold || b));
    if (b) out += '**';
    if (block && el.nextSibling && !out.endsWith('\n')) out += '\n';
  }
  root.childNodes.forEach((n) => walk(n, false));
  return out.replace(/\*\*\*\*/g, '').replace(/\*\*(\s*)\*\*/g, '$1').replace(/\n+$/, '');
}

/**
 * Сделать <div contenteditable> редактором. Возвращает { get, set, bold, isBold, insertVar, focus }.
 * Вставка из буфера — только текстом (без чужих шрифтов и цветов).
 */
export function richEditor(el, { onInput, valueOf, onSelect } = {}) {
  el.addEventListener('paste', (e) => {
    e.preventDefault();
    const text = (e.clipboardData && e.clipboardData.getData('text/plain')) || '';
    document.execCommand('insertText', false, text);
  });
  el.addEventListener('input', () => { if (onInput) onInput(); });
  // курсор внутри поля — запоминаем, чтобы кнопки вставляли туда, где он был
  let saved = null;
  const remember = () => {
    const sel = window.getSelection();
    if (sel.rangeCount && el.contains(sel.anchorNode)) saved = sel.getRangeAt(0).cloneRange();
  };
  ['keyup', 'mouseup', 'input', 'focus'].forEach((t) => el.addEventListener(t, remember));
  document.addEventListener('selectionchange', () => {
    const sel = window.getSelection();
    if (sel.rangeCount && el.contains(sel.anchorNode)) { remember(); if (onSelect) onSelect(); }
  });
  const restore = () => {
    el.focus();
    const sel = window.getSelection();
    if (saved && el.contains(saved.startContainer)) { sel.removeAllRanges(); sel.addRange(saved); return; }
    const r = document.createRange(); r.selectNodeContents(el); r.collapse(false);
    sel.removeAllRanges(); sel.addRange(r);
  };
  const hasSelection = () => Boolean(saved && el.contains(saved.startContainer) && !saved.collapsed);
  return {
    get: () => richToText(el),
    set: (text) => { el.innerHTML = textToRich(text, valueOf); saved = null; },
    hasSelection,
    /** Выделенное уже жирное? (тогда кнопка снимает жирный) */
    isBold: () => {
      if (!hasSelection()) return false;
      let n = saved.commonAncestorContainer;
      for (; n && n !== el; n = n.parentNode) if (n.nodeType === 1 && /^(B|STRONG)$/.test(n.tagName)) return true;
      return false;
    },
    /** Жирный для выделения (повторно — снять). Ничего не выделено — false. */
    bold: () => {
      if (!hasSelection()) return false;
      restore();
      document.execCommand('bold');
      remember();
      if (onInput) onInput();
      return true;
    },
    insertVar: (name) => {
      restore();
      document.execCommand('insertHTML', false, chip(name, valueOf) + '&#8203;');
      remember();
      if (onInput) onInput();
    },
    focus: () => el.focus(),
  };
}
