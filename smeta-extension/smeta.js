import * as pdfjs from './vendor/pdf.min.mjs';
import { parseEstimate, formatMoney } from './parser.js';
import { isConfigured, api, esc, readClipboard } from './core.js';

pdfjs.GlobalWorkerOptions.workerSrc = chrome.runtime.getURL('vendor/pdf.worker.min.mjs');

const $ = (s) => document.querySelector(s);
const els = {
  empty: $('#empty'), drop: $('#drop'), file: $('#file'), form: $('#form'), fileName: $('#fileName'),
  reset: $('#reset'), done: $('#done'), doneText: $('#doneText'), again: $('#again'),
  task: $('#task'), rowInfo: $('#rowInfo'), total: $('#total'), totalHint: $('#totalHint'),
  link: $('#link'), pasteLink: $('#pasteLink'), submit: $('#submit'), status: $('#status'),
};

let current = null; // { fileName } — смета, которая сейчас в форме

// ---------- Статус ----------
function setStatus(text, kind = 'info') {
  els.status.textContent = text;
  els.status.className = `status ${kind}`;
}

// ---------- Экраны: пусто → форма → готово ----------
function showScreen(name) {
  els.empty.hidden = name !== 'empty';
  els.form.hidden = name !== 'form';
  els.done.hidden = name !== 'done';
}

function resetToEmpty() {
  current = null;
  els.file.value = '';
  showRow(null);
  setStatus('');
  showScreen('empty');
}
els.reset.addEventListener('click', resetToEmpty);
els.again.addEventListener('click', resetToEmpty);

// ---------- Чтение PDF ----------
async function readPdfTexts(file) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const texts = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const content = await (await doc.getPage(i)).getTextContent();
    for (const item of content.items) if (item.str) texts.push(item.str);
  }
  return texts;
}

async function handleFile(file) {
  if (!file) return;
  if (!/\.pdf$/i.test(file.name) && file.type !== 'application/pdf') {
    setStatus('Это не PDF', 'err');
    return;
  }
  setStatus('Читаю PDF…');
  let parsed;
  try {
    parsed = parseEstimate(await readPdfTexts(file), file.name);
  } catch (e) {
    console.error(e);
    setStatus(`Не получилось прочитать PDF: ${e.message}`, 'err');
    return;
  }

  current = { fileName: file.name };
  showScreen('form');
  els.fileName.textContent = file.name;
  els.task.value = parsed.task || '';
  els.total.value = parsed.total ? formatMoney(parsed.total.gross) : '';
  els.totalHint.textContent = parsed.total
    ? `без НДС ${formatMoney(parsed.total.net)} + НДС ${Math.round(parsed.total.rate * 100)}% ${formatMoney(parsed.total.vat)}`
    : 'Итог с НДС не нашёлся — впишите вручную.';
  els.totalHint.classList.toggle('bad', !parsed.total);
  els.link.value = '';

  const missing = [!parsed.task && 'номер задачи', !parsed.total && 'итог с НДС'].filter(Boolean);
  setStatus(missing.length ? `Не нашлось: ${missing.join(', ')}. Проверьте поля.` : '', missing.length ? 'err' : 'info');
  lookup();
  (parsed.task && parsed.total ? els.link : els.task).focus();
}

els.file.addEventListener('change', () => handleFile(els.file.files[0]));
['dragenter', 'dragover'].forEach((t) => els.drop.addEventListener(t, (e) => {
  e.preventDefault();
  els.drop.classList.add('over');
}));
['dragleave', 'drop'].forEach((t) => els.drop.addEventListener(t, () => els.drop.classList.remove('over')));
els.drop.addEventListener('drop', (e) => {
  e.preventDefault();
  handleFile(e.dataTransfer.files[0]);
});
// Чтобы случайно брошенный мимо файл не открывался вместо панели.
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => e.preventDefault());

// ---------- Связь с таблицей ----------
let lookupSeq = 0;
async function lookup() {
  const task = normTask(els.task.value);
  const seq = ++lookupSeq;
  if (!task) { showRow(null); return; }
  if (!isConfigured()) {
    showRow({ kind: 'warn', html: 'Настройки не заполнены — строку в таблице не проверить.' });
    return;
  }
  showRow({ kind: 'pending', html: `Ищу ${esc(task)} в таблице…` });
  try {
    const r = await api({ action: 'lookup', task });
    if (seq !== lookupSeq) return;
    if (!r.ok) throw new Error(r.error);
    if (!r.found) {
      showRow({ kind: 'err', html: `${esc(task)} не найдена в листе задач.` });
      return;
    }
    const filled = r.link || r.total;
    showRow({
      kind: filled ? 'warn' : '',
      html: `<b>${esc(r.subject || task)}</b><span class="meta">${esc([r.contractor, r.manager].filter(Boolean).join(' · '))}</span>` +
        (filled ? `<span class="note">Смета уже внесена: ${esc(r.total ? formatMoney(Number(r.total)) + ' ₽' : '—')}. При записи перезапишется.</span>` : ''),
    });
  } catch (e) {
    if (seq !== lookupSeq) return;
    showRow({ kind: 'err', html: esc(e.message) });
  }
}

function showRow(info) {
  els.rowInfo.hidden = !info;
  if (!info) return;
  els.rowInfo.className = `rowinfo ${info.kind}`;
  els.rowInfo.innerHTML = info.html;
}

let lookupTimer;
els.task.addEventListener('input', () => {
  clearTimeout(lookupTimer);
  lookupTimer = setTimeout(lookup, 500);
});

els.pasteLink.addEventListener('click', async () => {
  try {
    els.link.value = await readClipboard();
  } catch {
    setStatus('Нет доступа к буферу — вставьте ссылку вручную (Ctrl+V)', 'err');
  }
});

els.form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const task = normTask(els.task.value);
  const total = Number(els.total.value.replace(/[\s  ₽]/g, '').replace(',', '.'));
  const link = els.link.value.trim();
  if (!task) return setStatus('Укажите номер', 'err');
  if (!(total > 0)) return setStatus('Сумма не похожа на число', 'err');
  if (!/^https?:\/\//i.test(link)) return setStatus('Вставьте ссылку на смету с ВБ Диска', 'err');

  els.submit.disabled = true;
  setStatus('Записываю…');
  try {
    let r = await api({ action: 'write', task, total, link, fileName: current?.fileName || '' });
    if (!r.ok && r.error === 'exists') {
      const was = r.total ? `${formatMoney(Number(r.total))} ₽` : 'пусто';
      if (!confirm(`В ${task} уже есть смета (${was}). Перезаписать?`)) {
        setStatus('Отменено — в таблице ничего не менялось.');
        return;
      }
      r = await api({ action: 'write', task, total, link, fileName: current?.fileName || '', overwrite: true });
    }
    if (!r.ok) throw new Error(r.error);
    els.doneText.textContent = `${task} · ${formatMoney(total)} ₽ и ссылка в строке ${r.row}`;
    setStatus('');
    current = null;
    els.file.value = '';
    showScreen('done');
  } catch (err) {
    setStatus(`Не записалось: ${err.message}`, 'err');
  } finally {
    els.submit.disabled = false;
  }
});

function normTask(s) {
  return String(s).trim().replace(/\s+/g, ' ').toUpperCase();
}

/** Настройки поменялись — перепроверить строку, если смета открыта. */
export function onSettingsSaved() {
  if (els.task.value) lookup();
}
