import * as pdfjs from './vendor/pdf.min.mjs';
import { parseEstimate, formatMoney, linesFromItems } from './parser.js';
import { settings, isConfigured, api, esc, readClipboard, DEFAULT_DISK_FOLDER } from './core.js';
import { uploadToDisk, listDiskPdfs, folderNameFromSubject } from './disk.js';

pdfjs.GlobalWorkerOptions.workerSrc = chrome.runtime.getURL('vendor/pdf.worker.min.mjs');

const $ = (s) => document.querySelector(s);
const els = {
  empty: $('#empty'), drop: $('#drop'), file: $('#file'), form: $('#form'), fileName: $('#fileName'),
  reset: $('#reset'), done: $('#done'), doneText: $('#doneText'), again: $('#again'),
  task: $('#task'), rowInfo: $('#rowInfo'), total: $('#total'), totalHint: $('#totalHint'),
  link: $('#link'), pasteLink: $('#pasteLink'), submit: $('#submit'), status: $('#status'), diskHint: $('#diskHint'),
  folder: $('#folder'), folderCard: $('#folderCard'), doneTitle: $('#doneTitle'), doneLinkBox: $('#doneLinkBox'),
  doneLink: $('#doneLink'), copyLink: $('#copyLink'), copyTotal: $('#copyTotal'),
  doneTotalBox: $('#doneTotalBox'), doneTotal: $('#doneTotal'), copyDoneTotal: $('#copyDoneTotal'),
};

// { file, fileName, subject, filled, oldTotal, version, found } — смета, которая сейчас в форме.
// found: true — строка в таблице есть; false — нет (или таблица недоступна): тогда только загрузка на диск.
let current = null;
let folderEdited = false; // название папки вписали руками — тему письма поверх не подставляем

/** Папка по умолчанию: тема письма из таблицы, а если задачи там нет — номер. */
function autoFolder() {
  return folderNameFromSubject(current && current.subject) || normTask(els.task.value);
}

/** Куда ляжет PDF на ВБ Диске: папка из настроек / название из поля «Папка на ВБ Диске». */
function diskFolders() {
  const base = String(settings.diskFolder || DEFAULT_DISK_FOLDER).split('/').map((x) => x.trim()).filter(Boolean);
  const name = folderNameFromSubject(els.folder.value) || autoFolder();
  return name ? base.concat([name]) : base;
}

/** Задачи в таблице нет (или таблица не настроена) — можно только загрузить PDF на диск и скопировать ссылку. */
function diskOnly() {
  return Boolean(current) && (!isConfigured() || current.found === false || !normTask(els.task.value));
}

/** Ссылка вписана — просто записываем; пусто — сначала загрузим PDF на диск. */
function updateDiskHint() {
  const manual = Boolean(els.link.value.trim());
  if (current && !folderEdited) els.folder.value = autoFolder();
  els.submit.textContent = manual ? 'Записать в таблицу' : diskOnly() ? 'Только загрузить на ВБ Диск' : 'Загрузить на ВБ Диск и записать';
  els.folderCard.hidden = manual || !current;
  if (current && !manual) {
    els.diskHint.textContent = `PDF ляжет в «${diskFolders().join(' / ')}»` +
      (diskOnly() ? '. В таблицу ничего не запишется — ссылку можно будет скопировать.' : '');
  }
}
els.link.addEventListener('input', updateDiskHint);
els.folder.addEventListener('input', () => {
  folderEdited = Boolean(els.folder.value.trim());
  updateDiskHint();
});

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
  const texts = [], items = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const content = await (await doc.getPage(i)).getTextContent();
    for (const item of content.items) {
      if (!item.str) continue;
      texts.push(item.str);
      items.push({ str: item.str, x: item.transform[4], y: item.transform[5], page: i });
    }
  }
  return { texts, lines: linesFromItems(items) };
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
    const pdf = await readPdfTexts(file);
    parsed = parseEstimate(pdf.texts, file.name, pdf.lines);
  } catch (e) {
    console.error(e);
    setStatus(`Не получилось прочитать PDF: ${e.message}`, 'err');
    return;
  }

  current = { file, fileName: file.name, subject: '', filled: false, oldTotal: '', version: parsed.version, found: null };
  folderEdited = false;
  els.folder.value = '';
  showScreen('form');
  els.fileName.textContent = file.name + (parsed.version ? ` · версия ${parsed.version}` : '');
  els.task.value = parsed.task || '';
  els.total.value = parsed.total ? formatMoney(parsed.total.gross) : '';
  const t = parsed.total;
  els.totalHint.textContent = t
    ? `без НДС ${formatMoney(t.net)} + НДС ${Math.round(t.rate * 100)}% ${formatMoney(t.vat)}` +
      (t.extraDigits ? ' · ⚠ в смете доли копейки — записываю как есть, без округления. Сверьте со счётом подрядчика' : '')
    : 'Итог с НДС не нашёлся — впишите вручную.';
  els.totalHint.classList.toggle('bad', !t || t.extraDigits);
  els.link.value = '';
  updateDiskHint();

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
  if (current) { current.found = null; current.subject = ''; current.filled = false; current.oldTotal = ''; updateDiskHint(); }
  if (!task) { showRow(null); return; }
  if (!isConfigured()) {
    showRow({ kind: 'warn', html: 'Настройки не заполнены — строку в таблице не проверить. PDF можно просто загрузить на ВБ Диск.' });
    return;
  }
  showRow({ kind: 'pending', html: `Ищу ${esc(task)} в таблице…` });
  try {
    const r = await api({ action: 'lookup', task });
    if (seq !== lookupSeq) return;
    if (!r.ok) throw new Error(r.error);
    if (!r.found) {
      if (current) { current.found = false; updateDiskHint(); }
      showRow({ kind: 'err', html: `${esc(task)} не найдена в листе задач.<span class="meta">PDF можно просто загрузить на ВБ Диск и скопировать ссылку.</span>` });
      return;
    }
    const filled = r.link || r.total;
    if (current) {
      current.found = true;
      current.subject = r.subject || '';
      current.filled = Boolean(filled);
      current.oldTotal = r.total;
      updateDiskHint();
    }
    showRow({
      kind: filled ? 'warn' : '',
      html: `<b>${esc(r.subject || task)}</b><span class="meta">${esc([r.contractor, r.manager].filter(Boolean).join(' · '))}</span>` +
        (filled ? `<span class="note">Смета уже внесена: ${esc(r.total ? formatMoney(Number(r.total)) + ' ₽' : '—')}. При записи перезапишется.</span>` : ''),
    });
  } catch (e) {
    if (seq !== lookupSeq) return;
    if (current) { current.found = false; updateDiskHint(); }
    showRow({ kind: 'err', html: `${esc(e.message)}<span class="meta">PDF можно просто загрузить на ВБ Диск и скопировать ссылку.</span>` });
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
    updateDiskHint();
  } catch {
    setStatus('Нет доступа к буферу — вставьте ссылку вручную (Ctrl+V)', 'err');
  }
});

els.form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const task = normTask(els.task.value);
  const total = parseTotal(els.total.value);
  let link = els.link.value.trim();
  if (diskOnly() && !link) return uploadOnly(task);
  if (!task) return setStatus('Укажите номер', 'err');
  if (!(total > 0)) return setStatus('Сумма не похожа на число', 'err');
  if (link && !/^https?:\/\//i.test(link)) return setStatus('Ссылка должна начинаться с https://', 'err');
  if (!link && !current) return setStatus('Перетащите PDF сметы или вставьте ссылку', 'err');
  if (current && current.found === false) return setStatus(`${task} нет в таблице — записывать некуда. Сотрите ссылку, чтобы просто загрузить PDF на диск.`, 'err');

  // Смета уже есть — спрашиваем до загрузки, чтобы не грузить зря, и показываем, как меняется сумма
  let overwrite = false;
  if (current && current.filled) {
    if (!confirm(replaceQuestion(task, current.oldTotal, total))) return setStatus('Отменено — ничего не менялось.');
    overwrite = true;
  }
  const version = current?.version || '';

  els.submit.disabled = true;
  try {
    let where = '', uploaded = false;
    if (!link) {
      const up = await uploadPdf();
      uploaded = true;
      link = up.link;
      where = ` · ${up.name} в «${up.folders.join(' / ')}»`;
      els.link.value = link;
      updateDiskHint();
    }
    setStatus('Записываю в таблицу…');
    const fileName = current?.fileName || '';
    let r = await api({ action: 'write', task, total, link, fileName, version, overwrite });
    if (!r.ok && r.error === 'exists') {
      if (!confirm(replaceQuestion(task, r.total, total))) {
        setStatus('Отменено — в таблице ничего не менялось.');
        return;
      }
      r = await api({ action: 'write', task, total, link, fileName, version, overwrite: true });
    }
    if (!r.ok && uploaded) {
      // PDF уже на диске — ссылку не теряем, даже если в таблицу записать не вышло
      showDone('Загружено на ВБ Диск', `В таблицу не записано: ${r.error}${where}`, link);
      return;
    }
    if (!r.ok) throw new Error(r.error);
    const was = r.was !== '' && r.was != null && Number.isFinite(Number(r.was)) && Number(r.was) !== total
      ? ` (было ${formatMoney(Number(r.was))} ₽ — осталось в заметке к ячейке)` : '';
    showDone('Записано', `${task} · ${formatMoney(total)} ₽${was} и ссылка в строке ${r.row}${where}`, link);
  } catch (err) {
    setStatus(`Не получилось: ${err.message}`, 'err');
  } finally {
    els.submit.disabled = false;
  }
});

/** PDF на ВБ Диск в папку из поля. В папке уже есть смета — спросим: заменить или оставить обе. */
async function uploadPdf() {
  const folders = diskFolders();
  setStatus('Смотрю папку на ВБ Диске…');
  const existing = await listDiskPdfs(folders);
  let remove = [], keepBoth = false;
  if (existing.length) {
    const replace = confirm(`В папке на ВБ Диске уже есть: ${existing.join(', ')}.\n\n` +
      'ОК — заменить новой сметой (старая уйдёт в корзину ВБ Диска, её можно восстановить)\n' +
      'Отмена — оставить и старую, и новую');
    if (replace) remove = existing; else keepBoth = true;
  }
  setStatus(`Загружаю на ВБ Диск: ${folders.join(' / ')}…`);
  const up = await uploadToDisk(current.file, folders, { remove, keepBoth });
  return { ...up, folders };
}

/** Задачи в таблице нет — только загрузка на диск, ссылку покажем, чтобы скопировать. */
async function uploadOnly(task) {
  if (!current) return setStatus('Перетащите PDF сметы', 'err');
  if (!folderNameFromSubject(els.folder.value) && !autoFolder()) return setStatus('Впишите название папки на ВБ Диске', 'err');
  els.submit.disabled = true;
  try {
    const up = await uploadPdf();
    const why = !isConfigured() ? 'настройки таблицы не заполнены' : task ? `${task} нет в таблице` : 'номер задачи не указан';
    showDone('Загружено на ВБ Диск', `${up.name} в «${up.folders.join(' / ')}». В таблицу не записано: ${why}.`, up.link);
  } catch (err) {
    setStatus(`Не получилось: ${err.message}`, 'err');
  } finally {
    els.submit.disabled = false;
  }
}

function showDone(title, text, link) {
  const total = parseTotal(els.total.value);
  els.doneTitle.textContent = title;
  els.doneText.textContent = text;
  els.doneLink.value = link || '';
  els.doneLinkBox.hidden = !link;
  els.doneTotal.value = total > 0 ? formatMoney(total) : '';
  els.doneTotalBox.hidden = !(total > 0);
  els.copyLink.textContent = 'Скопировать ссылку';
  els.copyDoneTotal.textContent = 'Скопировать сумму';
  setStatus('');
  current = null;
  els.file.value = '';
  showScreen('done');
}

/** «102 641,64» → 102641.64 */
function parseTotal(s) {
  return Number(String(s).replace(/[\s  ₽]/g, '').replace(',', '.'));
}

/** Сумма для вставки в Excel / Google Таблицы: «102641,64» — без пробелов и ₽, запятая перед копейками. */
function totalForPaste(x) {
  return formatMoney(x).replace(/[\s  ]/g, '');
}

async function copy(text, btn, field) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    if (!field) return;
    field.select();
    document.execCommand('copy');
  }
  btn.dataset.label = btn.dataset.label || btn.textContent;
  btn.textContent = 'Скопировано ✓';
  clearTimeout(btn.copyTimer);
  btn.copyTimer = setTimeout(() => { btn.textContent = btn.dataset.label; }, 1500);
}

els.copyLink.addEventListener('click', () => copy(els.doneLink.value, els.copyLink, els.doneLink));
els.copyDoneTotal.addEventListener('click', () => copy(totalForPaste(parseTotal(els.doneTotal.value)), els.copyDoneTotal));
els.copyTotal.addEventListener('click', () => {
  const total = parseTotal(els.total.value);
  if (!(total > 0)) return setStatus('Суммы нет — впишите её в поле', 'err');
  copy(totalForPaste(total), els.copyTotal);
});


/** «Было 102 641,64 ₽ → станет 105 210,30 ₽ (+2 568,66 ₽)» — чтобы обновлённую смету было видно сразу. */
function replaceQuestion(task, oldTotal, newTotal) {
  const old = Number(oldTotal);
  let text = `В ${task} уже есть смета.\n\n`;
  if (oldTotal !== '' && oldTotal != null && Number.isFinite(old)) {
    const diff = Math.round((newTotal - old) * 1000) / 1000;
    text += `Было: ${formatMoney(old)} ₽\nСтанет: ${formatMoney(newTotal)} ₽\n` +
      (diff ? `Разница: ${diff > 0 ? '+' : '−'}${formatMoney(Math.abs(diff))} ₽` : 'Сумма не меняется') + '\n\n';
  } else {
    text += `Станет: ${formatMoney(newTotal)} ₽\n\n`;
  }
  return text + 'Заменить ссылку и сумму? Прежняя сумма останется в заметке к ячейке.';
}

function normTask(s) {
  return String(s).trim().replace(/\s+/g, ' ').toUpperCase();
}

/** Настройки поменялись — перепроверить строку, если смета открыта. */
export function onSettingsSaved() {
  if (els.task.value) lookup();
}
