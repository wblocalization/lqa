// Запуск: node --test smeta-extension/test/
// Опционально: SMETA_PDF=/путь/к/смете.pdf — прогнать разбор на настоящем PDF (нужен pdfjs-dist).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseEstimate, taskFromFileName, taskFromTexts, extractAmounts, findTotalWithVat, parseMoney, formatMoney, linesFromItems } from '../parser.js';

test('parseMoney понимает пробелы и запятую', () => {
  assert.equal(parseMoney('110 929,150'), 110929.15);
  assert.equal(parseMoney('5 937,50'), 5937.5);
});

test('номер задачи берётся из имени файла', () => {
  assert.equal(taskFromFileName('LIT-26-2203.pdf'), 'LIT-26-2203');
  assert.equal(taskFromFileName('LIT-26-2203-РВБ.pdf'), 'LIT-26-2203');
  assert.equal(taskFromFileName('HELLO-3123-2133132-RWB.pdf'), 'HELLO-3123-2133132');
  assert.equal(taskFromFileName('HELLO-3123-2133132-РВБ (1).pdf'), 'HELLO-3123-2133132');
  assert.equal(taskFromFileName('Lit-26-2203-рвб.pdf'), 'LIT-26-2203');
  assert.equal(taskFromFileName('РВБ.pdf'), null);
  assert.equal(taskFromFileName('LIT-26-2203 РВБ.pdf'), 'LIT-26-2203');
  assert.equal(taskFromFileName('LIT-26-2203_РВБ.pdf'), 'LIT-26-2203');
  assert.equal(taskFromFileName('LIT-26-2203 RWB-1.pdf'), 'LIT-26-2203');
  assert.equal(taskFromFileName('419cc3c8-LIT-26-2217-___.pdf'), 'LIT-26-2217');
  assert.equal(taskFromFileName('Смета GLB-5-100 финал.pdf'), 'GLB-5-100');
  assert.equal(taskFromFileName('lit_7.pdf'), 'LIT-7');
  assert.equal(taskFromFileName('смета от 22.09.pdf'), null);
});

test('если в имени файла номера нет — берётся номер сметы из текста, а не чужой код', () => {
  assert.equal(taskFromTexts(['WB-2063_AZ_AM_HY', '1LIT-26-2217']), 'LIT-26-2217');
  assert.equal(taskFromTexts(['WB-2063_AZ_AM_HY']), null);
  assert.equal(parseEstimate(['1LIT-26-2217'], 'LIT-30-1.pdf').task, 'LIT-30-1');
  assert.equal(parseEstimate(['1LIT-26-2217'], 'смета.pdf').task, 'LIT-26-2217');
});

test('итог с НДС находится по арифметике, а не по подписям', () => {
  // Порядок кусков как в настоящем PDF: сначала числа, подписи — отдельно.
  const texts = [
    '₽ 90 925,530', '₽ 20 003,620', '₽ 110 929,150',
    '₽ 5 937,500 Итого без скидок: ₽ 95 201,352', '₽ 84 988,031 ₽ 4 275,821',
    '₽ 6 250,000 ₽ 5 937,500', '₽ 9 484,050 ₽ 9 009,848',
    'Без НДС', 'НДС (22%)', 'С НДС',
  ];
  const r = parseEstimate(texts, 'LIT-26-2217 смета.pdf');
  assert.equal(r.task, 'LIT-26-2217');
  assert.deepEqual(r.total, { net: 90925.53, vat: 20003.62, gross: 110929.15, rate: 0.22, extraDigits: false });
});

test('НДС 20% и отсутствие итога', () => {
  assert.equal(findTotalWithVat(extractAmounts(['1 000,00', '200,00', '1 200,00'])).gross, 1200);
  assert.equal(findTotalWithVat(extractAmounts(['1 000,00', '300,00', '1 300,00'])), null);
});

test('настоящий PDF (если задан SMETA_PDF)', { skip: !process.env.SMETA_PDF }, async () => {
  const fs = await import('node:fs');
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(process.env.SMETA_PDF)) }).promise;
  const texts = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const content = await (await doc.getPage(i)).getTextContent();
    for (const it of content.items) texts.push(it.str);
  }
  const r = parseEstimate(texts, process.env.SMETA_PDF);
  console.log(r);
  assert.ok(r.task);
  assert.ok(r.total);
});

test('сумма — ровно как в смете: доли копейки не округляются', () => {
  // Обычная смета: «102 641,640» → 102641.64
  const r = parseEstimate(['₽ 84 132,490', '₽ 18 509,150', '₽ 102 641,640', 'Версия', '1'], 'LIT-26-2198-РВБ.pdf');
  assert.deepEqual(r.total, { net: 84132.49, vat: 18509.15, gross: 102641.64, rate: 0.22, extraDigits: false });
  assert.equal(r.version, 1);
  // Третий знак не ноль — отдаём как есть и помечаем
  const x = parseEstimate(['1 000,000', '220,005', '1 220,005', 'Версия 2'], 'x.pdf');
  assert.equal(x.total.gross, 1220.005);
  assert.equal(x.total.extraDigits, true);
  assert.equal(x.version, 2);
  assert.equal(formatMoney(1220.005), '1\u00a0220,005');
  assert.equal(formatMoney(102641.64), '102\u00a0641,64');
  // Без ошибок дробей: 1,005 × 100 в обычной арифметике даёт 100,4999…
  assert.deepEqual(extractAmounts(['1,005']), [101]);
});

test('версия сметы — по строке, даже если значение лежит в другом месте текстового слоя', () => {
  const items = [
    { str: 'Версия', x: 382, y: 471.2, page: 1 }, { str: 'Проект(ы)', x: 53, y: 465, page: 1 },
    { str: '25.09.2026', x: 302, y: 471.6, page: 1 }, { str: '1', x: 402, y: 471.6, page: 1 }, { str: 'Дата', x: 289, y: 471.2, page: 1 },
  ];
  assert.deepEqual(linesFromItems(items), ['Дата 25.09.2026 Версия 1', 'Проект(ы)']);
  assert.equal(parseEstimate(['Версия', 'Проект(ы)', '1'], 'x.pdf', linesFromItems(items)).version, 1);
});
