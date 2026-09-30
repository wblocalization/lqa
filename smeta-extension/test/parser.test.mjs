// Запуск: node --test smeta-extension/test/
// Опционально: SMETA_PDF=/путь/к/смете.pdf — прогнать разбор на настоящем PDF (нужен pdfjs-dist).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseEstimate, findTaskId, extractAmounts, findTotalWithVat, parseMoney } from '../parser.js';

test('parseMoney понимает пробелы и запятую', () => {
  assert.equal(parseMoney('110 929,150'), 110929.15);
  assert.equal(parseMoney('5 937,50'), 5937.5);
});

test('номер задачи из номера сметы и из имени файла', () => {
  assert.equal(findTaskId(['1LIT-26-2217']), 'LIT-26');
  assert.equal(findTaskId(['WB-2063_AZ_AM', 'нет номера']), null);
  assert.equal(findTaskId(['lit_7 смета']), 'LIT-7');
  assert.equal(findTaskId(['GLB-12-1'], ['LIT', 'GLB']), 'GLB-12');
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
  assert.equal(r.task, 'LIT-26');
  assert.deepEqual(r.total, { net: 90925.53, vat: 20003.62, gross: 110929.15, rate: 0.22 });
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
