// Запуск: node --test band-tracker-extension/test/
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseOrder, parseLanguages, subtaskTitle, orderSummary, trackerKeys } from '../parse.js';

const ORDER = `Коллеги, добрый день!

Прошу взять в работу переводы для [Tracker](https://tracker.wb.ru).
(Сейчас используется машинный перевод, его нужно проверить, утвердить или внести правки).

**Языки:** английский (en), грузинский, казахский
**Ответственный за перевод:** @loc-managers
**Продукт:** Tracker
**Веблейт:** https://weblate.wb.ru/projects/tracker/tracker-frontend/ru/.
**Макеты:** [Фигма](https://figma.com/file/abc)`;

test('заказ из треда: продукт, языки, страны, название по сути', () => {
  const o = parseOrder(ORDER);
  assert.equal(o.product, 'Tracker');
  assert.deepEqual(o.languages.map((l) => l.code), ['en', 'ka', 'kk']);
  assert.equal(o.title, 'Tracker: переводы — en, ka, kk');
  assert.ok(o.links.some((l) => l.label === 'Фигма'));
  assert.equal(subtaskTitle(o, o.languages[1]), 'Tracker: перевод — ka');
  assert.match(orderSummary(o), /Страны:\*\* Международный, Грузия, Казахстан/);
});

test('языки кодами и с суффиксами, без повторов', () => {
  assert.deepEqual(parseLanguages('kk_cyr, ky; uz и en, Kazakh').map((l) => l.code), ['kk', 'ky', 'uz', 'en']);
});

test('без полей — первая содержательная строка, без приветствия', () => {
  const o = parseOrder('Привет!\nПрошу перевести баннер для акции «Осень»');
  assert.equal(o.title, 'Перевести баннер для акции «Осень»');
});

test('номера задач трекера из ответа в треде', () => {
  assert.deepEqual(trackerKeys('Завели: [LOC-101](https://tracker.wb.ru/i/EDITORS/LOC-101), https://tracker.wb.ru/i/EDITORS/LOC-102'), ['LOC-101', 'LOC-102']);
});

test('«перевести» — тоже переводы', () => {
  assert.equal(parseOrder('**Продукт:** Магазин\n**Языки:** uz\nПрошу перевести пуши').title, 'Магазин: переводы — uz');
});

test('дедлайн из заказа — как его хранит трекер', async () => {
  const { parseDue } = await import('../parse.js');
  const now = new Date(2026, 9, 14);
  assert.equal(parseDue('28.10', now), '2026-10-28T00:00:00+03:00');
  assert.equal(parseDue('до 3.01', now), '2027-01-03T00:00:00+03:00');
  assert.equal(parseDue('28.10.26', now), '2026-10-28T00:00:00+03:00');
  assert.equal(parseDue('нет', now), '');
  assert.equal(parseOrder('**Продукт:** X\n**Дедлайн:** 28.10.2026').due, '2026-10-28T00:00:00+03:00');
});

test('поля одной строкой — разносятся, название не «Языки: …»', () => {
  const o = parseOrder('Коллеги, привет!\n**Языки:** узбекский, кыргызский, армянский **Ответственный за перевод:** @loc-managers');
  assert.deepEqual(o.languages.map((l) => l.code), ['uz', 'ky', 'hy']);
  assert.equal(o.title, 'Локализация — uz, ky, hy');
  const p = parseOrder('Прошу перевести баннеры. Языки: узбекский, армянский Продукт: Магазин');
  assert.equal(p.product, 'Магазин');
  assert.equal(p.title, 'Магазин: переводы — uz, hy');
});
