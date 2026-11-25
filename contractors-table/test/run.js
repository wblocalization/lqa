// Проверка Расширение.gs без Google: node contractors-table/test/run.js
// FIXTURE=/путь/fixture.json — дополнительно прогнать на копии настоящего листа (в репозиторий её не кладём).
'use strict';
const fs = require('fs'), path = require('path'), vm = require('vm');
const { makeContext, mskDate, formatDate, TABLE_TZ } = require('./mock');

const SRC = fs.readFileSync(path.join(__dirname, '..', 'Расширение.gs'), 'utf8')
  .replace("const TABLE_LINK = '';", "const TABLE_LINK = 'https://docs.google.com/spreadsheets/d/OLDTABLEID0123456789abcdefghij/edit#gid=1';");
if (!SRC.includes('OLDTABLEID')) throw new Error('не нашла TABLE_LINK в скрипте');

const NOW = '2026-11-12T07:00:00Z'; // 10:00 по Москве
let passed = 0, failed = 0;
function check(name, fn) {
  try { fn(); passed++; console.log('  ✓ ' + name); } catch (e) { failed++; console.log('  ✗ ' + name + '\n      ' + String(e.message).split('\n').join('\n      ')); }
}
function eq(a, b, what) {
  const x = JSON.stringify(a), y = JSON.stringify(b);
  if (x !== y) throw new Error((what ? what + ': ' : '') + 'ждала ' + y + ', получила ' + x);
}
function ok(v, what) { if (!v) throw new Error(what || 'ожидалось истинное'); }

function load(fixture, patch) {
  const env = makeContext(JSON.parse(JSON.stringify(fixture)), { now: NOW });
  vm.createContext(env.ctx);
  vm.runInContext(patch ? patch(SRC) : SRC, env.ctx, { filename: 'Расширение.gs' });
  env.post = (req, token) => { // ms — время работы скрипта, в сравнениях не нужно
    const res = JSON.parse(env.ctx.doPost({ postData: { contents: JSON.stringify(Object.assign({ token: token || 'T', user: 'Тест' }, req)) } }));
    delete res.ms;
    return res;
  };
  env.calls = fn => { const before = env.sheet.calls; fn(); return env.sheet.calls - before; }; // сколько раз скрипт обратился к таблице
  env.day = (r, c) => { const v = env.sheet.peek(r, c).v; return v instanceof Date ? formatDate(v, TABLE_TZ, 'yyyy-MM-dd HH:mm') : v; };
  env.val = (r, c) => env.sheet.peek(r, c).v;
  env.row = r => { const out = []; for (let c = 1; c <= 19; c++) out.push(env.val(r, c)); return out; };
  env.sets = () => env.sheet.ops.filter(o => o[0] === 'set' || o[0] === 'formula').map(o => `R${o[1]}C${o[2]}`);
  env.reset = () => { env.sheet.ops.length = 0; };
  return env;
}

// ---------- Маленький лист, устроенный как настоящий ----------
const D = iso => ({ __date: iso });
const HEAD = ['А', 'Сторона', 'Дата', 'Тема письма в аутлуке', 'Задача', 'Тема в бенде (ссылка)', 'Ник заказчика в бенде', 'Языки',
  'Дедлайн (если есть)', 'Дедлайн', 'Статус', 'Смета (ссылка)', 'Итог(рубли,сНДС)', 'Подрядчик', 'Менеджер',
  'Статус отдачи задачи менеджером заказчику', 'SP', 'Дополнительный комментарий', 'Жалобы на заказчика'];
const B1 = 'https://band.wb.ru/wb/pl/aaa', B2 = 'https://band.wb.ru/wb/pl/bbb', B3 = 'https://band.wb.ru/wb/pl/ccc';
const SYN = {
  rows: [
    HEAD,
    /* 2 */ ['LIT-25-1', 'WBP', D('2025-06-05'), '[logrusit] [ka] старая', '', B1, '@a', 'Грузинский', '1-2 дня', D('2025-06-06'), 'Отдано', '', 100, 'LogrusIT', 'Анна Первая', 'Отдано в срок', '', '', ''],
    /* 3 */ ['LIT-26-10', 'Магазинка', D('2026-11-02'), '[LIT-26-10][LogrusIT][ka][Магазинка] Баннер', 'LOCAL-100', B2, '@b', 'Грузинский', '1-2 дня', D('2026-11-03'), 'Отдано', '', '', 'LogrusIT', 'Анна Первая', 'Отдано в срок', 2, '', ''],
    /* 4 */ ['LIT-26-11', 'Магазинка', D('2026-11-10'), '[LIT-26-11][LogrusIT][ka][az][Магазинка] Пуши', 'LOCAL-101', B3, '@c', 'Грузинский', 'До недели', D('2026-11-20'), 'В работе', '', '', 'LogrusIT', 'Анна Первая', '', '', '', ''],
    /* 5 */ ['', 'Межнар', D('2026-11-11'), '', 'LOCAL-102', 'https://band.wb.ru/wb/pl/ddd\nhttps://band.wb.ru/wb/pl/eee', '@d', 'Азербайджанский', 'До недели', D('2026-11-19'), 'В работе', '', '', '', 'Анна Первая', '', '', '', ''],
    /* 6 */ ['', 'Межнар', D('2026-11-11'), 'Перевод заявки (штат)', '', 'https://band.wb.ru/wb/pl/fff', '@e', 'ШТАТ армянский', '1-2 дня', D('2026-11-12'), 'В работе', '', '', '', 'Борис Второй', '', '', '', ''],
    /* 7 */ ['LoG-26-3', 'WBP', D('2026-11-11'), '[LoG-26-3][LogrusGlobal][hy][WBP] Тексты', 'LOCAL-103', 'https://band.wb.ru/wb/pl/ggg', '@f', 'Армянский', 'До недели', D('2026-11-17'), 'В работе', '', '', 'LogrusGlobal', 'Борис Второй', '', '', '', ''],
    /* 8 */ ['BP-26-7', 'WBP', D('2026-11-11'), '[BP-26-7][Бюро][ka][WBP] Тексты', '', '', '', 'Грузинский', 'До недели', D('2026-11-17'), 'В работе', '', '', 'Бюро переводов', 'Борис Второй', '', '', '', ''],
    /* 9 */ ['LIT-26-12', 'ПВЗ', D('2026-11-09'), '[LIT-26-12][LogrusIT][uz][ПВЗ] Озвучка', '', '', '@g', 'Узбекский', '1-2 дня', D('2026-11-11'), 'Принято', 'https://disk.wb.ru/f/1', 1000, 'LogrusIT', 'Анна Первая', '', '', 'коммент', '']
  ],
  links: { '9:12': 'https://disk.wb.ru/f/1' },
  merges: [[4, 1, 2, 1], [4, 4, 2, 1], [4, 12, 2, 1], [7, 5, 2, 1], [7, 6, 2, 1], [7, 7, 2, 1]],
  lists: {
    2: ['Сторона', 'Межнар', 'WBP', 'Магазинка', 'ПВЗ'],
    8: ['Языки', 'Азербайджанский', 'Армянский', 'Грузинский', 'Узбекский', 'Турецкий', 'ШТАТ армянский', 'ШТАТ en,hy,ka'],
    9: ['Дедлайн (если есть)', 'ASAP', '1-2 дня', 'До недели', 'Больше недели', 'Месяц и больше'],
    11: ['Статус', 'Принято', 'В работе', 'Отдано', 'Отменено', 'Холд'],
    14: ['LogrusIT', 'LogrusGlobal', 'Бюро переводов', 'JanusWW', 'Awatera'],
    15: ['Менеджер', 'Анна Первая', 'Борис Второй'],
    16: ['Отдано в срок', 'Сообщили о просрочке', 'Не сообщили о просрочке']
  },
  multi: [8], maxRows: 60
};

const NEW_TASK = {
  contractor: 'LogrusIT', subject: 'Новая акция', product: 'Магазинка', ticket: '555', link: 'https://band.wb.ru/wb/pl/new1',
  link2: 'https://band.wb.ru/wb/pl/new2', date: '2026-11-12', exactDeadline: '2026-11-16', customer: '@new', status: 'В работе',
  manager: 'Анна Первая', languages: ['Грузинский', 'Узбекский'], comment: 'к', sp: '3'
};

console.log('Маленький лист');
check('неверный токен — отказ, лист не открывается', () => {
  const e = load(SYN);
  eq(e.post({ action: 'taskForm' }, 'X'), { ok: false, error: 'Неверный токен' });
  ok(!e.ctx.__opened, 'таблицу открыли до проверки токена');
});

check('справочники — из проверки данных ячеек, без названия колонки и без пункта с запятыми', () => {
  const e = load(SYN);
  const l = e.post({ action: 'taskForm' }).lists;
  eq(l.contractors, ['LogrusIT', 'LogrusGlobal', 'Бюро переводов', 'JanusWW', 'Awatera']);
  eq(l.managers, ['Анна Первая', 'Борис Второй']);
  eq(l.products, ['Межнар', 'WBP', 'Магазинка', 'ПВЗ']);
  eq(l.statuses, ['Принято', 'В работе', 'Отдано', 'Отменено', 'Холд']);
  eq(l.languages, { regular: ['Азербайджанский', 'Армянский', 'Грузинский', 'Узбекский'], shtat: ['ШТАТ армянский'], rare: ['Турецкий'] });
  eq(l.features, { reports: false, mail: false });
  ok(/OLDTABLEID.*#gid=\d+$/.test(l.tableUrl), l.tableUrl);
  eq(e.ctx.__opened, 'OLDTABLEID0123456789abcdefghij');
  eq(e.sheet.ops, [], 'чтение ничего не меняет');
});

check('следующие номера: самый большой за этот год + 1 (и по теме, если в A номера ещё нет)', () => {
  const e = load(SYN);
  eq(e.post({ action: 'previewTaskIds' }).ids, { LogrusIT: 'LIT-26-13', LogrusGlobal: 'LoG-26-4', 'Бюро переводов': 'BP-26-8', JanusWW: 'JWW-26-1', Awatera: 'AWT-26-1' });
  const fx = JSON.parse(JSON.stringify(SYN));
  fx.rows.push(['', 'WBP', D('2026-11-12'), '[LIT-26-40][LogrusIT][ka] тему вписали, номер ещё нет', '', '', '', '', '', '', '', '', '', '', 'Анна Первая', '', '', '', '']);
  eq(load(fx).post({ action: 'previewTaskId', contractor: 'LogrusIT' }).id, 'LIT-26-41');
});

check('новая задача: строка вставляется под последней, номер, тема, даты по Москве, ссылки', () => {
  const e = load(SYN);
  const r = e.post({ action: 'addTask', task: NEW_TASK });
  eq([r.ok, r.id, r.row], [true, 'LIT-26-13', 10]);
  ok(/range=A10$/.test(r.url), r.url);
  eq(e.sheet.ops[0], ['insertRow', 10], 'сначала вставка строки, а не запись в пустую');
  eq(e.sheet.maxRows, 61);
  eq(e.row(10), ['LIT-26-13', 'Магазинка', e.val(10, 3), '[LIT-26-13][LogrusIT][ka][uz][Магазинка] Новая акция', 'LOCAL-555',
    'https://band.wb.ru/wb/pl/new1\nhttps://band.wb.ru/wb/pl/new2', '@new', 'Грузинский, Узбекский', '1-2 дня', e.val(10, 10), 'В работе', '', '',
    'LogrusIT', 'Анна Первая', '', 3, 'к', '']);
  eq([e.day(10, 3), e.day(10, 10)], ['2026-11-12 00:00', '2026-11-16 00:00'], 'даты — полночь по часам таблицы');
  eq(e.sheet.peek(10, 6).rt.getRuns().map(x => x.getLinkUrl()).filter(Boolean), ['https://band.wb.ru/wb/pl/new1', 'https://band.wb.ru/wb/pl/new2']);
  ok(e.sheet.ops.some(o => o[0] === 'copy' && o[2] === 11), 'оформление — с пустой строки-заготовки');
  eq(e.sets().filter(x => !/^R10C/.test(x)), [], 'другие строки не тронуты');
  // следующая задача получает следующий номер
  eq(e.post({ action: 'addTask', task: NEW_TASK }).id, 'LIT-26-14');
  eq(e.post({ action: 'addTask', task: Object.assign({}, NEW_TASK, { contractor: 'Бюро переводов' }) }).id, 'BP-26-8');
});

check('новая задача: значение не из выпадающего списка — строка не остаётся, понятная ошибка', () => {
  const e = load(SYN);
  const r = e.post({ action: 'addTask', task: Object.assign({}, NEW_TASK, { status: 'Чужой статус' }) });
  eq(r.ok, false);
  ok(/Таблица не приняла строку/.test(r.error), r.error);
  eq(e.sheet.maxRows, 60, 'вставленная строка убрана');
  eq(e.row(10).join(''), '');
});

check('колонки сдвинулись — запись останавливается, чтение работает', () => {
  const fx = JSON.parse(JSON.stringify(SYN));
  fx.rows[0][10] = 'Что-то новое';
  const e = load(fx);
  const r = e.post({ action: 'addTask', task: NEW_TASK });
  eq(r.ok, false);
  ok(/изменились колонки/.test(r.error), r.error);
  eq(e.sheet.ops, []);
  eq(e.post({ action: 'setStatus', row: 3, id: 'LIT-26-10', status: 'Отдано' }).ok, false);
});

check('«Мои»: строка-продолжение получает номер и тему заказа; просрочка и счётчики', () => {
  const e = load(SYN);
  const r = e.post({ action: 'myTasks', manager: 'Анна Первая' });
  eq(r.tasks.map(t => [t.row, t.id, t.title, t.status, t.overdue]), [
    [9, 'LIT-26-12', 'Озвучка', 'Принято', true], [5, 'LIT-26-11', 'Пуши', 'В работе', false], [4, 'LIT-26-11', 'Пуши', 'В работе', false]]);
  eq(r.counts, { open: 3, done: 2, cancelled: 0, all: 5 });
  eq(r.overdue, 1);
  eq(r.tasks[1].link, 'https://band.wb.ru/wb/pl/ddd');
  eq(e.post({ action: 'myTasks', manager: 'Борис Второй' }).tasks.map(t => [t.row, t.id, t.dueToday]), [[6, '', true], [7, 'LoG-26-3', false], [8, 'BP-26-7', false]]);
  eq(e.post({ action: 'myTasks', manager: 'Анна Первая', filter: 'done' }).tasks.map(t => t.row), [3, 2], 'закрытые — свежие сверху');
  eq(e.post({ action: 'myTasks', manager: '*' }).managers, ['Анна Первая', 'Борис Второй']);
});

check('задача для правки: объединённые «Задача», Band и ник берутся с верхней строки', () => {
  const e = load(SYN);
  const t = e.post({ action: 'getTask', row: 8, id: 'BP-26-7', origSubject: '[BP-26-7][Бюро][ka][WBP] Тексты' }).task;
  eq([t.ticket, t.link, t.customer, t.contractor, t.date, t.exactDeadline], ['LOCAL-103', 'https://band.wb.ru/wb/pl/ggg', '@f', 'Бюро переводов', '2026-11-11', '2026-11-17']);
  const c = e.post({ action: 'getTask', row: 5, id: 'LIT-26-11', origSubject: '[LIT-26-11][LogrusIT][ka][az][Магазинка] Пуши' }).task;
  eq([c.row, c.product, c.ticket, c.link, c.link2, c.languages], [5, 'Межнар', 'LOCAL-102', 'https://band.wb.ru/wb/pl/ddd', 'https://band.wb.ru/wb/pl/eee', 'Азербайджанский']);
  eq(e.post({ action: 'getTask', row: 9, id: 'LIT-26-12' }).task.estimateLink, 'https://disk.wb.ru/f/1');
});

check('правка без изменений не трогает ни одной ячейки; смена статуса — ровно одну', () => {
  const e = load(SYN);
  for (const [row, id] of [[3, 'LIT-26-10'], [4, 'LIT-26-11'], [5, 'LIT-26-11'], [6, ''], [8, 'BP-26-7'], [9, 'LIT-26-12']]) {
    const t = e.post({ action: 'getTask', row, id, origSubject: id ? undefined : 'Перевод заявки (штат)' }).task;
    const r = e.post({ action: 'saveTask', task: t });
    eq([r.ok, r.changed, e.sheet.ops.length], [true, 0, 0], 'строка ' + row);
  }
  const t = e.post({ action: 'getTask', row: 5, id: 'LIT-26-11' }).task;
  t.status = 'Отдано'; t.deliveryStatus = 'Отдано в срок';
  eq(e.post({ action: 'saveTask', task: t }).changed, 2);
  eq(e.sets(), ['R5C11', 'R5C16']);
});

check('правка строки-продолжения: тема и смета пишутся в верхнюю ячейку объединения', () => {
  const e = load(SYN);
  const t = e.post({ action: 'getTask', row: 5, id: 'LIT-26-11' }).task;
  t.subject = '[LIT-26-11][LogrusIT][ka][az][Магазинка] Пуши и баннеры'; t.estimateLink = 'https://disk.wb.ru/f/2'; t.total = '3000+2500,5';
  t.exactDeadline = '2026-11-25'; t.deadline = ''; t.languages = 'Азербайджанский, Турецкий';
  eq(e.post({ action: 'saveTask', task: t }).ok, true);
  eq(e.sets().sort(), ['R4C12', 'R4C4', 'R5C10', 'R5C13', 'R5C8', 'R5C9'].sort());
  eq([e.val(4, 4), e.val(4, 12), e.day(5, 10), e.val(5, 9), e.val(5, 8)],
    ['[LIT-26-11][LogrusIT][ka][az][Магазинка] Пуши и баннеры', 'https://disk.wb.ru/f/2', '2026-11-25 00:00', 'Больше недели', 'Азербайджанский, Турецкий']);
  eq(e.sheet.peek(5, 13).formula, '=3000+2500.5');
});

check('статус и статус отдачи из «Мои» — по одной ячейке; строка сдвинулась — находит по номеру', () => {
  const e = load(SYN);
  eq(e.post({ action: 'setStatus', row: 3, id: 'LIT-26-12', origSubject: '[LIT-26-12][LogrusIT][uz][ПВЗ] Озвучка', status: 'Отдано' }), { ok: true, row: 9 });
  eq(e.post({ action: 'setDelivery', row: 9, id: 'LIT-26-12', value: 'Сообщили о просрочке' }).ok, true);
  eq(e.sets(), ['R9C11', 'R9C16']);
  eq(e.post({ action: 'setStatus', row: 9, id: 'LIT-26-999', status: 'Отдано' }).ok, false);
});

check('ещё строка в заказ: под его строками, номер и тема объединяются, автономер таблицы не сбивается', () => {
  const e = load(SYN);
  const r = e.post({ action: 'addTaskRow', id: 'LIT-26-11', task: { product: 'ПВЗ', link: 'https://band.wb.ru/wb/pl/row3', customer: '@h',
    languages: ['Узбекский'], exactDeadline: '2026-11-18', status: 'Принято' } });
  eq([r.ok, r.id, r.row, r.subject], [true, 'LIT-26-11', 6, '[LIT-26-11][LogrusIT][ka][az][uz][Магазинка] Пуши']);
  eq(e.row(6).slice(0, 2).concat(e.row(6).slice(3, 9)), ['', 'ПВЗ', '', 'LOCAL-101', 'https://band.wb.ru/wb/pl/row3', '@h', 'Узбекский', 'Больше недели']);
  eq([e.day(6, 3), e.val(6, 14), e.val(6, 15), e.val(6, 11)], ['2026-11-10 00:00', 'LogrusIT', 'Анна Первая', 'Принято']);
  ok(e.sheet.merges.some(m => m.join() === '4,1,3,1') && e.sheet.merges.some(m => m.join() === '4,4,3,1'), 'A4:A6 и D4:D6 объединены: ' + JSON.stringify(e.sheet.merges));
  eq(e.val(4, 4), '[LIT-26-11][LogrusIT][ka][az][uz][Магазинка] Пуши');
  eq(e.val(7, 4), 'Перевод заявки (штат)', 'строки ниже просто сдвинулись');
  // автономер самой таблицы считает ячейки с номером — их число не изменилось
  let n = 0; for (let i = 2; i <= 11; i++) if (/^LIT-26-/.test(e.val(i, 1))) n++;
  eq(n, 3);
  eq(e.post({ action: 'myTasks', manager: 'Анна Первая' }).tasks.filter(t => t.id === 'LIT-26-11').map(t => t.row).sort(), [4, 5, 6]);
  eq(e.post({ action: 'previewTaskId', contractor: 'LogrusIT' }).id, 'LIT-26-13');
});

check('ещё строка в заказ из одной строки и в только что добавленный', () => {
  const e = load(SYN);
  const id = e.post({ action: 'addTask', task: NEW_TASK }).id;
  const r = e.post({ action: 'addTaskRow', id, task: { languages: ['Армянский'], link: 'https://band.wb.ru/wb/pl/x', status: 'В работе' } });
  eq([r.row, r.subject], [11, '[LIT-26-13][LogrusIT][ka][uz][hy][Магазинка] Новая акция']);
  ok(e.sheet.merges.some(m => m.join() === '10,1,2,1'));
  eq(e.post({ action: 'getTask', row: 11, id }).task.subject, r.subject);
  eq(e.post({ action: 'addTaskRow', id: 'LIT-26-777', task: {} }).ok, false);
});

check('смета: пишется в первую строку заказа, даже если ячейка объединена; повтор — «exists»', () => {
  const e = load(SYN);
  eq(e.post({ action: 'lookup', task: 'LIT-26-11' }), { ok: true, found: true, row: 4, matched: 'LIT-26-11',
    subject: '[LIT-26-11][LogrusIT][ka][az][Магазинка] Пуши', contractor: 'LogrusIT', manager: 'Анна Первая', link: '', total: '' });
  const w = e.post({ action: 'write', task: 'lit-26-11', total: 25200, link: 'https://disk.wb.ru/f/9', formula: '24000+1200' });
  eq([w.ok, w.row], [true, 4]);
  eq([e.val(4, 12), e.val(4, 13), e.sheet.peek(4, 13).formula, e.sheet.peek(4, 13).num], ['https://disk.wb.ru/f/9', 25200, '=24000+1200', '#,##0.00']);
  eq(e.sets().filter(x => !/^R4C1[23]$/.test(x)), []);
  eq(e.post({ action: 'write', task: 'LIT-26-11', total: 26000, link: 'https://disk.wb.ru/f/10' }), { ok: false, error: 'exists', link: 'https://disk.wb.ru/f/9', total: 25200 });
  const o = e.post({ action: 'write', task: 'LIT-26-11', total: 26000, link: 'https://disk.wb.ru/f/10', overwrite: true, version: 2 });
  eq([o.ok, o.was], [true, 25200]);
  ok(/12\.11\.2026 — смета обновлена \(версия 2\): было 25 200,00 ₽ → стало 26 000,00 ₽\nпрежняя смета: https:\/\/disk\.wb\.ru\/f\/9/.test(e.sheet.peek(4, 13).note), e.sheet.peek(4, 13).note);
  eq(e.post({ action: 'lookup', task: 'LIT-26-12' }).link, 'https://disk.wb.ru/f/1');
  eq(e.post({ action: 'lookup', task: 'XXX-26-1' }), { ok: true, found: false });
  eq(e.post({ action: 'write', task: 'XXX-26-1', total: 5, link: 'https://x' }).ok, false);
});

check('удаление: одиночная строка удаляется; первая строка заказа из нескольких — нет', () => {
  const e = load(SYN);
  const r = e.post({ action: 'deleteTask', row: 4, id: 'LIT-26-11' });
  ok(!r.ok && /первая строка заказа/.test(r.error), JSON.stringify(r));
  eq(e.sheet.ops, []);
  eq(e.post({ action: 'deleteTask', row: 9, id: 'LIT-26-12' }), { ok: true });
  eq([e.sheet.maxRows, e.row(9).join('')], [59, '']);
  // вторую строку заказа удалить можно: номер и тема остаются в первой
  eq(e.post({ action: 'deleteTask', row: 5, id: 'LIT-26-11', origSubject: '[LIT-26-11][LogrusIT][ka][az][Магазинка] Пуши' }), { ok: true });
  eq([e.val(4, 1), e.val(4, 4), e.val(5, 4)], ['LIT-26-11', '[LIT-26-11][LogrusIT][ka][az][Магазинка] Пуши', 'Перевод заявки (штат)']);
  eq(e.post({ action: 'deleteTask', row: 4, id: 'LIT-26-11' }), { ok: true }, 'теперь заказ из одной строки');
});

check('поиск и дубли', () => {
  const e = load(SYN);
  eq(e.post({ action: 'searchTasks', query: 'local-102' }).tasks.map(t => [t.row, t.id]), [[5, 'LIT-26-11']]);
  eq(e.post({ action: 'searchTasks', query: 'тексты' }).tasks.map(t => t.row), [8, 7]);
  eq(e.post({ action: 'searchTasks', query: '' }).tasks, []);
  eq(e.post({ action: 'checkDuplicates', task: { link: 'https://band.wb.ru/wb/pl/eee', subject: 'что-то' } }).duplicates.map(d => [d.id, d.why]), [['LIT-26-11', 'та же ссылка на Band']]);
  eq(e.post({ action: 'checkDuplicates', task: { subject: 'Баннер', ticket: '100' } }).duplicates.map(d => d.id), ['LIT-26-10']);
  eq(e.post({ action: 'checkDuplicates', task: { subject: 'Совсем новое', link: 'https://band.wb.ru/wb/pl/zzz' } }).duplicates, []);
});

check('«Задача»: новый тикет — только LOCAL-…; прежний текст в старой строке остаётся как есть', () => {
  const fx = JSON.parse(JSON.stringify(SYN));
  fx.rows[2][4] = 'Баннер Компас выгоды';
  const e = load(fx);
  ok(/LOCAL-1234/.test(e.post({ action: 'addTask', task: Object.assign({}, NEW_TASK, { ticket: 'что-то' }) }).error));
  const t = e.post({ action: 'getTask', row: 3, id: 'LIT-26-10' }).task;
  eq(t.ticket, 'Баннер Компас выгоды');
  t.comment = 'новый коммент';
  eq(e.post({ action: 'saveTask', task: t }).changed, 1);
  eq(e.sets(), ['R3C18']);
  t.ticket = 'ещё текст';
  eq(e.post({ action: 'saveTask', task: t }).ok, false);
  t.ticket = 'local 77';
  eq(e.post({ action: 'saveTask', task: t }).ok, true);
  eq(e.val(3, 5), 'LOCAL-77');
});

check('ещё строка: низ заказа объединён с соседней задачей — отказ без изменений', () => {
  const e = load(SYN);
  const r = e.post({ action: 'addTaskRow', id: 'LoG-26-3', task: { languages: ['Армянский'] } });
  ok(!r.ok && /\(E, F, G\) объединены со следующей задачей/.test(r.error), JSON.stringify(r));
  eq(e.sheet.ops, []);
});

check('выше «хвоста» листа строки-продолжения узнаются по виду (без номера и темы)', () => {
  const e = load(SYN, src => src.replace('const TAIL_ROWS = 400;', 'const TAIL_ROWS = 2;'));
  const r = e.post({ action: 'myTasks', manager: 'Анна Первая' });
  eq(r.tasks.map(t => [t.row, t.id]), [[9, 'LIT-26-12'], [5, 'LIT-26-11'], [4, 'LIT-26-11']]);
  eq(e.post({ action: 'lookup', task: 'LIT-26-11' }).row, 4);
});

check('строки заказа сдвинулись, пока панель была открыта: по «отпечатку» находится та же строка, а не соседняя', () => {
  const e = load(SYN);
  const t = e.post({ action: 'getTask', row: 5, id: 'LIT-26-11' }).task; // вторая строка заказа
  ok(t.key && t.key !== e.post({ action: 'getTask', row: 4, id: 'LIT-26-11' }).task.key, 'у строк одного заказа разные отпечатки');
  e.sheet.insertRowAfter(1); e.reset();                                   // кто-то вставил строку выше
  t.status = 'Отдано';
  eq(e.post({ action: 'saveTask', task: t }), { ok: true, row: 6, changed: 1 });
  eq(e.sets(), ['R6C11']);
  // из списка «Мои» — то же
  const mine = e.post({ action: 'myTasks', manager: 'Анна Первая' }).tasks.find(x => x.row === 5);
  e.sheet.insertRowAfter(1); e.reset();
  eq(e.post({ action: 'setStatus', row: mine.row, id: mine.id, origSubject: mine.subject, key: mine.key, status: 'Холд' }), { ok: true, row: 6 });
  eq(e.sets(), ['R6C11']);
  // старое расширение без отпечатка: строка с тем же номером и темой под тем же номером строки — не отличить, остальное — отказ
  const old = Object.assign({}, t, { key: undefined, row: 3 });
  const r = e.post({ action: 'saveTask', task: old });
  ok(!r.ok && /сдвинулись/.test(r.error), JSON.stringify(r));
  // строку изменили руками (другой ник) — отпечаток не сходится, правка не применяется
  e.sheet.cell(7, 7).v = '@другой'; e.reset();
  const r2 = e.post({ action: 'saveTask', task: Object.assign({}, t, { row: 7 }) });
  ok(!r2.ok && /не найдена/.test(r2.error), JSON.stringify(r2));
  eq(e.sheet.ops, []);
  // у задачи из одной строки сдвиг не страшен: находится по номеру
  eq(e.post({ action: 'setStatus', row: 9, id: 'LIT-26-12', origSubject: '[LIT-26-12][LogrusIT][uz][ПВЗ] Озвучка', status: 'Отдано' }), { ok: true, row: 11 });
});

check('сумма с длинным хвостом после запятой: сохранение без правок её не округляет', () => {
  const fx = JSON.parse(JSON.stringify(SYN));
  fx.rows[8][12] = 322919.6649999;
  const e = load(fx);
  const t = e.post({ action: 'getTask', row: 9, id: 'LIT-26-12' }).task;
  eq(e.post({ action: 'saveTask', task: Object.assign({}, t, { total: '322919.665' }) }).changed, 0);
  eq(e.sheet.ops, []);
  eq(e.post({ action: 'saveTask', task: Object.assign({}, t, { total: '322920' }) }).changed, 1);
  eq(e.val(9, 13), 322920);
  ok(/было 322 919,665 ₽ → стало 322 920,00 ₽/.test(e.sheet.peek(9, 13).note), e.sheet.peek(9, 13).note);
});

check('SP: прочерк в ячейке не стирается при сохранении; не число из формы — ошибка, а не NaN', () => {
  const fx = JSON.parse(JSON.stringify(SYN));
  fx.rows[8][16] = '-';
  const e = load(fx);
  const t = e.post({ action: 'getTask', row: 9, id: 'LIT-26-12' }).task;
  eq(t.sp, '-');
  eq(e.post({ action: 'saveTask', task: t }).changed, 0);
  eq(e.post({ action: 'saveTask', task: Object.assign({}, t, { sp: '' }) }).changed, 0, 'числовое поле формы прочерк не показывает');
  eq(e.sheet.ops, []);
  ok(/SP должен быть числом/.test(e.post({ action: 'saveTask', task: Object.assign({}, t, { sp: 'много' }) }).error));
  eq(e.sheet.ops, []);
  eq(e.post({ action: 'saveTask', task: Object.assign({}, t, { sp: '2,5' }) }).changed, 1);
  eq(e.val(9, 17), 2.5);
  const r = e.post({ action: 'addTask', task: Object.assign({}, NEW_TASK, { sp: 'x' }) });
  ok(!r.ok && /SP должен быть числом/.test(r.error), JSON.stringify(r));
  eq(e.sheet.maxRows, 60, 'строка не вставлена');
});

check('даты пишутся номером дня: в ячейке ровно выбранный день, даже если таблица «сдвигает» даты-объекты', () => {
  const fx = JSON.parse(JSON.stringify(SYN));
  fx.dateBug = 2; // как будто запись объектом Date уезжает на два дня
  const e = load(fx);
  const r = e.post({ action: 'addTask', task: Object.assign({}, NEW_TASK, { date: '2026-11-24', exactDeadline: '2026-11-26' }) });
  eq([r.ok, r.date, r.due, r.warn], [true, '24.11.2026', '26.11.2026', undefined]);
  eq([e.sheet.getRange(10, 3).getDisplayValue(), e.sheet.getRange(10, 10).getDisplayValue()], ['24.11', '26.11']);
  eq([e.day(10, 3), e.day(10, 10)], ['2026-11-24 00:00', '2026-11-26 00:00']);
  const t = e.post({ action: 'getTask', row: 10, id: r.id }).task;
  eq([t.date, t.exactDeadline], ['2026-11-24', '2026-11-26']);
  t.exactDeadline = '2026-12-01'; t.date = '2026-11-23'; t.deadline = '';
  const s2 = e.post({ action: 'saveTask', task: t });
  eq([s2.ok, s2.warn], [true, undefined]);
  eq([e.sheet.getRange(10, 3).getDisplayValue(), e.sheet.getRange(10, 10).getDisplayValue()], ['23.11', '1.12']);
  const row = e.post({ action: 'addTaskRow', id: r.id, task: { languages: ['Армянский'], exactDeadline: '2026-11-30' } });
  eq([e.sheet.getRange(row.row, 3).getDisplayValue(), e.sheet.getRange(row.row, 10).getDisplayValue()], ['23.11', '30.11']);
  // на конец года и високосный день
  const y = e.post({ action: 'addTask', task: Object.assign({}, NEW_TASK, { date: '2026-12-31', exactDeadline: '2028-02-29' }) });
  eq([e.sheet.getRange(y.row, 3).getDisplayValue(), e.sheet.getRange(y.row, 10).getDisplayValue(), e.day(y.row, 10)], ['31.12', '29.2', '2028-02-29 00:00']);
});

check('если таблица всё же показала не тот день — расширение получает предупреждение', () => {
  const e = load(SYN);
  const real = e.sheet.getRange(10, 3).constructor.prototype.getDisplayValue;
  e.sheet.getRange(10, 3).constructor.prototype.getDisplayValue = function () { return this.c === 3 && this.r === 10 ? '26.11' : real.call(this); };
  try {
    const r = e.post({ action: 'addTask', task: Object.assign({}, NEW_TASK, { date: '2026-11-24' }) });
    eq(r.ok, true);
    ok(/Проверьте дату в таблице: в ячейке C10 таблица показывает «26\.11», а должно быть 24\.11\.2026/.test(r.warn), r.warn);
  } finally { e.sheet.getRange(10, 3).constructor.prototype.getDisplayValue = real; }
});

check('проверка дат из редактора (checkDates) ничего не меняет', () => {
  const e = load(SYN);
  e.ctx.checkDates();
  ok(e.ctx.__log.some(l => /Часовой пояс таблицы: Europe\/Moscow · скрипта: America\/New_York/.test(l)), e.ctx.__log[0]);
  ok(e.ctx.__log.some(l => /C9 · LIT-26-12 · на экране «9\.11» · скрипт читает 09\.11\.2026 00:00/.test(l)), e.ctx.__log.join('\n'));
  eq(e.sheet.ops, []);
});

check('скорость: справочники из памяти лист не читают; добавление и смена статуса не перебирают объединения', () => {
  const e = load(SYN);
  e.post({ action: 'taskForm' }); // первый раз — читает и запоминает
  ok(e.calls(() => e.post({ action: 'taskForm' })) <= 1, 'taskForm из памяти');
  ok(e.calls(() => e.post({ action: 'managers' })) <= 1, 'managers из памяти');
  const add = e.calls(() => e.post({ action: 'addTask', task: NEW_TASK }));
  ok(add <= 16, 'addTask: обращений ' + add);
  const st = e.calls(() => e.post({ action: 'setStatus', row: 9, id: 'LIT-26-12', status: 'Отдано' }));
  ok(st <= 8, 'setStatus: обращений ' + st);
  const ids = e.calls(() => e.post({ action: 'previewTaskIds' }));
  ok(ids <= 4, 'previewTaskIds: обращений ' + ids);
});

check('отчёты и письма по расписанию — понятный отказ, неизвестное действие — ошибка', () => {
  const e = load(SYN);
  ok(/пока не подключено/.test(e.post({ action: 'reportLists' }).error));
  ok(/пока не подключено/.test(e.post({ action: 'mailStatus' }).error));
  eq(e.post({ action: 'чтото' }), { ok: false, error: 'Неизвестное действие' });
  eq(e.sheet.ops, []);
});

check('проверка из редактора (checkTable) ничего не меняет и пишет, что видит', () => {
  const e = load(SYN);
  e.ctx.checkTable();
  ok(e.ctx.__log.some(l => /строк с задачами: 8/.test(l)), e.ctx.__log.join(' | '));
  ok(e.ctx.__log.some(l => /LIT-26-13/.test(l)));
  eq(e.sheet.ops, []);
});

// ---------- Копия настоящего листа ----------
if (process.env.FIXTURE) {
  console.log('\nКопия настоящего листа');
  const REAL = JSON.parse(fs.readFileSync(process.env.FIXTURE, 'utf8'));
  const lastRow = REAL.rows.length;
  const maxOf = (base) => Math.max(0, ...REAL.rows.map(r => { const m = String(r[0]).match(new RegExp('^' + base + '-(\\d+)$')); return m ? Number(m[1]) : 0; }));

  check('справочники и номера', () => {
    const e = load(REAL);
    const t0 = Date.now();
    const l = e.post({ action: 'taskForm' }).lists;
    console.log('      подрядчики: ' + l.contractors.join(', ') + ' · менеджеров ' + l.managers.length + ' · продуктов ' + l.products.length +
      ' · языков ' + (l.languages.regular.length + l.languages.shtat.length + l.languages.rare.length) + ' · ' + (Date.now() - t0) + ' мс');
    ok(l.contractors.length === 5 && l.managers.length > 5 && l.products.length > 10 && l.languages.shtat.length >= 5);
    ok(l.statuses.indexOf('Статус') === -1 && l.managers.indexOf('Менеджер') === -1);
    const ids = e.post({ action: 'previewTaskIds' }).ids;
    console.log('      следующие номера: ' + JSON.stringify(ids));
    eq(ids.LogrusIT, 'LIT-26-' + (maxOf('LIT-26') + 1));
    eq(ids.LogrusGlobal, 'LoG-26-' + (maxOf('LoG-26') + 1));
    eq(ids['Бюро переводов'], 'BP-26-' + (maxOf('BP-26') + 1));
  });

  check('новая задача встаёт сразу под последней, остальной лист не меняется', () => {
    const e = load(REAL);
    const before = JSON.stringify(e.sheet.cells.slice(0, lastRow).map(r => r.map(c => c.v)));
    const l = e.post({ action: 'taskForm' }).lists;
    const r = e.post({ action: 'addTask', task: { contractor: 'LogrusIT', subject: 'Проверка', product: l.products[0], date: '2026-11-12', exactDeadline: '2026-11-13',
      status: 'В работе', manager: l.managers[0], languages: [l.languages.regular[0], l.languages.shtat[0]], link: 'https://band.wb.ru/wb/pl/test' } });
    eq([r.ok, r.row, r.id], [true, lastRow + 1, 'LIT-26-' + (maxOf('LIT-26') + 1)]);
    eq(JSON.stringify(e.sheet.cells.slice(0, lastRow).map(row => row.map(c => c.v))), before);
    console.log('      ' + e.row(r.row).map(v => (v instanceof Date ? formatDate(v, TABLE_TZ, 'dd.MM') : v)).filter(v => v !== '').join(' | '));
  });

  check('«Мои» у каждого менеджера: у всех открытых задач есть тема; правка без изменений ничего не трогает', () => {
    const e = load(REAL);
    const managers = e.post({ action: 'taskForm' }).lists.managers;
    let open = 0, noId = 0, checked = 0;
    managers.forEach(m => {
      const r = e.post({ action: 'myTasks', manager: m });
      open += r.tasks.length;
      r.tasks.forEach(t => {
        ok(t.subject, 'нет темы: строка ' + t.row);
        if (!t.id) noId++;
        if (checked < 400) {
          checked++;
          const g = e.post({ action: 'getTask', row: t.row, id: t.id, origSubject: t.subject });
          ok(g.ok, 'getTask ' + t.row + ': ' + g.error);
          const s = e.post({ action: 'saveTask', task: g.task });
          eq([s.ok, s.changed, e.sheet.ops.length], [true, 0, 0], 'строка ' + t.row + ' ' + JSON.stringify(e.sheet.ops.slice(0, 3)) + (s.error || ''));
        }
      });
    });
    console.log('      открытых задач: ' + open + ', из них без номера (штат): ' + noId + ', правка «вхолостую» проверена на ' + checked);
    ok(open > 50);
  });

  check('смета в последний заказ из нескольких строк — в его первую строку', () => {
    const e = load(REAL);
    const m = REAL.merges.filter(x => x[1] === 1 && x[2] > 1).sort((a, b) => b[0] - a[0])[0];
    const id = REAL.rows[m[0] - 1][0];
    const look = e.post({ action: 'lookup', task: id });
    eq([look.found, look.row], [true, m[0]]);
    const w = e.post({ action: 'write', task: id, total: 1234.5, link: 'https://disk.wb.ru/f/test', overwrite: true });
    eq([w.ok, w.row], [true, m[0]]);
    eq(e.sets().every(x => new RegExp('^R' + m[0] + 'C1[23]$').test(x)), true, e.sets().join());
    console.log('      ' + id + ' → строка ' + m[0]);
  });

  check('ещё строка в последние заказы из нескольких строк (или понятный отказ, если низ заказа объединён с соседней задачей)', () => {
    const orders = REAL.merges.filter(x => x[1] === 1 && x[2] > 1).sort((a, b) => b[0] - a[0]).slice(0, 25);
    let added = 0, refused = 0;
    orders.forEach(m => {
      const e = load(REAL);
      const id = REAL.rows[m[0] - 1][0];
      const l = e.post({ action: 'taskForm' }).lists;
      const r = e.post({ action: 'addTaskRow', id, task: { languages: [l.languages.regular[1]], status: 'В работе', link: 'https://band.wb.ru/wb/pl/more' } });
      if (!r.ok) {
        ok(/объединены со следующей задачей/.test(r.error), id + ': ' + r.error);
        eq(e.sheet.ops, [], id + ': при отказе ничего не меняется');
        refused++;
        return;
      }
      added++;
      eq(r.row, m[0] + m[2], id);
      ok(e.sheet.merges.some(x => x[0] === m[0] && x[1] === 1 && x[2] === m[2] + 1), id + ': номер объединён на все строки заказа');
      eq(e.post({ action: 'getTask', row: r.row, id }).task.id, id);
    });
    console.log('      заказов: ' + orders.length + ', строка добавлена: ' + added + ', отказ: ' + refused);
    ok(added > 10);
  });
}

console.log(`\n${passed} проверок прошло` + (failed ? `, ${failed} НЕ прошло` : ''));
process.exit(failed ? 1 : 0);
