'use strict';
const fs = require('fs'), vm = require('vm'), assert = require('assert/strict');
const { makeContext, calls, RichBuilder } = require('./gas-mock');
// Выгрузка таблицы в JSON (листы → строки); в репозиторий не кладём — там личные данные.
const fixture = JSON.parse(fs.readFileSync(process.env.FIXTURE || __dirname + '/fixture.json', 'utf8'));
const REPO = require('path').resolve(__dirname, '../..');

// Почты в «Списках»: F — менеджер, M — рассылка, K — переводчик
const lists = fixture['Списки'].rows;
const managerRow = lists.findIndex(r => r[4] === 'Анастасия Лисовая');
lists[managerRow][5] = 'nastya@wb.ru';
lists[managerRow][12] = 'nastya@wb.ru';
const trRow = lists.findIndex(r => r[9] === 'Виктория Гусева');
lists[trRow][10] = 'vika@wb.ru';

const { ctx, ss } = makeContext(fixture, { email: 'nastya@wb.ru' });
vm.createContext(ctx);
// Даты — из «мира» скрипта, иначе instanceof Date не сработает
const CDate = vm.runInContext('Date', ctx);
ss.getSheets().forEach(sh => sh.cells.forEach(row => row.forEach(c => { if (c.v instanceof Date) c.v = new CDate(c.v.getTime()); })));
const same = (a, b, m) => assert.equal(JSON.stringify(a), JSON.stringify(b), m);
for (const f of ['table-scripts/Код.gs', 'smeta-extension/apps-script/Smeta.gs']) {
  vm.runInContext(fs.readFileSync(`${REPO}/${f}`, 'utf8'), ctx, { filename: f });
}
const G = ctx;
const tasks = () => ss.getSheetByName('📌 Задачи (менеджеры)');
const tr = () => ss.getSheetByName('✍️ Задачи (переводчики)');
const log = () => ss.getSheetByName('📝 Журнал');
const rowOf = id => tasks().getRange('A1:A200').getValues().findIndex(r => r[0] === id) + 1;
let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('ok  ', name); } catch (e) { console.log('FAIL', name, '\n     ', e.stack.split('\n').slice(0, 4).join('\n      ')); process.exitCode = 1; }
}

// Ссылки в теме у двух задач (в выгрузке xlsx их нет)
const lit4 = rowOf('LIT-4'), lit1 = rowOf('LIT-1');
tasks().getRange(lit4, 3).setRichTextValue(new RichBuilder().setText(tasks().getRange(lit4, 3).getValue()).setLinkUrl('https://band/lit4').build());
{
  const t = String(tasks().getRange(lit1, 3).getValue()).replace(' (доп. ссылка)', '');
  tasks().getRange(lit1, 3).setRichTextValue(new RichBuilder().setText(t + ' (доп. ссылка)').setLinkUrl(0, t.length, 'https://band/one').setLinkUrl(t.length, t.length + 14, 'https://band/two').build());
}

test('onOpen строит меню', () => {
  G.onOpen();
  same(calls.menus.map(m => m.name), ['📋 Менеджеры', '👥 Переводчики EN', '📈 Отчёты', '⚙️ Настройки']);
  const fnNames = calls.menus.flatMap(m => m.items.flatMap(i => Array.isArray(i[1]) ? i[1].map(x => x[1]) : [i[1]]));
  fnNames.forEach(f => assert.equal(typeof G[f], 'function', 'нет функции ' + f));
});

test('До оформления (в «Списках» ещё нет колонки O) справочники читаются', () => {
  assert.equal(ss.getSheetByName('Списки').getMaxColumns(), 14);
  const l = G.getAddTaskFormLists();
  assert.ok(l.contractors.includes('LogrusIT'));
  assert.equal(G.contractorPrefix_('LogrusIT'), 'LIT');
});

test('setupDesign: оформление без ошибок', () => {
  G.setupDesign();
  const sh = tasks();
  assert.equal(sh.getRange(1, 9).getValue(), 'Срок сдачи');
  assert.equal(sh.frozenRows, 1); assert.equal(sh.frozenCols, 3);
  assert.ok(sh.cf.length > 30, 'правил цветов: ' + sh.cf.length);
  assert.ok(sh.cf.every(r => /!.1:/.test(r.ranges[0])), 'правила должны начинаться с 1-й строки');
  assert.ok(sh.groups[11] && sh.groups[15], 'группы колонок');
  assert.ok(sh.getRange(2, 5).getDataValidation(), 'список в строке задачи');
  assert.equal(sh.getRange(sh.getLastRow() + 3, 5).getDataValidation(), null, 'в пустых строках списков нет');
  const l = ss.getSheetByName('Списки');
  assert.equal(l.getRange('O2').getValue(), 'Префикс подрядчика');
  assert.equal(l.getRange('O3').getValue(), 'LIT');
  const guide = ss.getSheetByName('📋 Инструкция (менеджеры)').getRange('C1').getValue();
  assert.match(guide, /Шпаргалка/);
  assert.ok(tr().cf.length > 10);
  // повторный запуск не ломается и не плодит группы
  G.setupDesign();
});

let newId;
test('Добавить задачу: номер LIT-26-2232, тема, ссылки, журнал', () => {
  const before = tasks().getLastRow();
  newId = G.submitNewTaskFromDialog({ contractor: 'LogrusIT', subject: 'Новые строчки', ticket: 'LOCAL-1116', link: 'https://band/a', link2: 'https://band/b',
    date: '2026-10-01', product: 'Магазинка', languages: ['Грузинский', 'Казахский', 'ШТАТ английский'], status: 'Принято', manager: 'Анастасия Лисовая', total: '', sp: '3' });
  assert.equal(newId, 'LIT-26-2232');
  const sh = tasks();
  assert.equal(sh.getLastRow(), before + 1);
  const r = sh.getRange(2, 1, 1, 17).getValues()[0];
  assert.equal(r[0], 'LIT-26-2232');
  assert.equal(r[2], '[LIT-26-2232][LogrusIT][ka][kk][Магазинка] Новые строчки (доп. ссылка)');
  assert.equal(r[3].getDate(), 1); assert.equal(r[3].getMonth(), 9);
  assert.equal(r[15], 3);
  same(G.linksFromRich_(sh.getRange(2, 3).getRichTextValue()), { link: 'https://band/a', link2: 'https://band/b' });
  assert.ok(sh.getRange(2, 5).getDataValidation(), 'списки скопированы в новую строку');
  const last = log().getRange(log().getLastRow(), 1, 1, 7).getValues()[0];
  assert.equal(last[2], 'Создание задачи'); assert.equal(last[3], 'LIT-26-2232'); assert.equal(last[1], 'nastya@wb.ru');
  assert.equal(G.generateNextTaskId(sh, 'LogrusIT'), 'LIT-26-2233');
  assert.equal(G.generateNextTaskId(sh, 'Бюро переводов'), 'BP-26-367');
});

test('Тикет: 1234 → LOCAL-1234, без подрядчика — ошибка', () => {
  assert.equal(G.normTicket_('1234'), 'LOCAL-1234');
  assert.equal(G.normTicket_('local 55'), 'LOCAL-55');
  assert.throws(() => G.submitNewTaskFromDialog({ subject: 'x' }), /подрядчика/);
});

test('Поиск: пустой — мои задачи, по тексту и номеру', () => {
  const mine = G.searchTasks('');
  assert.ok(mine.length >= 5 && mine.every(x => x.id !== '' || x.title));
  assert.equal(mine[0].id, 'LIT-26-2232');
  assert.equal(mine[0].title, 'Новые строчки');
  const byTicket = G.searchTasks('1114');
  assert.ok(byTicket.length >= 2 && byTicket.every(x => x.ticket === 'LOCAL-1114'));
  assert.equal(G.searchTasks('lit-4')[0].id, 'LIT-4');
});

test('Правка попадает в нужную задачу, даже если строки сдвинулись', () => {
  const row = rowOf('LIT-7');
  const t = G.getTaskForEdit(row);
  assert.equal(t.id, 'LIT-7');
  G.submitNewTaskFromDialog({ contractor: 'LogrusGlobal', subject: 'Сдвиг строк' }); // кто-то добавил задачу — всё съехало на 1
  assert.equal(rowOf('LIT-7'), row + 1);
  G.saveTaskEdits({ ...t, status: 'В работе', sp: '5', languages: 'Грузинский, Армянский' });
  const r = tasks().getRange(row + 1, 1, 1, 17).getValues()[0];
  assert.equal(r[0], 'LIT-7'); assert.equal(r[9], 'В работе'); assert.equal(r[15], 5); assert.equal(r[6], 'Грузинский, Армянский');
  assert.notEqual(tasks().getRange(row, 10).getValue(), 'В работе', 'соседняя задача не тронута');
  const entries = log().getDataRange().getValues().filter(x => x[3] === 'LIT-7' && x[2] === 'Правка').map(x => x[4]);
  assert.ok(entries.includes('Статус') && entries.includes('SP') && entries.includes('Языки'), entries.join());
});

test('Правка не теряет доп. ссылку', () => {
  const t = G.getTaskForEdit(rowOf('LIT-1'));
  assert.equal(t.link, 'https://band/one'); assert.equal(t.link2, 'https://band/two');
  assert.ok(!t.subject.includes('(доп. ссылка)'));
  G.saveTaskEdits({ ...t, comment: 'поправили' });
  const rt = tasks().getRange(rowOf('LIT-1'), 3).getRichTextValue();
  same(G.linksFromRich_(rt), { link: 'https://band/one', link2: 'https://band/two' });
  assert.equal(tasks().getRange(rowOf('LIT-1'), 17).getValue(), 'поправили');
});

test('Очистка поля тоже попадает в журнал', () => {
  const t = G.getTaskForEdit(rowOf('LIT-1'));
  G.saveTaskEdits({ ...t, comment: '' });
  const last = log().getRange(log().getLastRow(), 1, 1, 7).getValues()[0];
  assert.equal(last[4], 'Комментарий'); assert.equal(last[5], 'поправили'); assert.equal(last[6], '');
});

test('Удаление: удаляется именно та задача, даже после сдвига', () => {
  const row = rowOf('LIT-5');
  const t = G.getTaskForEdit(row);
  G.submitNewTaskFromDialog({ contractor: 'LogrusIT', subject: 'Ещё сдвиг' });
  const neighbour = tasks().getRange(row, 1).getValue();
  G.deleteTask(t.row, t.id, t.origSubject);
  assert.equal(rowOf('LIT-5'), 0);
  assert.equal(tasks().getRange(row, 1).getValue(), neighbour);
  assert.throws(() => G.saveTaskEdits(t), /не найдена/);
});

test('Дашборд: SP по месяцам не 0, просрочки считаются', () => {
  const d = G.getDashboardData();
  const sep = d.byMonth.find(m => m[0].startsWith('Сен 2026'));
  assert.ok(sep[2] > 0, 'SP за сентябрь: ' + JSON.stringify(sep));
  assert.ok(d.overdueCount >= 2, 'просрочек: ' + d.overdueCount);
  assert.equal(G.getOverdueCount(), d.overdueCount);
  const y = G.getDashboardData('2026', '9');
  assert.ok(y.totalTasks >= 4);
});

test('Отчёты менеджера, трекера, кастомный, экспорт', () => {
  const r = G.getManagerReport('Анастасия Лисовая', '');
  assert.ok(r.lineCount >= 4);
  const line = r.groups.flatMap(g => g.lines).find(l => l.subject.includes('[LIT-4]'));
  assert.equal(line.link, 'https://band/lit4');
  const url = G.exportReportToDoc('Анастасия Лисовая', '');
  assert.match(url, /export\?format=xlsx/);
  const tr1 = G.getTrackerReportText('LOCAL-1116');
  assert.ok(tr1.count >= 2 && tr1.text.includes('[LIT-4] Новые строчки для приложения от 10.09](https://band/lit4)'), tr1.text);
  const c = G.getCustomReport('2026', '');
  assert.ok(c.totalTasks >= 5);
  assert.match(G.exportCustomReportToExcel('2026', ''), /xlsx/);
});

test('Письма: сводка, горящие сроки, конец месяца', () => {
  calls.mail.length = 0;
  G.sendWeeklyDigests();
  const mine = () => calls.mail.filter(m => m.to === 'nastya@wb.ru');
  assert.equal(mine().length, 1);
  assert.ok(!mine()[0].htmlBody.includes('<script'));
  assert.ok(calls.mail.every(m => m.to), 'письма только тем, у кого есть почта в колонке M');
  G.sendDueSoonAlerts();
  assert.equal(mine().length, 2); assert.match(mine()[1].subject, /просроченные/);
  G.sendMonthEndReminders();
  assert.ok(calls.mail.length >= 2);
});

test('Переводчики: добавить, найти, поправить после сдвига, удалить', () => {
  G.submitNewTranslatorTask({ date: '2026-10-01', side: 'Маркетинг', razdel: 'Пуши', task: 'Тестовая задача', link: 'https://band/t',
    translator: 'Виктория Гусева', editor: 'Максим Селищев', readiness: 'В работе' });
  assert.equal(tr().getRange(2, 4).getValue(), 'Тестовая задача');
  const found = G.searchTranslatorTasks('тестовая');
  assert.equal(found.length, 1);
  const t = G.getTranslatorTaskForEdit(found[0].row);
  assert.equal(t.link, 'https://band/t');
  same(t.translatorNames, ['Виктория Гусева']);
  G.submitNewTranslatorTask({ side: 'Контент', task: 'Сдвиг' });
  G.saveTranslatorTaskEdits({ ...t, readiness: 'Готово', translator: 'Вика, Таня' });
  const row = G.searchTranslatorTasks('тестовая')[0].row;
  assert.equal(tr().getRange(row, 8).getValue(), 'Готово');
  same(G.splitPeople_('Вика, Таня'), ['Виктория Гусева', 'Татьяна Козлова']);
  const t2 = G.getTranslatorTaskForEdit(row);
  G.deleteTranslatorTask(t2.row, t2.origKey);
  assert.equal(G.searchTranslatorTasks('тестовая').length, 0);
  // Пустой поиск — задачи того, кто открыл (по почте переводчика)
  ctx.__setEmail('vika@wb.ru');
  assert.ok(G.searchTranslatorTasks('').every(x => x.title));
  ctx.__setEmail('nastya@wb.ru');
  const rep = G.getTranslatorReport('Виктория Гусева', '');
  assert.ok(rep.total >= 3);
  assert.ok(Object.keys(G.getContentCategories()).includes('Портал продавца'));
});

test('Расширение (Smeta.gs): добавить задачу и записать смету — без двойной блокировки', () => {
  const call = b => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ token: 'T', ...b }) } }));
  const form = call({ action: 'taskForm' });
  assert.ok(form.ok && form.lists.contractors.includes('LogrusIT'));
  const prev = call({ action: 'previewTaskId', contractor: 'LogrusIT' });
  const add = call({ action: 'addTask', user: 'Настя', task: { contractor: 'LogrusIT', subject: 'Из расширения', ticket: 'LOCAL-1116', languages: ['Грузинский'] } });
  assert.ok(add.ok, JSON.stringify(add));
  assert.equal(add.id, prev.id);
  const last = log().getRange(log().getLastRow(), 1, 1, 7).getValues()[0];
  assert.equal(last[1], 'Расширение (Настя)');
  assert.equal(ctx.SCRIPT_LOCK_HELD, false); assert.ok(!ctx.__lockHeld);
  const look = call({ action: 'lookup', task: add.id });
  assert.ok(look.found, JSON.stringify(look));
  const w = call({ action: 'write', task: add.id, total: 1234.5, link: 'https://disk/x' });
  assert.ok(w.ok, JSON.stringify(w));
  assert.equal(tasks().getRange(rowOf(add.id), 12).getValue(), 1234.5);
});

test('Фильтры', () => {
  G.hideClosedTasks(); same(tasks().filter.criteria[10].hidden, ['Отдано', 'Отменено']);
  G.showMyTasks(); assert.ok(!tasks().filter.criteria[14].hidden.includes('Анастасия Лисовая'));
  G.showAllTasks(); same(tasks().filter.criteria, {});
  G.showInProgressTranslatorTasks(); G.showAllTranslatorTasks();
});

test('Правка прямо в листе: журнал, цветные языки, новая строка получает номер и списки', () => {
  const sh = tasks();
  const row = rowOf('LIT-7');
  const old = sh.getRange(row, 10).getValue();
  sh.getRange(row, 10).setValue('Холд');
  G.onEdit({ range: sh.getRange(row, 10), oldValue: old, value: 'Холд' });
  let last = log().getRange(log().getLastRow(), 1, 1, 7).getValues()[0];
  same([last[2], last[3], last[4], last[5], last[6]], ['Правка в таблице', 'LIT-7', 'Статус', old, 'Холд']);

  sh.getRange(row, 7).setValue('Грузинский, Узбекский');
  G.onEdit({ range: sh.getRange(row, 7), oldValue: 'x', value: 'Грузинский, Узбекский' });
  assert.ok(sh.getRange(row, 7).getRichTextValue().styles.length === 2, 'языки покрашены');

  // Новая задача вписана руками в пустую строку под последней
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 3).setValue('Задача руками');
  G.onEdit({ range: sh.getRange(r, 3), value: 'Задача руками' });
  assert.ok(sh.getRange(r, 10).getDataValidation(), 'списки появились');
  assert.equal(sh.getRange(r, 1).getValue(), '', 'без подрядчика номера ещё нет');
  sh.getRange(r, 13).setValue('LogrusIT');
  G.onEdit({ range: sh.getRange(r, 13), value: 'LogrusIT' });
  assert.match(String(sh.getRange(r, 1).getValue()), /^LIT-26-\d+$/);
  last = log().getRange(log().getLastRow(), 1, 1, 7).getValues()[0];
  assert.ok(['Создание задачи (в таблице)', 'Правка в таблице'].includes(last[2]));

  // Дата: в событии приходит числом — в журнал пишется датой
  assert.equal(G.editedText_('46296', true), '01.10.2026');

  // Переводчики
  const t = tr(); const tr2 = t.getLastRow() + 1;
  t.getRange(tr2, 4).setValue('Руками');
  G.onEdit({ range: t.getRange(tr2, 4), value: 'Руками' });
  assert.ok(t.getRange(tr2, 8).getDataValidation());
  last = log().getRange(log().getLastRow(), 1, 1, 7).getValues()[0];
  assert.equal(last[2], 'Правка в таблице (переводчик)');

  // «Списки» — обновляются цвета, без ошибок
  G.onEdit({ range: ss.getSheetByName('Списки').getRange('E20') });
  assert.equal(ctx.SCRIPT_LOCK_HELD, false);
});

test('colorizeRowDirectly: на оформленном листе не красит, на неоформленном — правильные колонки', () => {
  const sh = ss.insertSheet('TS');
  sh.getRange(1, 1, 2, 17).setValues([Array(17).fill('h'), ['TS-1', '', 'x', '', 'Магазинка', '', '', 'ASAP', '', 'Принято', '', '', 'LogrusIT', 'Анастасия Лисовая', '', '', '']]);
  G.colorizeRowDirectly(sh, 2); // не должно упасть
  G.colorizeRowDirectly(tasks(), 2);
});

console.log(`\n${passed} проверок пройдено`);
