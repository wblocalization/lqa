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

// Старая таблица (лист «Localization Misc») для проверки переноса — тоже не в репозитории
const oldFixture = process.env.OLD_FIXTURE ? JSON.parse(fs.readFileSync(process.env.OLD_FIXTURE, 'utf8')) : null;
const OLD_ID = 'OLDTABLE_1234567890abcdefghij';
const { ctx, ss, others } = makeContext(fixture, { email: 'nastya@wb.ru', others: oldFixture ? { [OLD_ID]: oldFixture } : {} });
vm.createContext(ctx);
// Даты — из «мира» скрипта, иначе instanceof Date не сработает
const CDate = vm.runInContext('Date', ctx);
[ss, ...Object.values(others)].forEach(book => book.getSheets().forEach(sh => sh.cells.forEach(row => row.forEach(c => { if (c.v instanceof Date) c.v = new CDate(c.v.getTime()); }))));
const same = (a, b, m) => assert.equal(JSON.stringify(a), JSON.stringify(b), m);
// BUNDLE=1 — проверяем сборку «всё в одном файле» вместо отдельных файлов
const FILES = process.env.BUNDLE ? ['table-scripts/ВсёВОдном.gs'] : ['table-scripts/Код.gs', 'table-scripts/Перенос.gs', 'smeta-extension/apps-script/Smeta.gs'];
for (const f of FILES) {
  vm.runInContext(fs.readFileSync(`${REPO}/${f}`, 'utf8'), ctx, { filename: f });
}
const G = ctx;
const tasks = () => ss.getSheetByName('📌 Задачи (менеджеры)');
const tr = () => ss.getSheetByName('✍️ Задачи (переводчики)');
const log = () => ss.getSheetByName('📝 Журнал');
const rowOf = id => tasks().getRange('A1:A200').getValues().findIndex(r => r[0] === id) + 1;
let passed = 0;
function test(name, opts, fn) {
  if (typeof opts === 'function') { fn = opts; opts = {}; }
  if (opts.skip) { console.log('skip', name); return; }
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

test('Окна открываются', { skip: !process.env.BUNDLE }, () => {
  assert.match(G.include('Общее'), /<style>/);
  ['showAddTaskDialog', 'showSearchEditSidebar', 'showDashboard', 'showManagerReportSidebar', 'showCustomReportSidebar',
   'showAddTranslatorTaskDialog', 'showSearchEditTranslatorSidebar', 'showTranslatorReportSidebar'].forEach(f => G[f]());
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
  assert.ok(!sh.groups['11:collapsed'] && !sh.groups['15:collapsed'], 'колонки не свёрнуты — «Смета» видна');
  assert.equal(sh.cell(2, 3).f.bold, true, 'тема письма жирная');
  assert.ok(ss.getSheetByName('📋 Инструкция (менеджеры)').cards.length >= 6, 'карточки в шпаргалке');
  assert.ok(sh.getRange(2, 5).getDataValidation(), 'список в строке задачи');
  assert.equal(sh.getRange(sh.getLastRow() + 3, 5).getDataValidation(), null, 'в пустых строках списков нет');
  const l = ss.getSheetByName('Списки');
  assert.equal(l.getRange('O2').getValue(), 'Префикс подрядчика');
  assert.equal(l.getRange('O3').getValue(), 'LIT');
  const guide = ss.getSheetByName('📋 Инструкция (менеджеры)').getRange('B1').getValue();
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

test('Повторить задачу: последние задачи, похожие темы схлопываются', () => {
  const list = G.getRecentTasksForRepeat('Анастасия Лисовая');
  assert.ok(list.length >= 2);
  const app = list.filter(t => /Новые строчки для приложения/.test(t.title) && t.contractor === 'LogrusIT');
  assert.equal(app.length, 1, 'одна «для приложения», а не по штуке на каждую дату');
  assert.ok(Array.isArray(app[0].languages) && app[0].languages.length > 3);
  const all = G.getRecentTasksForRepeat('*');
  assert.ok(all.length >= list.length);
});

test('Проверка дублей', () => {
  const t = G.getTaskForEdit(rowOf('LIT-8'));
  const d = G.findDuplicateTasks({ subject: t.subject, ticket: '1116' });
  assert.ok(d.some(x => x.id === 'LIT-8'), JSON.stringify(d));
  assert.equal(G.findDuplicateTasks({ subject: 'Совсем новая тема', ticket: '1116' }).length, 0);
  const byLink = G.findDuplicateTasks({ subject: 'другое', link: 'https://band/lit4' });
  assert.equal(byLink[0].id, 'LIT-4'); assert.equal(byLink[0].why, 'та же ссылка на Band');
});

test('Мои задачи и статус из расширения', () => {
  const call = b => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ token: 'T', ...b }) } }));
  const none = call({ action: 'myTasks' });
  assert.equal(none.tasks.length, 0, 'без выбранного менеджера — пусто (а не задачи владельца)');
  assert.ok(none.managers.includes('Анастасия Лисовая'));
  const mine = call({ action: 'myTasks', manager: 'Анастасия Лисовая' });
  assert.ok(mine.tasks.length >= 3 && mine.overdue >= 1, JSON.stringify(mine).slice(0, 300));
  assert.ok(mine.tasks[0].overdue, 'просроченные первыми');
  assert.ok(mine.statuses.includes('Отдано'));
  const t = mine.tasks[0];
  const res = call({ action: 'setStatus', user: 'Настя', row: t.row, id: t.id, origSubject: t.subject, status: 'Отдано' });
  assert.ok(res.ok, JSON.stringify(res));
  const after = call({ action: 'myTasks', manager: 'Анастасия Лисовая' });
  assert.ok(!after.tasks.some(x => x.id === t.id && x.subject === t.subject), 'закрытая пропала из списка');
  const last = log().getRange(log().getLastRow(), 1, 1, 7).getValues()[0];
  same([last[1], last[4], last[6]], ['Расширение (Настя)', 'Статус', 'Отдано']);
  const rec = call({ action: 'recentTasks' });
  assert.ok(rec.ok && rec.tasks.length);
  const dup = call({ action: 'checkDuplicates', task: { subject: 'x', link: 'https://band/lit4' } });
  assert.equal(dup.duplicates.length, 1);
});

test('Перенос истории из старой таблицы', { skip: !oldFixture }, () => {
  const before = tasks().getLastRow();
  ctx.__prompt = 'https://docs.google.com/spreadsheets/d/' + OLD_ID + '/edit#gid=0';
  calls.alerts.length = 0;
  G.migrateFromOldTable();
  const added = tasks().getLastRow() - before;
  assert.ok(added > 3700, 'перенесено ' + added + '; ' + calls.alerts.join(' // '));
  assert.match(calls.alerts[0], /Будет добавлено задач: \d+/);
  const migrationAlert = calls.alerts[0];
  const rows = tasks().getRange(1, 1, tasks().getLastRow(), 17).getValues();
  // Одна задача подрядчику на три строки (разные запросы из Band): номер и тема у всех, сумма — только у первой
  const g = rows.filter(r => r[0] === 'LIT-26-1953');
  assert.equal(g.length, 3);
  assert.ok(g.every(r => /^\[LIT-26-1953\]/.test(r[2])));
  assert.equal(g.filter(r => r[11] !== '').length, 1, 'сумма не задвоилась');
  assert.ok(g.every(r => r[1] === 'LOCAL-493'), 'тикет из колонки «Задача»');
  assert.ok(g.every(r => r[4] === 'WBP'), 'сторона → продукт по памятке');
  const firstOld = rows.findIndex(r => r[0] === 'LIT-25-1') + 1;
  assert.equal(G.linksFromRich_(tasks().getRange(firstOld, 3).getRichTextValue()).link.slice(0, 26), 'https://band.wb.ru/wb/pl/g');
  assert.equal(rows[rows.length - 1][0], 'LIT-25-1', 'самые старые — внизу');
  // Повторный запуск ничего не задваивает
  calls.alerts.length = 0;
  G.migrateFromOldTable();
  assert.equal(tasks().getLastRow() - before, added);
  assert.match(calls.alerts[0], /Переносить нечего/);
  // Правка второй строки общего номера попадает во вторую, а не в первую
  const second = rows.findIndex((r, i) => r[0] === 'LIT-26-1953' && i > rows.findIndex(x => x[0] === 'LIT-26-1953')) + 1;
  const t = G.getTaskForEdit(second);
  G.saveTaskEdits({ ...t, comment: 'вторая строка' });
  assert.equal(tasks().getRange(second, 17).getValue(), 'вторая строка');
  assert.notEqual(tasks().getRange(second - 1, 17).getValue(), 'вторая строка');
  // Номера новых задач продолжаются после перенесённых
  const next = G.generateNextTaskId(tasks(), 'LogrusIT');
  assert.ok(Number(next.split('-')[2]) > 2231 && !rows.some(r => r[0] === next), next);
  console.log('      окно подтверждения:', calls.alerts.length ? '' : '', migrationAlert.replace(/\n/g, ' | '));
  // Старая таблица запомнилась: новые номера сверяются с ней
  others[OLD_ID].getSheetByName('Localization Misc').appendRow(['LIT-26-9999']);
  ctx.__clearCache();
  assert.equal(G.generateNextTaskId(tasks(), 'LogrusIT'), 'LIT-26-10000');
  const d = G.getDashboardData('2025');
  assert.ok(d.totalTasks > 800, 'дашборд видит историю: ' + d.totalTasks);
});

test('colorizeRowDirectly: на оформленном листе не красит, на неоформленном — правильные колонки', () => {
  const sh = ss.insertSheet('TS');
  sh.getRange(1, 1, 2, 17).setValues([Array(17).fill('h'), ['TS-1', '', 'x', '', 'Магазинка', '', '', 'ASAP', '', 'Принято', '', '', 'LogrusIT', 'Анастасия Лисовая', '', '', '']]);
  G.colorizeRowDirectly(sh, 2); // не должно упасть
  G.colorizeRowDirectly(tasks(), 2);
});

// Картинка шпаргалки: GUIDE_HTML=путь — сохранить лист инструкции как HTML
if (process.env.GUIDE_HTML) {
  const g = ss.getSheetByName(process.env.GUIDE_SHEET || '📋 Инструкция (менеджеры)');
  const covered = {};
  (g.mergeList || []).forEach(([r, c, nr, nc]) => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (i || j) covered[(r + i) + ':' + (c + j)] = true; });
  const span = (r, c) => { const m = (g.mergeList || []).find(x => x[0] === r && x[1] === c); return m ? m[3] : 1; };
  let html = '<table style="border-collapse:collapse;font:10pt Arial;table-layout:fixed">';
  html += '<colgroup>' + [1, 2, 3, 4, 5, 6, 7].map(c => `<col style="width:${(g.colW || {})[c] || 100}px">`).join('') + '</colgroup>';
  for (let r = 1; r <= g.getLastRow() + 1; r++) {
    html += `<tr style="height:${(g.rowH || {})[r] || 21}px">`;
    for (let c = 1; c <= 7; c++) {
      if (covered[r + ':' + c]) continue;
      const cell = g.peek(r, c), f = cell.f || {};
      let text = String(cell.v ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;');
      if (cell.rt && cell.rt.styles.length) text = text; // цвета легенды — упрощённо
      const inCard = (g.cards || []).find(([cr, cc, nr, nc]) => r >= cr && r < cr + nr && c >= cc && c < cc + nc);
      const b = inCard ? `border-left:${c === inCard[1] ? '1px solid #DADDE2' : '0'};border-right:${c + span(r, c) - 1 === inCard[1] + inCard[3] - 1 ? '1px solid #DADDE2' : '0'};border-top:${r === inCard[0] ? '1px solid #DADDE2' : '0'};border-bottom:${r === inCard[0] + inCard[2] - 1 || f.underline ? '1px solid #DADDE2' : '0'};` : '';
      html += `<td colspan="${span(r, c)}" style="padding:2px 8px;overflow:hidden;background:${f.bg || '#fff'};color:${f.color || '#1F2328'};font-weight:${f.bold ? 700 : 400};font-size:${f.size || 10}pt;text-align:${f.align || 'left'};${b}">${text}</td>`;
    }
    html += '</tr>';
  }
  fs.writeFileSync(process.env.GUIDE_HTML, html + '</table>');
}
console.log(`\n${passed} проверок пройдено`);
