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
// «Журнал» убран: ни одна операция не должна в него писать
const logRows0 = log() ? log().getLastRow() : 0;
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
   'showAddTranslatorTaskDialog', 'showSearchEditTranslatorSidebar', 'showTranslatorReportSidebar', 'migrateFromOldTable'].forEach(f => G[f]());
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
  assert.equal(sh.getRange(1, 18).getValue(), 'Жалобы на заказчика', 'колонка R «Жалобы»');
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
  assert.equal(r[2], '[LIT-26-2232][LogrusIT][ka][kk][Магазинка] Новые строчки (ссылка 2)');
  assert.equal(r[3].getDate(), 1); assert.equal(r[3].getMonth(), 9);
  assert.equal(r[15], 3);
  same(G.linksFromRich_(sh.getRange(2, 3).getRichTextValue()), { link: 'https://band/a', link2: 'https://band/b' });
  assert.ok(sh.getRange(2, 5).getDataValidation(), 'списки скопированы в новую строку');
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

test('Несколько доп. ссылок: «(ссылка 2)», «(ссылка 3)» после темы', () => {
  const id = G.submitNewTaskFromDialog({ contractor: 'LogrusIT', subject: 'Три переписки', link: 'https://band/m',
    link2: 'https://band/x\nhttps://band/y  https://band/z', languages: [], manager: 'Анастасия Лисовая' });
  const row = rowOf(id);
  const rt = tasks().getRange(row, 3).getRichTextValue();
  assert.match(rt.getText(), /Три переписки \(ссылка 2\) \(ссылка 3\) \(ссылка 4\)$/);
  same(G.linksFromRich_(rt), { link: 'https://band/m', link2: 'https://band/x\nhttps://band/y\nhttps://band/z' });
  const t = G.getTaskForEdit(row);
  assert.ok(!/ссылка \d/.test(t.subject), t.subject);
  assert.equal(G.findDuplicateTasks({ subject: 'другое', link: 'https://band/y' }).length, 1, 'дубль и по доп. ссылке');
  // Тема для копирования — без «(ссылка N)»: окно правки, расширение, отчёты
  const look = JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ token: 'T', action: 'lookup', task: id }) } }));
  assert.ok(look.found && !/ссылка \d/.test(look.subject), look.subject);
  const line = G.getManagerReport('Анастасия Лисовая', '').groups.flatMap(g => g.lines).find(l => l.link === 'https://band/m');
  assert.ok(line && !/ссылка \d/.test(line.subject) && line.link2.split('\n').length === 3, JSON.stringify(line));
  G.saveTaskEdits({ ...t, link2: 'https://band/y' });
  const rt2 = tasks().getRange(row, 3).getRichTextValue();
  assert.match(rt2.getText(), /Три переписки \(ссылка 2\)$/);
  same(G.linksFromRich_(rt2), { link: 'https://band/m', link2: 'https://band/y' });
  // Прежний вид «(доп. ссылка)» читается и при правке переходит в новый
  const old = G.getTaskForEdit(rowOf('LIT-1'));
  G.saveTaskEdits({ ...old, link2: old.link2 + '\nhttps://band/three' });
  const rt3 = tasks().getRange(rowOf('LIT-1'), 3).getRichTextValue();
  assert.match(rt3.getText(), / \(ссылка 2\) \(ссылка 3\)$/);
  assert.ok(!rt3.getText().includes('доп. ссылка'));
  same(G.linksFromRich_(rt3).link2.split('\n'), ['https://band/two', 'https://band/three']);
  G.saveTaskEdits({ ...G.getTaskForEdit(rowOf('LIT-1')), link2: 'https://band/two' }); // как было для следующих проверок
  G.deleteTask(row, id, tasks().getRange(row, 3).getValue());
});

test('Жалобы на заказчика: добавление, правка, старые окна не стирают', () => {
  const id = G.submitNewTaskFromDialog({ contractor: 'LogrusIT', subject: 'С жалобой', languages: [], complaints: 'Прислал задачу после 19:00' });
  const row = rowOf(id);
  assert.equal(tasks().getRange(row, 18).getValue(), 'Прислал задачу после 19:00');
  const t = G.getTaskForEdit(row);
  assert.equal(t.complaints, 'Прислал задачу после 19:00');
  G.saveTaskEdits({ ...t, complaints: 'Не отвечает в Band' });
  assert.equal(G.getTaskForEdit(row).complaints, 'Не отвечает в Band');
  const old = { ...G.getTaskForEdit(row) }; delete old.complaints; // окно старой версии поля не знает
  G.saveTaskEdits({ ...old, comment: 'ок' });
  assert.equal(tasks().getRange(row, 18).getValue(), 'Не отвечает в Band');
  G.deleteTask(row, id, tasks().getRange(row, 3).getValue());
});

test('Очистка поля тоже попадает в журнал', () => {
  const t = G.getTaskForEdit(rowOf('LIT-1'));
  G.saveTaskEdits({ ...t, comment: '' });
  assert.equal(G.getTaskForEdit(rowOf('LIT-1')).comment, '');
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
  assert.equal(ctx.SCRIPT_LOCK_HELD, false); assert.ok(!ctx.__lockHeld);
  // Как у Насти: openById без разрешения — скрипт в таблице, поэтому он не нужен
  const openById = ctx.SpreadsheetApp.openById;
  ctx.SpreadsheetApp.openById = () => { throw new Error('У вас нет разрешения на вызов функции "SpreadsheetApp.openById"'); };
  const look = call({ action: 'lookup', task: add.id });
  ctx.SpreadsheetApp.openById = openById;
  assert.ok(look.found, JSON.stringify(look));
  const w = call({ action: 'write', task: add.id, total: 1234.5, link: 'https://disk/x' });
  assert.ok(w.ok, JSON.stringify(w));
  assert.equal(tasks().getRange(rowOf(add.id), 12).getValue(), 1234.5);
  // В ячейке сметы — сам адрес ссылкой (таблица читает его мгновенно); повторная запись видит старую ссылку
  const r = rowOf(add.id);
  assert.equal(tasks().getRange(r, 11).getValue(), 'https://disk/x');
  assert.equal(G.linksFromRich_(tasks().getRange(r, 11).getRichTextValue()).link, 'https://disk/x');
  const again = call({ action: 'write', task: add.id, total: 1, link: 'https://disk/y' });
  same([again.error, again.link], ['exists', 'https://disk/x']);
  assert.equal(call({ action: 'lookup', task: add.id }).link, 'https://disk/x');
  // Окно правки видит адрес, а сохранение без изменений не пишет в журнал «Смета»
  const t = G.getTaskForEdit(r);
  assert.equal(t.estimateLink, 'https://disk/x');
  G.saveTaskEdits(t);
  assert.equal(G.getTaskForEdit(r).estimateLink, 'https://disk/x');
  G.saveTaskEdits({ ...t, estimateLink: 'https://disk/z' });
  same([tasks().getRange(r, 11).getValue(), G.getTaskForEdit(r).estimateLink], ['https://disk/z', 'https://disk/z']);
  // Адрес, вписанный руками, так и остаётся адресом
  tasks().getRange(r, 11).setValue('https://disk/hand');
  G.onEdit({ range: tasks().getRange(r, 11), source: ss, oldValue: 'Смета', value: 'https://disk/hand' });
  same([tasks().getRange(r, 11).getValue(), G.getTaskForEdit(r).estimateLink], ['https://disk/hand', 'https://disk/hand']);
  // Прежний вид («Ссылка на смету …», «Смета [номер]» — адрес под текстом) переделывается в адрес («Оформить таблицу»)
  tasks().getRange(r, 11).setRichTextValue(new RichBuilder().setText('Ссылка на смету ' + add.id).setLinkUrl('https://disk/old').build());
  const other = r + 1, otherWas = tasks().getRange(other, 11).getValue();
  G.estimateLabelsToUrls_(tasks());
  same([tasks().getRange(r, 11).getValue(), G.getTaskForEdit(r).estimateLink], ['https://disk/old', 'https://disk/old']);
  assert.equal(tasks().getRange(other, 11).getValue(), otherWas, 'остальные ячейки не тронуты');
  tasks().getRange(r, 11).setRichTextValue(new RichBuilder().setText('Смета [' + add.id + ']').setLinkUrl('https://disk/old2').build());
  assert.equal(G.estimateLabelsToUrls_(tasks()), 1);
  assert.equal(tasks().getRange(r, 11).getValue(), 'https://disk/old2');
  assert.equal(G.estimateLabelsToUrls_(tasks()), 0);
});

test('Расширение: удалить задачу', () => {
  const call = b => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ token: 'T', ...b }) } }));
  const add = call({ action: 'addTask', user: 'Настя', task: { contractor: 'LogrusIT', subject: 'Удалить меня', languages: ['Грузинский'] } });
  assert.ok(add.ok, JSON.stringify(add));
  const r = rowOf(add.id), n = tasks().getLastRow();
  const del = call({ action: 'deleteTask', row: r, id: add.id, origSubject: tasks().getRange(r, 3).getValue() });
  assert.ok(del.ok, JSON.stringify(del));
  assert.equal(tasks().getLastRow(), n - 1);
  assert.equal(call({ action: 'lookup', task: add.id }).found, false);
  const again = call({ action: 'deleteTask', row: r, id: add.id, origSubject: 'Удалить меня' });
  assert.ok(!again.ok && /не найдена/.test(again.error), JSON.stringify(again));
  assert.ok(!ctx.__lockHeld);
});

test('Фильтры', () => {
  G.hideClosedTasks(); same(tasks().filter.criteria[10].hidden, ['Отдано', 'Отменено']);
  G.showMyTasks(); assert.ok(!tasks().filter.criteria[14].hidden.includes('Анастасия Лисовая'));
  G.showAllTasks(); same(tasks().filter.criteria, {});
  G.showInProgressTranslatorTasks(); G.showAllTranslatorTasks();
});

test('Мои задачи: с Sheets API — личное представление, без него — общий фильтр', () => {
  const reqs = [];
  let existing = [];
  ctx.Sheets = { Spreadsheets: {
    get: () => ({ sheets: [{ properties: { sheetId: tasks().getSheetId() }, filterViews: existing }] }),
    batchUpdate: (body, id) => {
      reqs.push(body.requests[0]);
      return { replies: [{ addFilterView: { filterView: { filterViewId: 777 } } }] };
    } } };
  try {
    G.showAllTasks();
    G.showMyTasks();
    const add = reqs[0].addFilterView.filter;
    same([add.title, add.filterSpecs[0].columnIndex, add.filterSpecs[0].filterCriteria.condition.values[0].userEnteredValue],
      ['Задачи: Анастасия Лисовая', 13, 'Анастасия Лисовая']);
    assert.match(calls.alerts.at(-1), /Режимы фильтрации/);
    same(tasks().filter && tasks().filter.criteria, {}, 'общий фильтр не тронут');
    G.showMyTasks(); // второй раз — то же представление обновляется, а не плодится
    assert.equal(reqs[1].updateFilterView.filter.filterViewId, 777);
    // Представление уже есть в таблице, а номер не запомнен — берём его, а не создаём второе с тем же названием
    existing = [{ filterViewId: 555, title: 'Задачи: Ольга Шешина' }];
    ctx.__prompt = 'Ольга Шешина'; G.filterByManager(); ctx.__prompt = '';
    assert.equal(reqs.at(-1).updateFilterView.filter.filterViewId, 555);
    // Сразу для всех менеджеров из «Списков»
    const before = reqs.length;
    G.createAllManagerViews();
    assert.equal(reqs.length - before, G.getListsData().managers.length);
    assert.match(calls.alerts.at(-1), /Готово: \d+/);
  } finally { delete ctx.Sheets; }
  G.showMyTasks(); // сервис не включён — как раньше, общий фильтр
  assert.ok(!tasks().filter.criteria[14].hidden.includes('Анастасия Лисовая'));
  G.showAllTasks();
});

test('Правка прямо в листе: цветные языки, новая строка получает номер и списки', () => {
  const sh = tasks();
  const row = rowOf('LIT-7');
  const old = sh.getRange(row, 10).getValue();
  sh.getRange(row, 10).setValue('Холд');
  G.onEdit({ range: sh.getRange(row, 10), oldValue: old, value: 'Холд' });

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

  // Пока ждали очереди за номером, сверху добавили задачу — номер уходит в сдвинутую строку, а не в чужую
  const r2 = sh.getLastRow() + 1;
  sh.getRange(r2, 3).setValue('Ещё руками');
  sh.getRange(r2, 13).setValue('LogrusIT');
  const realLock = ctx.LockService;
  let top = '';
  ctx.LockService = { getScriptLock: () => ({ waitLock() {
    ctx.LockService = realLock; // пока мы ждём, коллега добавляет задачу через расширение
    top = G.submitNewTaskFromDialog({ contractor: 'LogrusIT', subject: 'Сверху', languages: [] });
    realLock.getScriptLock().waitLock();
  }, releaseLock() { realLock.getScriptLock().releaseLock(); } }) };
  G.onEdit({ range: sh.getRange(r2, 13), value: 'LogrusIT' });
  ctx.LockService = realLock;
  same([sh.getRange(r2 + 1, 3).getValue(), sh.getRange(r2, 3).getValue()], ['Ещё руками', 'Задача руками']);
  assert.match(String(sh.getRange(r2 + 1, 1).getValue()), /^LIT-26-\d+$/);
  assert.notEqual(sh.getRange(r2 + 1, 1).getValue(), top, 'номера разные');
  assert.notEqual(sh.getRange(r2, 1).getValue(), sh.getRange(r2 + 1, 1).getValue(), 'чужой номер не перезаписан');

  // Дедлайн сам из срока сдачи: от даты получения
  sh.getRange(r, 4).setValue(new CDate(2026, 9, 1));
  sh.getRange(r, 9).setValue(new CDate(2026, 9, 6));
  G.onEdit({ range: sh.getRange(r, 9) });
  assert.equal(sh.getRange(r, 8).getValue(), 'До недели');
  sh.getRange(r, 9).setValue(new CDate(2026, 9, 2));
  G.onEdit({ range: sh.getRange(r, 9) });
  assert.equal(sh.getRange(r, 8).getValue(), '1-2 дня');
  sh.getRange(r, 4).setValue(new CDate(2026, 8, 1));
  G.onEdit({ range: sh.getRange(r, 4) });
  assert.equal(sh.getRange(r, 8).getValue(), 'Месяц и больше');
  sh.getRange(r, 8).setValue('Холд'); // «Холд» руками — не перебиваем
  sh.getRange(r, 9).setValue(new CDate(2026, 8, 1));
  G.onEdit({ range: sh.getRange(r, 9) });
  assert.equal(sh.getRange(r, 8).getValue(), 'Холд');
  const auto = G.submitNewTaskFromDialog({ contractor: 'LogrusIT', subject: 'С дедлайном', date: '2026-10-01', exactDeadline: '2026-10-01', languages: [] });
  assert.equal(sh.getRange(rowOf(auto), 8).getValue(), 'ASAP');
  const manual = G.submitNewTaskFromDialog({ contractor: 'LogrusIT', subject: 'Свой дедлайн', date: '2026-10-01', exactDeadline: '2026-10-01', deadline: 'До недели', languages: [] });
  assert.equal(sh.getRange(rowOf(manual), 8).getValue(), 'До недели');

  // Переводчики
  const t = tr(); const tr2 = t.getLastRow() + 1;
  t.getRange(tr2, 4).setValue('Руками');
  G.onEdit({ range: t.getRange(tr2, 4), value: 'Руками' });
  assert.ok(t.getRange(tr2, 8).getDataValidation());

  // «Списки» — обновляются цвета, без ошибок
  G.onEdit({ range: ss.getSheetByName('Списки').getRange('E20') });
  assert.equal(ctx.SCRIPT_LOCK_HELD, false);
});

test('Шаблоны задач: свои у каждого, большой список не теряется', () => {
  same(G.getTaskTemplates(), []);
  const many = Array.from({ length: 60 }, (_, i) => ({ name: 'Шаблон ' + i, contractor: 'LogrusIT', ticket: 'LOCAL-1116',
    subject: 'Новые строчки для приложения от 22.09 — '.repeat(5) + i, languages: ['Грузинский', 'Иврит'], evil: '<x>' }));
  G.saveTaskTemplates(many.concat([{ name: '' }, { name: 'плохие языки', languages: 'нет' }]));
  const got = G.getTaskTemplates();
  assert.equal(got.length, 61);
  assert.equal(got[59].subject.slice(-2), '59');
  assert.ok(!('evil' in got[0]));
  same(got[60].languages, []);
  G.saveTaskTemplates(got.slice(0, 1)); // старые куски удаляются
  assert.equal(G.getTaskTemplates().length, 1);
  assert.equal(G.getAddTaskFormLists().templates[0].name, 'Шаблон 0');
  // Редкие языки получили коды; код из «Списков» (колонка I) главнее
  const codes = G.getAddTaskFormLists().langCodes;
  same([codes['Иврит'], codes['Турецкий'], codes['Корейский'], codes['Армянский']], ['he', 'tr', 'ko', 'hy']);
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
  // Фильтр: отданные / отменённые / все; счётчики сходятся
  const done = call({ action: 'myTasks', manager: 'Анастасия Лисовая', filter: 'done' });
  assert.ok(done.tasks.some(x => x.id === t.id) && done.tasks.every(x => x.status === 'Отдано'));
  assert.ok(done.tasks.every(x => !x.overdue), 'у закрытых нет «просрочено»');
  const all = call({ action: 'myTasks', manager: 'Анастасия Лисовая', filter: 'all' });
  // Ссылки читаются только для показанных строк — но у каждой задачи та же, что в таблице
  all.tasks.forEach(x => assert.equal(x.link, G.linksFromRich_(tasks().getRange(x.row, 3).getRichTextValue()).link, x.id));
  assert.ok(all.tasks.some(x => x.link), 'хотя бы у одной задачи есть ссылка');
  const c = all.counts;
  assert.equal(c.all, all.tasks.length + all.more);
  assert.ok(c.open + c.done + c.cancelled <= c.all && c.done >= 1, JSON.stringify(c));
  assert.equal(after.overdue, done.overdue, 'бейдж просрочек не зависит от фильтра');
  const dates = all.tasks.map(x => x.date).filter(Boolean);
  assert.ok(dates.length, 'у задач есть дата');
  assert.equal(call({ action: 'myTasks', manager: 'Анастасия Лисовая', filter: 'cancelled' }).tasks.filter(x => x.status !== 'Отменено').length, 0);
  const ids = call({ action: 'previewTaskIds' });
  assert.ok(ids.ok && ids.ids.LogrusIT === G.previewNextTaskId('LogrusIT'), JSON.stringify(ids));
  const dup = call({ action: 'checkDuplicates', task: { subject: 'x', link: 'https://band/lit4' } });
  assert.equal(dup.duplicates.length, 1);
  // Правка из расширения: поиск → задача целиком → сохранение (в журнал — от расширения)
  same(call({ action: 'searchTasks', query: '' }).tasks, []);
  const f = call({ action: 'searchTasks', query: 'LIT-1' });
  assert.ok(f.ok && f.tasks.some(x => x.id === 'LIT-1'), JSON.stringify(f).slice(0, 200));
  const hit = f.tasks.find(x => x.id === 'LIT-1');
  const got = call({ action: 'getTask', row: hit.row + 5, id: 'LIT-1' }); // строка сдвинулась — найдёт по номеру
  assert.ok(got.ok && got.task.id === 'LIT-1' && got.task.link === 'https://band/one', JSON.stringify(got).slice(0, 200));
  const saved = call({ action: 'saveTask', user: 'Настя', task: { ...got.task, sp: 7, link2: 'https://band/two\nhttps://band/three' } });
  assert.ok(saved.ok, JSON.stringify(saved));
  const edited = call({ action: 'getTask', row: got.task.row, id: 'LIT-1' }).task;
  same([edited.sp, edited.link2], [7, 'https://band/two\nhttps://band/three']);
  const bad = call({ action: 'saveTask', task: { ...edited, link2: 'не ссылка' } });
  assert.ok(!bad.ok && /http/.test(bad.error));
  call({ action: 'saveTask', task: { ...edited, link2: 'https://band/two' } }); // как было
});

// Файл Excel старой таблицы: OLD_XLSX=путь, XLSX_MODULE=путь к пакету xlsx (SheetJS), как в окне переноса
function xlsxSource() {
  const XLSX = require(process.env.XLSX_MODULE || 'xlsx');
  const html = fs.readFileSync(__dirname + '/../MigrateDialog.html', 'utf8');
  const code = html.split('// xlsx-parse:start')[1].split('// xlsx-parse:end')[0];
  const parse = new Function('XLSX', 'wb', code + '\nreturn oldSheetFromWorkbook(XLSX, wb);');
  const wb = XLSX.read(new Uint8Array(fs.readFileSync(process.env.OLD_XLSX)), { type: 'array', cellNF: true, cellDates: false });
  // Как через google.script.run: только простые данные
  return JSON.parse(JSON.stringify(parse(XLSX, wb)));
}

// Файл в формате новой таблицы (TASKS_XLSX): задачи с сентября + лист «Данные до сентября 2026»
test('Перенос из файла с колонками новой таблицы: задачи и архив', { skip: !process.env.TASKS_XLSX }, () => {
  const XLSX = require(process.env.XLSX_MODULE || 'xlsx');
  const html = fs.readFileSync(__dirname + '/../MigrateDialog.html', 'utf8');
  const code = html.split('// xlsx-parse:start')[1].split('// xlsx-parse:end')[0];
  const parse = new Function('XLSX', 'wb', code + '\nreturn tasksFileFromWorkbook(XLSX, wb);');
  const wb = XLSX.read(new Uint8Array(fs.readFileSync(process.env.TASKS_XLSX)), { type: 'array', cellNF: true, cellDates: false });
  const src = JSON.parse(JSON.stringify(parse(XLSX, wb)));
  same(src.sheets.map(x => [x.name, x.archive]), [['С сентября 2026', false], ['Данные до сентября 2026', true]]);
  const before = tasks().getLastRow();
  const p = G.migrationPreviewFromTasksFile(src);
  assert.ok(p.main > 300 && p.archives[0].count > 3000, JSON.stringify(p));
  const r = G.migrationApplyFromTasksFile(src);
  assert.equal(tasks().getLastRow() - before, r.added);
  const arch = ss.getSheetByName('Данные до сентября 2026');
  assert.equal(arch.getLastRow() - 1, r.archived);
  // Ссылки: тема — Band, «(ссылка 2)» из комментария, смета — «Ссылка на смету …»
  const row = rowOf('LIT-26-2188');
  const rt = tasks().getRange(row, 3).getRichTextValue();
  assert.equal(G.linksFromRich_(rt).link, 'https://band.wb.ru/wb/pl/tpw8mh66tpd7ung4siq9g96hoh');
  assert.equal(G.linksFromRich_(rt).link2.split('\n').length, 2, rt.getText());
  assert.ok(!/Доп\. ссылка|Ещё ссылки/.test(tasks().getRange(row, 17).getValue()));
  same([tasks().getRange(row, 11).getValue(), G.getTaskForEdit(row).estimateLink], ['https://disk.wb.ru/f/83009693', 'https://disk.wb.ru/f/83009693']);
  assert.ok(tasks().getRange(row, 4).getValue() instanceof CDate);
  // Повторно — ничего не задваивается
  const again = G.migrationPreviewFromTasksFile(src);
  same([again.main, again.archives[0].count], [0, 0]);
});

test('Перенос из файла Excel даёт то же, что по ссылке', { skip: !oldFixture || !process.env.OLD_XLSX }, () => {
  const byLink = G.planMigration_(others[OLD_ID]);
  const byFile = G.planFromSource_(G.sourceFromUpload_(xlsxSource()));
  // В выгрузке для теста вместо значения формулы — её текст (=29440+…); Google и Excel дают число
  const formulas = byLink.rows.map(r => r.values.map(v => /^=/.test(v)));
  const norm = plan => plan.rows.map((r, i) => JSON.stringify([r.link, r.link2, r.values.map((v, j) =>
    formulas[i] && formulas[i][j] ? '=' : v instanceof CDate ? 'D' + v.getFullYear() + '-' + v.getMonth() + '-' + v.getDate() : String(v))]));
  const a = norm(byLink), b = norm(byFile);
  assert.equal(b.length, a.length, 'строк: ' + b.length + ' против ' + a.length);
  const diff = a.map((x, i) => x === b[i] ? null : i).filter(i => i !== null);
  assert.equal(diff.length, 0, diff.length + ' различий, первое:\n' + a[diff[0]] + '\n' + b[diff[0]]);
  same(G.migrationPreviewFromData(xlsxSource()).count, byLink.rows.length);
});

test('Перенос истории из старой таблицы', { skip: !oldFixture }, () => {
  const before = tasks().getLastRow();
  const url = 'https://docs.google.com/spreadsheets/d/' + OLD_ID + '/edit#gid=0';
  const preview = G.migrationPreviewFromLink(url);
  const res = G.migrationApplyFromLink(url);
  const added = tasks().getLastRow() - before;
  assert.ok(added > 3700 && res.added === added && preview.count === added, 'перенесено ' + added + '; ' + JSON.stringify(preview));
  const rows = tasks().getRange(1, 1, tasks().getLastRow(), 17).getValues();
  // Одна задача подрядчику на три строки (разные запросы из Band): номер и тема у всех, сумма — только у первой
  const g = rows.filter(r => r[0] === 'LIT-26-1953');
  assert.equal(g.length, 3);
  assert.ok(g.every(r => /^\[LIT-26-1953\]/.test(r[2])));
  assert.equal(g.filter(r => r[11] !== '').length, 1, 'сумма не задвоилась');
  assert.ok(g.every(r => r[1] === 'LOCAL-493'), 'тикет из колонки «Задача»');
  assert.ok(g.every(r => r[4] === 'WBP'), 'сторона → продукт по памятке');
  // Три ссылки в одной ячейке: первая — на тему, остальные — «(ссылка 2)», «(ссылка 3)»
  const multi = rows.findIndex(r => r[0] === 'LIT-26-2188') + 1;
  same(G.linksFromRich_(tasks().getRange(multi, 3).getRichTextValue()),
    { link: 'https://band.wb.ru/wb/pl/tpw8mh66tpd7ung4siq9g96hoh', link2: 'https://band.wb.ru/wb/pl/tjfa9pqwnprk3eorhzjinzkhjy\nhttps://band.wb.ru/wb/pl/mytom1ojftg3789opd3c7ioe1o' });
  assert.match(tasks().getRange(multi, 3).getValue(), / \(ссылка 2\) \(ссылка 3\)$/);
  // «Жалобы на заказчика» — в свою колонку, а не в комментарий
  const complained = rows.findIndex(r => r[0] === 'LIT-25-159') + 1;
  same([tasks().getRange(complained, 18).getValue(), /Жалобы/.test(tasks().getRange(complained, 17).getValue())], ['Прислал задачу после 19:00', false]);
  // Смета из старой таблицы — тоже «Ссылка на смету номер» со ссылкой
  same([tasks().getRange(multi, 11).getValue(), G.getTaskForEdit(multi).estimateLink], ['https://disk.wb.ru/f/83009693', 'https://disk.wb.ru/f/83009693']);
  assert.ok(!/Ещё ссылки/.test(tasks().getRange(multi, 17).getValue()));
  const firstOld = rows.findIndex(r => r[0] === 'LIT-25-1') + 1;
  assert.equal(G.linksFromRich_(tasks().getRange(firstOld, 3).getRichTextValue()).link.slice(0, 26), 'https://band.wb.ru/wb/pl/g');
  // Весь лист по дате, свежие сверху (без дат — внизу)
  const dates = rows.slice(1).map(r => r[3]).filter(d => d instanceof CDate).map(d => d.getTime());
  assert.ok(dates.every((t, i) => i === 0 || dates[i - 1] >= t), 'даты по убыванию');
  const oldest = rows.slice(1).filter(r => r[3] instanceof CDate).pop();
  assert.equal(oldest[3].getTime(), Math.min(...dates), 'самая старая дата — внизу');
  // Повторный запуск ничего не задваивает
  assert.equal(G.migrationPreviewFromLink(url).count, 0);
  assert.equal(G.migrationApplyFromLink(url).added, 0);
  assert.equal(tasks().getLastRow() - before, added);
  // Правка второй строки общего номера попадает во вторую, а не в первую
  const second = rows.findIndex((r, i) => r[0] === 'LIT-26-1953' && i > rows.findIndex(x => x[0] === 'LIT-26-1953')) + 1;
  const t = G.getTaskForEdit(second);
  G.saveTaskEdits({ ...t, comment: 'вторая строка' });
  assert.equal(tasks().getRange(second, 17).getValue(), 'вторая строка');
  assert.notEqual(tasks().getRange(second - 1, 17).getValue(), 'вторая строка');
  // Номера новых задач продолжаются после перенесённых
  const next = G.generateNextTaskId(tasks(), 'LogrusIT');
  assert.ok(Number(next.split('-')[2]) > 2231 && !rows.some(r => r[0] === next), next);
  console.log('      предпросмотр:', JSON.stringify(preview));
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
test('Суммы смет: без округления, прежняя сумма — в заметке', () => {
  same([G.rub_(102641.64), G.rub_(1220.005), G.rub_(0.1 + 0.2), G.rub_('1 000 ₽')], ['102 641,64', '1 220,005', '0,30', '1 000 ₽']);
  const call = b => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ token: 'T', ...b }) } }));
  const id = G.submitNewTaskFromDialog({ contractor: 'LogrusIT', subject: 'Сумма 1', manager: 'Тест Сумма', date: '2026-09-15', languages: ['Грузинский'] });
  const r = () => rowOf(id), cell = () => tasks().getRange(r(), 12);
  assert.ok(call({ action: 'write', task: id, total: 102641.64, link: 'https://disk/a' }).ok);
  same([cell().getValue(), cell().getNumberFormat(), cell().getNote()], [102641.64, '#,##0.00', '']);
  // Подрядчик обновил смету: сумма и ссылка заменяются, прежние — в заметке
  const w = call({ action: 'write', task: id, total: 105210.3, link: 'https://disk/b', version: 2, overwrite: true });
  same([w.ok, w.was, cell().getValue()], [true, 102641.64, 105210.3]);
  assert.match(cell().getNote(), /смета обновлена \(версия 2\): было 102 641,64 ₽ → стало 105 210,30 ₽\nпрежняя смета: https:\/\/disk\/a/);
  // Доли копейки не теряются и видны в таблице
  call({ action: 'write', task: id, total: 1220.005, link: 'https://disk/b', overwrite: true });
  same([cell().getValue(), cell().getNumberFormat()], [1220.005, '#,##0.000']);
  assert.match(cell().getNote(), /было 105 210,30 ₽ → стало 1 220,005 ₽$/);
  // Та же смета ещё раз — заметка не растёт
  const note = cell().getNote();
  call({ action: 'write', task: id, total: 1220.005, link: 'https://disk/b', overwrite: true });
  assert.equal(cell().getNote(), note);
  // Правка руками — тоже в заметку
  G.saveTaskEdits({ ...G.getTaskForEdit(r()), total: '1000' });
  same([cell().getValue(), cell().getNumberFormat()], [1000, '#,##0.00']);
  assert.match(cell().getNote(), /сумма изменена вручную: было 1 220,005 ₽ → стало 1 000,00 ₽$/);
  G.saveTaskEdits({ ...G.getTaskForEdit(r()), comment: 'без смены суммы' });
  assert.match(cell().getNote(), /стало 1 000,00 ₽$/);
});

test('Отчёты по дате поступления и по дате закрытия (срок сдачи)', () => {
  const id = G.submitNewTaskFromDialog({ contractor: 'LogrusIT', subject: 'Пришла в марте, закрыта в апреле', manager: 'Период Тест',
    date: '2031-03-28', exactDeadline: '2031-04-02', total: '100', languages: [] });
  const lines = (m, by) => G.getManagerReport('Период Тест', m, by).lineCount;
  same([lines('2031-03'), lines('2031-04'), lines('2031-03', 'due'), lines('2031-04', 'due')], [1, 0, 0, 1]);
  same([G.getCustomReport('2031', '3').totalTasks, G.getCustomReport('2031', '4', 'due').totalTasks, G.getCustomReport('2031', '3', 'due').totalTasks], [1, 1, 0]);
  assert.match(G.getCustomReport('2031', '4', 'due').periodLabel, /Апрель 2031 \(по дате закрытия\)/);
  same([G.getDashboardData('2031', '3').totalTasks, G.getDashboardData('2031', '4', 'due').totalTasks], [1, 1]);
  assert.ok(G.getDashboardYears().indexOf(2031) !== -1);
  assert.ok(id);
});

test('Журнал больше не пишется', () => {
  assert.equal(log() ? log().getLastRow() : 0, logRows0);
  assert.equal(typeof G.logChange, 'undefined');
});
console.log(`\n${passed} проверок пройдено`);
