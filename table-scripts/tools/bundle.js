// Собирает всё в один файл: node table-scripts/tools/bundle.js → table-scripts/ВсёВОдном.gs
// Код.gs + Перенос.gs + Smeta.gs и все окна (HTML) как строки — вставляется в Apps Script одним файлом.
'use strict';
const fs = require('fs'), path = require('path');
const dir = path.resolve(__dirname, '..');
const read = f => fs.readFileSync(path.join(dir, f), 'utf8');

const htmlNames = ['Общее', 'AddTaskDialog', 'SearchEditSidebar', 'DashboardSidebar', 'ManagerReportSidebar',
  'CustomReportSidebar', 'AddTranslatorTaskDialog', 'SearchEditTranslatorSidebar', 'TranslatorReportSidebar', 'MigrateDialog', 'ReconcileDialog'];
const html = htmlNames.map(n => `  ${JSON.stringify(n)}: ${JSON.stringify(read(n + '.html'))}`).join(',\n');

const out = `/*************************************************************
 * ВсёВОдном.gs — вся таблица задач в одном файле.
 *
 * Как поставить: в таблице «Расширения → Apps Script», удалить все старые файлы,
 * создать один файл-скрипт, вставить этот текст целиком, сохранить (Ctrl+S).
 * Один раз: «Сервисы» → «+» → Google Sheets API → «Добавить» — тогда «Мои задачи» личные.
 * Дальше — README, шаги 4–6 (токен, развёртывание, «Оформить таблицу»).
 *
 * Собран автоматически из Код.gs, Перенос.gs, Smeta.gs и HTML-окон (tools/bundle.js).
 * Править лучше исходные файлы и пересобирать.
 *************************************************************/

${read('Код.gs')}

${read('Перенос.gs')}

${read('../smeta-extension/apps-script/Smeta.gs')}

// ==================== ОКНА (HTML) ====================
// Текст всех окон — чтобы не создавать 9 отдельных HTML-файлов.
const HTML_FILES = {
${html}
};
`;
fs.writeFileSync(path.join(dir, 'ВсёВОдном.gs'), out);
console.log('ВсёВОдном.gs:', (out.length / 1024).toFixed(0), 'КБ');
