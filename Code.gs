const SPREADSHEET_ID = '1R6jOHFsM7MrUG-4Lv-9_Q7dCzA3KtZcIY6h3OT7m41o';
const SHEET_NAME = 'Чек-лист';

const CHECK_LABELS = [
  'Подготовить тетрадь',
  'Обсудить задачу с напарником',
  'Имя и отчество преподавателя',
  'Познакомиться с одногруппниками',
  'Вспомнить дроби',
  'Вспомнить решение уравнений',
  'Позиционные системы счисления',
  'Двоичная и другие системы счисления',
  'Домашние животные',
  'Занять две области другой команды',
  'Решить уравнение',
  'Zoom: реакции и отметки',
  'Лучшее решение на 4-м уроке',
  'Сломать мозг об тренажёр',
  'Обсудить занятия с родителями',
  'Выбрать любимое занятие',
  'Заполнить чек-лист'
];

const HEADERS = [
  'Прозвище',
  ...CHECK_LABELS,
  'Выполнено пунктов',
  'Всё выполнено',
  'Последнее изменение'
];

function sheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  const current = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (current.join('|') !== HEADERS.join('|')) sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  return sh;
}

function normalizeNick_(value) {
  return String(value || '').trim().toLocaleLowerCase();
}

function findRow_(sh, nickname) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const target = normalizeNick_(nickname);
  const names = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < names.length; i++) {
    if (normalizeNick_(names[i][0]) === target) return i + 2;
  }
  return 0;
}

function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    return saveChecklist_(data);
  } catch (err) {
    return json_({ok:false, error:String(err && err.message || err)});
  }
}

function doGet(e) {
  const params = e.parameter || {};
  const callback = String(params.callback || '').replace(/[^a-zA-Z0-9_.$]/g, '');
  let payload;

  if (params.action === 'save') {
    payload = saveChecklist_({
      nickname: params.nickname,
      checks: String(params.checks || '').split('').map(function(value) { return value === '1'; }),
      updatedAt: params.updatedAt
    }, true);
  } else {
    payload = loadChecklist_(params.nickname);
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}

function saveChecklist_(data, asObject) {
  try {
    const nickname = String(data.nickname || '').trim();
    if (!nickname) throw new Error('nickname is required');

    const checks = Array.from({length: CHECK_LABELS.length}, function(_, i) {
      return Boolean((data.checks || [])[i]);
    });
    const completed = checks.filter(Boolean).length;
    const ready = completed === checks.length;
    const updatedAt = data.updatedAt ? new Date(data.updatedAt) : new Date();

    const sh = sheet_();
    let row = findRow_(sh, nickname);
    if (!row) row = sh.getLastRow() + 1;

    sh.getRange(row, 1, 1, HEADERS.length).setValues([[
      nickname,
      ...checks,
      completed,
      ready,
      updatedAt
    ]]);

    const result = {ok:true, row:row, state:{checks:checks, updatedAt:updatedAt.toISOString()}};
    return asObject ? result : json_(result);
  } catch (err) {
    const result = {ok:false, error:String(err && err.message || err)};
    return asObject ? result : json_(result);
  }
}

function loadChecklist_(nicknameValue) {
  const nickname = String(nicknameValue || '').trim();
  let payload = {found:false};

  if (nickname) {
    const sh = sheet_();
    const row = findRow_(sh, nickname);
    if (row) {
      const values = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      payload = {
        found:true,
        state:{
          checks: values.slice(1, 1 + CHECK_LABELS.length).map(Boolean),
          updatedAt: values[HEADERS.length - 1] instanceof Date ? values[HEADERS.length - 1].toISOString() : String(values[HEADERS.length - 1] || '')
        }
      };
    }
  }

  return payload;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
