const SPREADSHEET_ID = '1R6jOHFsM7MrUG-4Lv-9_Q7dCzA3KtZcIY6h3OT7m41o';
const SHEET_NAME = 'Чек-лист';

const HEADERS = [
  'Прозвище',
  'Компьютер или ноутбук',
  'Микрофон и камера',
  'Камера настроена',
  'Ручка и черновик',
  'Интернет ≥ 5 МБайт/с',
  'Переключение между окнами',
  'Classkick с рабочего устройства',
  'Работа в пробной тетради Classkick',
  'Выполнено пунктов',
  'Готово',
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
    const nickname = String(data.nickname || '').trim();
    if (!nickname) throw new Error('nickname is required');

    const checks = Array.from({length: 8}, (_, i) => Boolean((data.checks || [])[i]));
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

    return json_({ok:true, row:row});
  } catch (err) {
    return json_({ok:false, error:String(err && err.message || err)});
  }
}

function doGet(e) {
  const nickname = String((e.parameter && e.parameter.nickname) || '').trim();
  const callback = String((e.parameter && e.parameter.callback) || '').replace(/[^a-zA-Z0-9_.$]/g, '');
  let payload = {found:false};

  if (nickname) {
    const sh = sheet_();
    const row = findRow_(sh, nickname);
    if (row) {
      const values = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      payload = {
        found:true,
        state:{
          checks: values.slice(1, 9).map(Boolean),
          updatedAt: values[11] instanceof Date ? values[11].toISOString() : String(values[11] || '')
        }
      };
    }
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
