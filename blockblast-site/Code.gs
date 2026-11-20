/**
 * CrocoBlast — серверная часть для Google Apps Script.
 *
 * doGet()    — отдаёт игру (Index.html) как веб-приложение.
 * getTop()   — возвращает лучшие результаты из Google-таблицы.
 * addScore() — добавляет результат игрока в таблицу.
 *
 * При первом запуске таблица создаётся автоматически и её ID сохраняется
 * в Script Properties, чтобы при повторных вызовах использовалась та же таблица.
 */

var SHEET_TITLE = 'CrocoBlast — Leaderboard';
var SHEET_TAB = 'Leaderboard';
var PROP_SHEET_ID = 'crocoblast_sheet_id';
var HEADERS = ['Ник', 'Очки', 'Дата'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CrocoBlast — играй и попади в лидерборд')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Создаёт или открывает таблицу лидерборда. */
function getSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(PROP_SHEET_ID);

  if (id) {
    try {
      var existing = SpreadsheetApp.openById(id);
      var sheet = existing.getSheets()[0];
      ensureHeaders_(sheet);
      return sheet;
    } catch (err) {
      // Старый ID больше не работает — создадим новую таблицу.
    }
  }

  var created = SpreadsheetApp.create(SHEET_TITLE);
  var newSheet = created.getSheets()[0];
  newSheet.setName(SHEET_TAB);
  ensureHeaders_(newSheet);
  props.setProperty(PROP_SHEET_ID, created.getId());
  return newSheet;
}

/** Гарантирует, что в таблице есть правильная шапка. */
function ensureHeaders_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } else {
    var current = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    var same = true;
    for (var i = 0; i < HEADERS.length; i++) {
      if (current[i] !== HEADERS[i]) { same = false; break; }
    }
    if (!same) sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#1b2444')
    .setFontColor('#ffffff');
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 190);
}

/** Топ результатов. Возвращает массив объектов {name, score, date}. */
function getTop(limit) {
  var n = parseInt(limit, 10);
  if (!isFinite(n) || n < 1) n = 10;
  if (n > 50) n = 50;

  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];

  var data = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0] == null ? '' : data[i][0]).trim();
    var score = Number(data[i][1]);
    if (!name || !isFinite(score)) continue;
    rows.push({
      name: name.slice(0, 16),
      score: Math.max(0, Math.floor(score)),
      date: data[i][2] ? String(data[i][2]) : ''
    });
  }

  rows.sort(function (a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return String(a.date).localeCompare(String(b.date));
  });

  return rows.slice(0, n);
}

/** Добавляет результат и сразу возвращает обновлённый топ-10. */
function addScore(name, score) {
  var nick = String(name == null ? '' : name)
    .replace(/[\r\n\t<>]/g, '')
    .trim()
    .slice(0, 16) || 'Игрок';

  var pts = Math.floor(Number(score));
  if (!isFinite(pts) || pts < 0) pts = 0;
  if (pts > 1000000000) pts = 1000000000;

  var sheet = getSheet_();
  sheet.appendRow([nick, pts, new Date()]);
  return getTop(10);
}

/** Можно запустить вручную в редакторе, чтобы сразу создать таблицу. */
function setupSheet() {
  var sheet = getSheet_();
  return 'Таблица готова: ' + sheet.getParent().getUrl() + '\nID: ' + sheet.getParent().getId();
}

/** Для проверки: какой ID таблицы используется. */
function getSheetId() {
  return getSheet_().getParent().getId();
}
