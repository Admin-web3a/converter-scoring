/**
 * Бэкенд для страницы скоринга гипотез конвертеров.
 *
 * Публикуется как веб-приложение: Deploy → New deployment → Web app,
 * Execute as: Me, Who has access: Anyone.
 * Полученный /exec URL вставляется в docs/index.html в константу API_URL.
 */

var SHEET_NAME = 'scores';
var HEADERS = [
  'Отправлено',
  'Кто',
  'ID гипотезы',
  'Гипотеза',
  'Impact',
  'Confidence',
  'Effort',
  'TimeToSignal',
  'Комментарий',
];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

/** Возвращает все сохранённые оценки. */
function doGet() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getSheet_();
    var last = sheet.getLastRow();
    if (last < 2) return json_({ ok: true, rows: [] });

    var values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
    var rows = values
      .filter(function (r) {
        return r[1] && r[2];
      })
      .map(function (r) {
        return {
          submittedAt: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
          person: String(r[1]),
          hypothesisId: String(r[2]),
          hypothesisName: String(r[3]),
          impact: Number(r[4]),
          confidence: Number(r[5]),
          effort: Number(r[6]),
          signal: Number(r[7]),
          comment: r[8] ? String(r[8]) : '',
        };
      });
    return json_({ ok: true, rows: rows });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Сохраняет оценки одного человека. Предыдущие строки этого же человека
 * удаляются, поэтому в таблице всегда лежит последняя версия — по одной
 * строке на пару «человек + гипотеза».
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var payload = JSON.parse(e.postData.contents);
    var person = String(payload.person || '').trim();
    var scores = payload.scores;

    if (!person) return json_({ ok: false, error: 'Не указано имя' });
    if (!scores || !scores.length) return json_({ ok: false, error: 'Пустой набор оценок' });

    var sheet = getSheet_();
    var last = sheet.getLastRow();

    if (last > 1) {
      var existing = sheet.getRange(2, 2, last - 1, 1).getValues();
      for (var i = existing.length - 1; i >= 0; i--) {
        if (String(existing[i][0]) === person) sheet.deleteRow(i + 2);
      }
    }

    var now = new Date();
    var rows = scores.map(function (s) {
      return [
        now,
        person,
        String(s.id),
        String(s.name || ''),
        Number(s.impact),
        Number(s.confidence),
        Number(s.effort),
        Number(s.signal),
        s.comment ? String(s.comment) : '',
      ];
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    return json_({ ok: true, saved: rows.length, person: person });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
