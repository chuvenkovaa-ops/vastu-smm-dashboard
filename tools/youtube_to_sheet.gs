/**
 * YouTube → «Таблица для дашборда VASTU»
 *
 * Раз в день забирает статистику канала и пишет её в блок YOUTUBE
 * на листах «Еженедельная статистика» и «Ежемесячная статистика».
 *
 * Заполняет строки: Подписчики, Подписки, Отписки, Просмотры,
 * Время просмотра, ч, Вовлеченность (лайки + комментарии + репосты), Количество видео.
 * Строки «Прирост подписчиков» (формула), «Лиды» и GetCourse-строки не трогает.
 *
 * Установка (один раз):
 * 1. Открыть таблицу → Расширения → Apps Script.
 * 2. Создать файл, вставить этот код целиком, сохранить.
 * 3. Слева «Сервисы» (+) → добавить «YouTube Data API v3» и «YouTube Analytics API».
 * 4. Выбрать функцию installYouTubeTrigger → Выполнить → войти в Google-аккаунт,
 *    к которому привязан YouTube-канал школы, и разрешить доступ.
 * 5. Чтобы заполнить таблицу сразу, выполнить функцию updateYouTube.
 */

/* Запуск вручную или по расписанию */
function installYouTubeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'updateYouTube') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('updateYouTube').timeBased().everyDays(1).atHour(6).create();
  YT_.run();
}

function updateYouTube() { YT_.run(); }

/* Всё остальное спрятано внутри YT_, чтобы не пересекаться с другими скриптами таблицы */
var YT_ = (function () {
  var WEEKLY_SHEET = 'Еженедельная статистика';
  var MONTHLY_SHEET = 'Ежемесячная статистика';
  var SECTION = 'YOUTUBE';
  var MONTHS = ['январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь'];


  function run() {
    var ss = SpreadsheetApp.getActive();
    var subscribersNow = Number(YouTube.Channels.list('statistics', { mine: true }).items[0].statistics.subscriberCount);
    fillSheet_(ss.getSheetByName(WEEKLY_SHEET), weekPeriod_, subscribersNow);
    fillSheet_(ss.getSheetByName(MONTHLY_SHEET), monthPeriod_, subscribersNow);
  }

  /* Заполняет все прошедшие и текущий период листа. */
  function fillSheet_(sheet, parsePeriod, subscribersNow) {
    if (!sheet) return;
    var data = sheet.getDataRange().getDisplayValues();
    var rows = sectionRows_(data);
    if (!rows) throw new Error('На листе «' + sheet.getName() + '» нет блока ' + SECTION);

    var today = startOfDay_(new Date());
    var prevStart = null, latestCol = -1;
    for (var col = 1; col < data[0].length; col++) {
      var p = parsePeriod(data[0][col], prevStart);
      if (!p) continue;
      prevStart = p.start;
      if (p.start > today) continue;                       // будущие периоды пропускаем
      var end = p.end > today ? today : p.end;             // текущий период — по сегодня
      var s = analytics_(p.start, end);
      put_(sheet, rows['подписки'], col, s.subscribersGained);
      put_(sheet, rows['отписки'], col, s.subscribersLost);
      put_(sheet, rows['просмотры'], col, s.views);
      put_(sheet, rows['время просмотра, ч'], col, Math.round(s.estimatedMinutesWatched / 60));
      put_(sheet, rows['вовлеченность'], col, s.likes + s.comments + s.shares);
      put_(sheet, rows['количество видео'], col, videosPublished_(p.start, end));
      latestCol = col;
    }
    // Число подписчиков YouTube отдаёт только на сегодня — пишем его в последний период
    if (latestCol >= 0) put_(sheet, rows['подписчики'], latestCol, subscribersNow);
  }

  /* Строки блока YOUTUBE: { 'подписчики': номер строки, ... } */
  function sectionRows_(data) {
    var start = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === SECTION) { start = i; break; }
    }
    if (start < 0) return null;
    var rows = {};
    for (var j = start + 1; j < data.length; j++) {
      var name = String(data[j][0]).trim();
      if (!name) continue;
      if (name === name.toUpperCase() && /[A-ZА-Я]/.test(name)) break;  // следующий раздел
      rows[name.toLowerCase()] = j + 1;
    }
    return rows;
  }

  function put_(sheet, row, col, value) {
    if (!row || value === null || value === undefined || isNaN(value)) return;
    sheet.getRange(row, col + 1).setValue(value);
  }

  function analytics_(start, end) {
    var metrics = ['views', 'estimatedMinutesWatched', 'likes', 'comments', 'shares', 'subscribersGained', 'subscribersLost'];
    var res = YouTubeAnalytics.Reports.query({
      ids: 'channel==MINE',
      startDate: ymd_(start),
      endDate: ymd_(end),
      metrics: metrics.join(',')
    });
    var out = {};
    var row = (res.rows && res.rows[0]) || [];
    metrics.forEach(function (m, i) { out[m] = Number(row[i] || 0); });
    return out;
  }

  function videosPublished_(start, end) {
    var n = 0, token;
    var after = new Date(start), before = new Date(end.getTime() + 86400000);
    do {
      var res = YouTube.Search.list('id', {
        forMine: true, type: 'video', maxResults: 50, pageToken: token,
        publishedAfter: after.toISOString(), publishedBefore: before.toISOString()
      });
      n += (res.items || []).length;
      token = res.nextPageToken;
    } while (token);
    return n;
  }

  /* «5.10-11.10» → {start, end}; end — последний день недели. */
  function weekPeriod_(label, prevStart) {
    var m = String(label).match(/(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\s*[-–—]\s*(\d{1,2})\.(\d{1,2})/);
    if (!m) return null;
    var year = m[3] ? (m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]))
                    : (prevStart ? prevStart.getFullYear() : new Date().getFullYear());
    var start = new Date(year, Number(m[2]) - 1, Number(m[1]));
    if (prevStart && start < prevStart) start.setFullYear(year + 1);   // переход через Новый год
    var end = new Date(start.getFullYear(), Number(m[5]) - 1, Number(m[4]));
    if (end < start) end.setFullYear(end.getFullYear() + 1);
    return { start: start, end: end };
  }

  /* «Ноябрь 26» → {start, end} */
  function monthPeriod_(label) {
    var m = String(label).toLowerCase().match(/([а-яё]+)\s*(\d{2,4})/);
    if (!m) return null;
    var idx = MONTHS.indexOf(m[1]);
    if (idx < 0) return null;
    var year = m[2].length === 2 ? 2000 + Number(m[2]) : Number(m[2]);
    return { start: new Date(year, idx, 1), end: new Date(year, idx + 1, 0) };
  }

  function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  function ymd_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

  return { run: run };
})();
