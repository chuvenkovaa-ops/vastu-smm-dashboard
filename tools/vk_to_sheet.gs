/**
 * VK → «Таблица для дашборда VASTU»
 *
 * Раз в день забирает статистику сообщества и пишет её в блок ВКОНТАКТЕ
 * на листах «Еженедельная статистика» и «Ежемесячная статистика».
 *
 * Заполняет строки: Подписчики, Прирост подписчиков, Охват, Посещения профиля,
 * Количество публикаций. «Просмотры» (клипы и записи), «Лиды» и GetCourse-строки не трогает.
 *
 * Прошлые периоды заполняются, только если ячейка пустая: то, что вы вносили руками,
 * скрипт не перезаписывает. Текущий период (неделя и месяц, которые ещё идут)
 * и период, закончившийся за последние 3 дня, обновляются каждый день.
 *
 * Установка (один раз):
 * 1. Открыть таблицу → Расширения → Apps Script → новый файл → вставить этот код → сохранить.
 * 2. Слева «Настройки проекта» (шестерёнка) → «Свойства скрипта» → добавить два свойства:
 *      VK_TOKEN — ключ доступа (как получить — в README репозитория, раздел про VK)
 *      VK_GROUP — короткое имя сообщества из адреса, например vastuhome
 * 3. Выбрать функцию installVKTrigger → Выполнить → разрешить доступ.
 *    Если VK ответит, что у ключа нет доступа к статистике, скрипт напишет это в журнале выполнения.
 */

/* Запуск вручную или по расписанию */
function installVKTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'updateVK') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('updateVK').timeBased().everyDays(1).atHour(6).create();
  VK_.run();
}

function updateVK() { VK_.run(); }

/* Всё остальное спрятано внутри VK_, чтобы не пересекаться с другими скриптами таблицы */
var VK_ = (function () {
  var VK_VERSION = '5.199';
  var WEEKLY_SHEET = 'Еженедельная статистика';
  var MONTHLY_SHEET = 'Ежемесячная статистика';
  var SECTION = 'ВКОНТАКТЕ';
  var MONTHS = ['январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь'];


  function run() {
    var props = PropertiesService.getScriptProperties();
    var token = String(props.getProperty('VK_TOKEN') || '').replace(/\s+/g, '').replace(/^["']|["']$/g, '');
    var groupName = String(props.getProperty('VK_GROUP') || '').trim()
      .replace(/^https?:\/\/(m\.)?vk\.(com|ru)\//, '').replace(/^@/, '').replace(/\/.*$/, '');
    if (!token || !groupName) throw new Error('Добавьте свойства скрипта VK_TOKEN и VK_GROUP (Настройки проекта → Свойства скрипта).');
    if (token.length < 50) throw new Error('VK_TOKEN слишком короткий (' + token.length + ' символов): похоже, скопирован не весь ключ. Ключ VK — это длинная строка из 80+ символов.');

    var group = groupInfo_(token, groupName);
    var ss = SpreadsheetApp.getActive();
    var weekly = ss.getSheetByName(WEEKLY_SHEET), monthly = ss.getSheetByName(MONTHLY_SHEET);

    // Публикации за весь охват обеих таблиц получаем одним проходом по стене
    var earliest = earliestStart_([[weekly, weekPeriod_], [monthly, monthPeriod_]]);
    var postDates = earliest ? wallDates_(token, group.id, earliest) : [];

    fillSheet_(weekly, weekPeriod_, 'week', token, group, postDates);
    fillSheet_(monthly, monthPeriod_, 'month', token, group, postDates);
  }

  function fillSheet_(sheet, parsePeriod, interval, token, group, postDates) {
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
      if (p.start > today) continue;
      var current = p.end >= today;
      var fresh = p.end >= new Date(today.getTime() - 3 * 86400000);   // закончился недавно — досчитываем последние дни
      var end = current ? today : p.end;
      var cell = function (name) { var r = rows[name]; return r ? data[r - 1][col] : 'x'; };
      var need = function (name) { return fresh || String(cell(name)).trim() === ''; };

      if (need('прирост подписчиков') || need('охват') || need('посещения профиля')) {
        var s = stats_(token, group.id, p.start, end, interval);
        if (s) {
          if (need('прирост подписчиков')) put_(sheet, rows['прирост подписчиков'], col, s.subscribed - s.unsubscribed);
          if (need('охват')) put_(sheet, rows['охват'], col, s.reach);
          if (need('посещения профиля')) put_(sheet, rows['посещения профиля'], col, s.visitors);
        }
      }
      if (need('количество публикаций')) {
        var from = p.start.getTime() / 1000, to = (end.getTime() + 86400000) / 1000;
        put_(sheet, rows['количество публикаций'], col, postDates.filter(function (d) { return d >= from && d < to; }).length);
      }
      latestCol = col;
    }
    // Число подписчиков VK отдаёт только на сегодня — пишем его в текущий период
    if (latestCol >= 0) put_(sheet, rows['подписчики'], latestCol, group.members);
  }

  /* ---------- VK API ---------- */

  function vk_(method, params, token) {
    params.access_token = token;
    params.v = VK_VERSION;
    var q = Object.keys(params).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
    Utilities.sleep(350);   // VK разрешает не больше 3 запросов в секунду
    var res = JSON.parse(UrlFetchApp.fetch('https://api.vk.com/method/' + method + '?' + q, { muteHttpExceptions: true }).getContentText());
    if (res.error) {
      var e = res.error;
      var hint = e.error_code === 5
        ? ' VK не принял ключ. Создайте ключ заново (Управление → Работа с API → Ключи доступа), скопируйте его целиком кнопкой «Показать» и вставьте в VK_TOKEN без пробелов.'
        : (e.error_code === 15 || e.error_code === 27 || e.error_code === 7)
        ? ' У ключа нет доступа к этому методу. Нужен ключ администратора сообщества с доступом к статистике (см. README).'
        : '';
      throw new Error('VK ' + method + ': ' + e.error_code + ' ' + e.error_msg + '.' + hint);
    }
    return res.response;
  }

  function groupInfo_(token, groupName) {
    var r = vk_('groups.getById', { group_id: groupName, fields: 'members_count' }, token);
    var g = (r.groups || r)[0];
    return { id: g.id, members: g.members_count };
  }

  /* Сумма подписок/отписок, уникальный охват и посетители за период. */
  function stats_(token, groupId, start, end, interval) {
    var from = Math.floor(start.getTime() / 1000), to = Math.floor((end.getTime() + 86400000) / 1000) - 1;
    var periods = vk_('stats.get', { group_id: groupId, timestamp_from: from, timestamp_to: to, interval: interval }, token);
    if (!periods || !periods.length) return null;
    // Берём период, который начинается с нужной даты; если VK разбил иначе — суммируем
    var key = ymd_(start);
    var match = periods.filter(function (p) { return dateKey_(p.period_from) === key; });
    var list = match.length ? match : periods;
    var out = { subscribed: 0, unsubscribed: 0, reach: 0, visitors: 0 };
    list.forEach(function (p) {
      var a = p.activity || {}, r = p.reach || {}, v = p.visitors || {};
      out.subscribed += Number(a.subscribed || 0);
      out.unsubscribed += Number(a.unsubscribed || 0);
      out.reach += Number(r.reach || 0);
      out.visitors += Number(v.visitors || 0);
    });
    return out;
  }

  /* Даты (unix) всех записей на стене начиная с from. */
  function wallDates_(token, groupId, from) {
    var limit = from.getTime() / 1000, dates = [], offset = 0;
    while (true) {
      var r = vk_('wall.get', { owner_id: -groupId, count: 100, offset: offset, filter: 'owner' }, token);
      var items = r.items || [];
      if (!items.length) break;
      var older = false;
      items.forEach(function (it) {
        if (it.date >= limit) dates.push(it.date);
        else if (!it.is_pinned) older = true;      // закреп может быть старым — его не считаем концом
      });
      offset += items.length;
      if (older || offset >= r.count) break;
    }
    return dates;
  }

  /* ---------- Таблица ---------- */

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
      if (name === name.toUpperCase() && /[A-ZА-Я]/.test(name)) break;
      rows[name.toLowerCase()] = j + 1;
    }
    return rows;
  }

  function earliestStart_(pairs) {
    var min = null, today = startOfDay_(new Date());
    pairs.forEach(function (pair) {
      var sheet = pair[0]; if (!sheet) return;
      var head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0], prev = null;
      for (var c = 1; c < head.length; c++) {
        var p = pair[1](head[c], prev); if (!p) continue; prev = p.start;
        if (p.start <= today && (!min || p.start < min)) min = p.start;
      }
    });
    return min;
  }

  function put_(sheet, row, col, value) {
    if (!row || value === null || value === undefined || isNaN(value)) return;
    sheet.getRange(row, col + 1).setValue(value);
  }

  function weekPeriod_(label, prevStart) {
    var m = String(label).match(/(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\s*[-–—]\s*(\d{1,2})\.(\d{1,2})/);
    if (!m) return null;
    var year = m[3] ? (m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]))
                    : (prevStart ? prevStart.getFullYear() : new Date().getFullYear());
    var start = new Date(year, Number(m[2]) - 1, Number(m[1]));
    if (prevStart && start < prevStart) start.setFullYear(year + 1);
    var end = new Date(start.getFullYear(), Number(m[5]) - 1, Number(m[4]));
    if (end < start) end.setFullYear(end.getFullYear() + 1);
    return { start: start, end: end };
  }

  function monthPeriod_(label) {
    var m = String(label).toLowerCase().match(/([а-яё]+)\s*(\d{2,4})/);
    if (!m) return null;
    var idx = MONTHS.indexOf(m[1]);
    if (idx < 0) return null;
    var year = m[2].length === 2 ? 2000 + Number(m[2]) : Number(m[2]);
    return { start: new Date(year, idx, 1), end: new Date(year, idx + 1, 0) };
  }

  function dateKey_(v) {
    if (typeof v === 'number' || /^\d+$/.test(String(v))) return ymd_(new Date(Number(v) * 1000));
    return String(v).slice(0, 10);
  }
  function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  function ymd_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

  return { run: run };
})();
