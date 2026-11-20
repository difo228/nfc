// Petroretail NFC — приём отзывов со страницы гостя в Google Таблицу.
// Установка: Таблица → Расширения → Apps Script → вставить этот файл целиком →
// Развернуть → Новое развертывание → Веб-приложение (запуск от: «Я», доступ: «Все») →
// ссылку …/exec вставить в CONFIG.endpoint в index.html.

const SHEET = 'Отзывы';
const HEAD = ['Время', 'АЗС', 'Зона', 'Оценка', 'Причины', 'Комментарий', 'Телефон', 'Статус',
              'Язык', 'Источник', 'Устройство', 'ID', 'zone', 'chips'];
const COL = { phone: 7, status: 8, id: 12, chips: 14 };   // номера столбцов (с 1)
const ZONE = { kassa: 'Касса', fuel: 'Заправщик', exit: 'Выход', wc: 'Туалет' };
const STATUS_IN = { 'Новая': 'new', 'В работе': 'new', 'Решено': 'done' };

// Создаёт лист с заголовком при первом отзыве
function sheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
    sh.getRange(2, COL.status, 999, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['Новая', 'В работе', 'Решено'], true)
        .setAllowInvalid(true).build());
    sh.hideColumns(13, 2);   // служебные коды для страницы результатов
  }
  return sh;
}

// Текст от гостя: обрезаем и не даём начинаться с = + - @ (иначе таблица примет за формулу)
function txt_(v, max) {
  v = String(v == null ? '' : v).slice(0, max || 600);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function rowOf_(sh, id) {
  const n = sh.getLastRow();
  if (n < 2) return -1;
  const ids = sh.getRange(2, COL.id, n - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

// Страница присылает оценку сразу после касания смайла, затем уточнения с тем же id
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const id = String(d.id || '').slice(0, 40);
    if (!id || !d.station || !ZONE[d.zone]) return out_({ ok: false });
    const sh = sheet_();
    const r = rowOf_(sh, id);
    const chips = (Array.isArray(d.chips) ? d.chips : []).map(String).join(',').slice(0, 300);
    const phone = d.phone ? "'" + String(d.phone).replace(/[^\d+()\-\s]/g, '').slice(0, 30) : '';
    if (r < 0) {
      const rating = Math.max(1, Math.min(5, Math.round(Number(d.rating)) || 1));
      sh.appendRow([new Date(Number(d.t) || Date.now()), txt_(d.station, 20), ZONE[d.zone], rating,
        txt_(d.chipsText, 300), txt_(d.comment), phone, rating <= 2 ? 'Новая' : '',
        txt_(d.lang, 5), txt_(d.src, 10), txt_(d.device, 10), id, d.zone, chips]);
    } else {
      // уточнение: меняем только причины, комментарий и телефон; время, оценку и статус не трогаем
      sh.getRange(r, 5, 1, 2).setValues([[txt_(d.chipsText, 300), txt_(d.comment)]]);
      sh.getRange(r, COL.chips).setValue(chips);
      if (phone) sh.getRange(r, COL.phone).setValue(phone);
    }
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Страница результатов (#results) читает последние 500 отзывов. Телефоны сюда НЕ отдаются.
function doGet() {
  const sh = sheet_();
  const n = sh.getLastRow();
  if (n < 2) return out_({ reviews: [] });
  const cnt = Math.min(500, n - 1);
  const rows = sh.getRange(n - cnt + 1, 1, cnt, HEAD.length).getValues().reverse();
  return out_({ reviews: rows.filter(x => x[COL.id - 1]).map(x => ({
    id: String(x[11]), t: new Date(x[0]).getTime(), station: String(x[1]), zone: String(x[12]),
    rating: Number(x[3]), chips: String(x[13]).split(',').filter(String), comment: String(x[5]),
    status: STATUS_IN[x[7]] || (Number(x[3]) <= 2 ? 'new' : 'ok'), hasContact: !!x[6],
    lang: String(x[8]), src: String(x[9])
  })) });
}
