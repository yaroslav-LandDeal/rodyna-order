/**
 * Бекенд для додатка «Замовлення з Корчми Родина».
 * Прив'язаний до Google Таблиці: Розширення → Apps Script.
 *
 * POST  {name, tg_id, items:[{code, name, qty, price}]}  → зберігає замовлення
 *       (попереднє замовлення тієї ж людини замінюється)
 * GET   ?summary=1                                       → підсумок у JSON
 * GET   ?my=<telegram id>                                → збережене замовлення людини
 * GET   без параметрів                                   → перевірка, що працює
 * POST  від Telegram (webhook)                           → відповіді бота: /start, /order
 *
 * Токен бота: Налаштування проєкту → Властивості скрипту → BOT_TOKEN.
 * Webhook: відкрити в браузері
 *   https://api.telegram.org/bot<TOKEN>/setWebhook?url=<URL веб-додатка /exec>
 */

var SHEET_ORDERS = 'Замовлення';
var SHEET_SUMMARY = 'Підсумок';
var TZ = 'Europe/Madrid';

function doGet(e) {
  setup_();
  if (e && e.parameter && e.parameter.summary) return json_(summary_());
  if (e && e.parameter && e.parameter.my) return json_({ ok: true, order: myOrder_('tg:' + String(e.parameter.my)) });
  return json_({ ok: true, service: 'rodyna-order', time: new Date().toISOString() });
}

function doPost(e) {
  setup_();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    if (body.update_id) { handleUpdate_(body); return json_({ ok: true }); }
    var name = String(body.name || '').trim().slice(0, 80) || 'Без імені';
    var tgId = String(body.tg_id || '').trim();
    var items = Array.isArray(body.items) ? body.items : [];
    var key = tgId ? 'tg:' + tgId : 'name:' + name.toLowerCase();
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);

    if (body.delete) {
      var removed = removeByKey_(sh, key);
      if (tgId && removed) sendMessage_(tgId, '🗑 Ваше замовлення видалено. Щоб зробити нове, натисніть кнопку «Замовити» внизу.');
      return json_({ ok: true, removed: removed });
    }
    if (!items.length) return json_({ ok: false, error: 'empty' });

    removeByKey_(sh, key); // попереднє замовлення цієї людини замінюється

    var now = Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm');
    var rows = [];
    items.forEach(function (it, i) {
      var qty = Math.max(0, Math.min(20, parseInt(it.qty, 10) || 0));
      var price = Math.round((parseFloat(it.price) || 0) * 100) / 100;
      if (!qty) return;
      rows.push([
        now, tgId, name,
        String(it.code || '').slice(0, 5),
        String(it.name || '').slice(0, 120),
        qty, price, Math.round(qty * price * 100) / 100,
        parseInt(it.order, 10) || i,
        key
      ]);
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    var notified = false;
    if (tgId && rows.length) notified = sendMessage_(tgId, orderText_(name, rows));
    return json_({ ok: true, saved: rows.length, name: name, notified: notified });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function summary_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  var last = sh.getLastRow();
  var people = {}, dishes = {}, grand = 0;
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 10).getValues().forEach(function (r) {
      var name = r[2], code = r[3], dish = r[4], qty = +r[5], sum = +r[7], ord = +r[8], key = r[9];
      if (!qty) return;
      people[key] = people[key] || { name: name, sum: 0, time: r[0] };
      people[key].sum = Math.round((people[key].sum + sum) * 100) / 100;
      dishes[code] = dishes[code] || { code: code, name: dish, qty: 0, sum: 0, order: ord };
      dishes[code].qty += qty;
      dishes[code].sum = Math.round((dishes[code].sum + sum) * 100) / 100;
      grand = Math.round((grand + sum) * 100) / 100;
    });
  }
  var dishList = Object.keys(dishes).map(function (k) { return dishes[k]; })
    .sort(function (a, b) { return a.order - b.order; });
  var peopleList = Object.keys(people).map(function (k) { return people[k]; })
    .sort(function (a, b) { return a.name.localeCompare(b.name, 'uk'); });
  return { ok: true, people: peopleList, dishes: dishList, total: grand, updated: new Date().toISOString() };
}

/** Видаляє всі рядки з цим ключем; повертає кількість видалених. */
function removeByKey_(sh, key) {
  var last = sh.getLastRow(), n = 0;
  if (last < 2) return 0;
  var keys = sh.getRange(2, 10, last - 1, 1).getValues();
  for (var r = keys.length - 1; r >= 0; r--) {
    if (keys[r][0] === key) { sh.deleteRow(r + 2); n++; }
  }
  return n;
}

/** Замовлення однієї людини за ключем (масив рядків аркуша). */
function myOrder_(key) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 10).getValues()
    .filter(function (r) { return r[9] === key && +r[5] > 0; })
    .sort(function (a, b) { return +a[8] - +b[8]; })
    .map(function (r) { return { code: r[3], name: r[4], qty: +r[5], price: +r[6], sum: +r[7], time: r[0], person: r[2] }; });
}

function eur_(n) { return (Math.round(n * 100) / 100).toFixed(2).replace('.', ',') + ' €'; }

/** Текст замовлення для повідомлення в Telegram (rows = рядки аркуша або об'єкти myOrder_). */
function orderText_(name, rows) {
  var total = 0;
  var lines = rows.map(function (r) {
    var code = r.code || r[3], dish = r.name || r[4], qty = r.qty || r[5], sum = r.sum != null ? r.sum : r[7];
    total += +sum;
    return code + ' ' + dish + ' ×' + qty + ' — ' + eur_(sum);
  });
  return '🧾 Ваше замовлення, ' + name + ':\n' + lines.join('\n') + '\nРазом: ' + eur_(total) +
    '\n\nЗмінити можна в додатку: кнопка «Замовити» внизу. Команда /order покаже поточне замовлення.';
}

/** Повідомлення від Telegram (webhook). */
function handleUpdate_(u) {
  var m = u.message || u.edited_message;
  if (!m || !m.chat || m.chat.type !== 'private') return;
  var chatId = String(m.chat.id);
  var text = String(m.text || '').trim();
  var from = m.from || {};
  var name = [from.first_name, from.last_name].filter(Boolean).join(' ') || 'Без імені';
  if (/^\/start/.test(text)) {
    sendMessage_(chatId, 'Привіт! Щоб зробити замовлення з «Корчми Родина», натисніть кнопку «Замовити» внизу, оберіть страви й надішліть. Після цього я пришлю сюди ваше замовлення по стравах. Команда /order покаже поточне замовлення.');
    return;
  }
  var order = myOrder_('tg:' + chatId);
  if (!order.length) { sendMessage_(chatId, 'Замовлення ще немає. Натисніть кнопку «Замовити» внизу, щоб обрати страви.'); return; }
  sendMessage_(chatId, orderText_(order[0].person || name, order));
}

function sendMessage_(chatId, text) {
  var token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  if (!token) return false;
  try {
    var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: chatId, text: text, disable_web_page_preview: true })
    });
    return res.getResponseCode() === 200;
  } catch (err) { return false; }
}

/** Створює аркуші й формули, якщо їх ще немає. */
function setup_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_ORDERS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ORDERS, 0);
    sh.getRange(1, 1, 1, 10).setValues([[
      'Час', 'Telegram ID', 'Ім\'я', 'Код', 'Страва', 'К-сть', 'Ціна, €', 'Сума, €', 'Порядок', 'Ключ'
    ]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(5, 320);
    sh.hideColumns(9, 2);
  }
  var sum = ss.getSheetByName(SHEET_SUMMARY);
  if (!sum) {
    sum = ss.insertSheet(SHEET_SUMMARY, 1);
    sum.getRange('A1').setValue('Разом, €').setFontWeight('bold');
    sum.getRange('B1').setFormula('=SUM(\'' + SHEET_ORDERS + '\'!H2:H)');
    sum.getRange('A2').setValue('Учасників').setFontWeight('bold');
    sum.getRange('B2').setFormula('=COUNTUNIQUE(\'' + SHEET_ORDERS + '\'!J2:J)');
    sum.getRange('A4').setValue('Замовити в ресторані').setFontWeight('bold');
    sum.getRange('A5').setFormula(
      '=IFERROR(QUERY(\'' + SHEET_ORDERS + '\'!D2:I, ' +
      '"select D, E, sum(F), sum(H) where D is not null group by D, E, I order by I ' +
      'label D \'Код\', E \'Страва\', sum(F) \'К-сть\', sum(H) \'Сума, €\'", 0), "Замовлень ще немає")');
    sum.getRange('G4').setValue('Хто скільки платить').setFontWeight('bold');
    sum.getRange('G5').setFormula(
      '=IFERROR(QUERY(\'' + SHEET_ORDERS + '\'!C2:H, ' +
      '"select C, sum(H) where C is not null group by C order by C ' +
      'label C \'Учасник\', sum(H) \'Сума, €\'", 0), "")');
    sum.setColumnWidth(2, 320);
    var def = ss.getSheets()[0];
    if (def.getName() !== SHEET_ORDERS && def.getName() !== SHEET_SUMMARY && ss.getSheets().length > 2
        && def.getLastRow() === 0) ss.deleteSheet(def);
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
