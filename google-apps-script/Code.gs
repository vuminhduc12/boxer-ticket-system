/**
 * 応援チケット管理 - Google Apps Script バックエンド
 *
 * セットアップ手順:
 * 1. Google スプレッドシートを新規作成
 * 2. 拡張機能 → Apps Script を開き、このコードを貼り付け
 * 3. デプロイ → 新しいデプロイ → 種類: ウェブアプリ
 *    - 実行ユーザー: 自分
 *    - アクセス: 全員
 * 4. 表示された URL をサイト管理画面の「API URL」に入力
 */

var SHEET_CONFIG = '設定';
var SHEET_APPS = '申込';

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  try {
    if (action === 'config') return json_(loadPublicConfig_());
    if (action === 'applications') {
      if (!checkPin_(e.parameter.pin)) return json_({ error: 'PINが正しくありません' }, 401);
      return json_(loadApplications_());
    }
    if (action === 'stats') {
      if (!checkPin_(e.parameter.pin)) return json_({ error: 'PINが正しくありません' }, 401);
      return json_(calcStats_());
    }
    return json_({ error: '不明な action' }, 400);
  } catch (err) {
    return json_({ error: String(err) }, 500);
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'application') return json_(saveApplication_(body.data));
    if (body.action === 'config') {
      if (!checkPin_(body.pin)) return json_({ error: 'PINが正しくありません' }, 401);
      saveConfig_(body.data);
      return json_({ ok: true });
    }
    if (body.action === 'updateStatus') {
      if (!checkPin_(body.pin)) return json_({ error: 'PINが正しくありません' }, 401);
      updateStatus_(body.id, body.status);
      return json_({ ok: true });
    }
    if (body.action === 'deleteApp') {
      if (!checkPin_(body.pin)) return json_({ error: 'PINが正しくありません' }, 401);
      deleteApp_(body.id);
      return json_({ ok: true });
    }
    if (body.action === 'verifyPin') {
      return json_({ ok: checkPin_(body.pin) });
    }
    return json_({ error: '不明な action' }, 400);
  } catch (err) {
    return json_({ error: String(err) }, 500);
  }
}

function getSs_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureSheets_() {
  var ss = getSs_();
  if (!ss.getSheetByName(SHEET_CONFIG)) {
    var s = ss.insertSheet(SHEET_CONFIG);
    s.getRange('A1:B1').setValues([['key', 'value']]);
  }
  if (!ss.getSheetByName(SHEET_APPS)) {
    var a = ss.insertSheet(SHEET_APPS);
    a.getRange(1, 1, 1, 12).setValues([[
      'id', 'date', 'name', 'kana', 'email', 'phone',
      'ticketType', 'quantity', 'payment', 'paymentStatus', 'message', 'total'
    ]]);
  }
}

function loadConfig_() {
  ensureSheets_();
  var sheet = getSs_().getSheetByName(SHEET_CONFIG);
  var data = sheet.getDataRange().getValues();
  var config = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) config[data[i][0]] = data[i][1];
  }
  config.spreadsheetUrl = getSs_().getUrl();
  return config;
}

/** 公開サイト用 — PIN など管理情報を除外 */
function loadPublicConfig_() {
  var config = loadConfig_();
  delete config.pin;
  return config;
}

function saveConfig_(config) {
  ensureSheets_();
  var sheet = getSs_().getSheetByName(SHEET_CONFIG);
  sheet.clear();
  sheet.getRange('A1:B1').setValues([['key', 'value']]);
  var rows = [];
  for (var key in config) {
    if (config.hasOwnProperty(key)) rows.push([key, String(config[key])]);
  }
  if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function loadApplications_() {
  ensureSheets_();
  var sheet = getSs_().getSheetByName(SHEET_APPS);
  var data = sheet.getDataRange().getValues();
  var apps = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    apps.push({
      id: data[i][0], date: data[i][1], name: data[i][2], kana: data[i][3],
      email: data[i][4], phone: data[i][5], ticketType: data[i][6],
      quantity: Number(data[i][7]), payment: data[i][8],
      paymentStatus: data[i][9], message: data[i][10], total: Number(data[i][11] || 0)
    });
  }
  return apps;
}

function saveApplication_(app) {
  ensureSheets_();
  var sheet = getSs_().getSheetByName(SHEET_APPS);
  var config = loadConfig_();
  var price = app.ticketType === '一般席' ? Number(config.priceGeneral || 3000)
    : app.ticketType === 'リングサイド' ? Number(config.priceRingside || 5000)
    : Number(config.priceGroup || 2500);
  var total = price * Number(app.quantity);
  var id = app.id || Utilities.getUuid();
  sheet.appendRow([
    id, app.date, app.name, app.kana, app.email, app.phone || '',
    app.ticketType, app.quantity, app.payment, app.paymentStatus || 'unpaid',
    app.message || '', total
  ]);
  try {
    if (config.notifyEmail || config.email) {
      var to = config.notifyEmail || config.email;
      MailApp.sendEmail(to, '【新規申込】' + app.name + ' さん',
        '新しいチケット申込がありました。\n\n' +
        '名前: ' + app.name + '\n' +
        '種別: ' + app.ticketType + '\n' +
        '枚数: ' + app.quantity + '\n' +
        '合計: ¥' + total + '\n' +
        'メール: ' + app.email + '\n' +
        '電話: ' + (app.phone || '-') + '\n' +
        '支払: ' + app.payment + '\n' +
        '備考: ' + (app.message || '-')
      );
    }
  } catch (e) {}
  return { ok: true, id: id };
}

function updateStatus_(id, status) {
  var sheet = getSs_().getSheetByName(SHEET_APPS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 10).setValue(status);
      return;
    }
  }
}

function deleteApp_(id) {
  var sheet = getSs_().getSheetByName(SHEET_APPS);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === id) sheet.deleteRow(i + 1);
  }
}

function checkPin_(pin) {
  var config = loadConfig_();
  return String(pin) === String(config.pin || '1234');
}

function calcStats_() {
  var apps = loadApplications_();
  var active = apps.filter(function(a) { return a.paymentStatus !== 'cancelled'; });
  return {
    totalApps: active.length,
    totalTickets: active.reduce(function(s, a) { return s + Number(a.quantity); }, 0),
    totalRevenue: active.reduce(function(s, a) { return s + Number(a.total || 0); }, 0),
    paid: apps.filter(function(a) { return a.paymentStatus === 'paid'; }).length,
    unpaid: apps.filter(function(a) { return a.paymentStatus === 'unpaid'; }).length,
    cancelled: apps.filter(function(a) { return a.paymentStatus === 'cancelled'; }).length
  };
}

function json_(obj, code) {
  code = code || 200;
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
