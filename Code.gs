/**
 * 学習リンク集 ― バックエンド (Google Apps Script)
 *
 * スプレッドシートに紐づけて使います（拡張機能 → Apps Script）。
 * 手順は README.md を参照してください。
 *
 * 合い言葉・管理者パスワードはこのファイルには書きません。
 * 「プロジェクトの設定 → スクリプト プロパティ」に次の2つを登録します。
 *   EDIT_PASSWORD  … 職員共通の合い言葉（追加・編集に必要）
 *   ADMIN_PASSWORD … 管理者パスワード（削除に必要）
 */

var SHEET_LINKS    = 'links';
var SHEET_SUBJECTS = 'subjects';
var SHEET_LOG      = 'log';

var DEFAULT_SUBJECTS = [
  '国語', '社会', '算数', '理科', '生活', '音楽', '図画工作', '家庭',
  '体育', '道徳', '外国語', '総合', '学級活動・行事', '特別支援',
  '保健・食育', 'その他'
];

var LINK_HEADERS = [
  'id', 'サイト名', 'URL', '教科', '学年', 'ひとことメモ',
  '追加した人', '追加日時', '更新日時', '表示'
];

var MAX_URL  = 2000;
var MAX_NAME = 200;
var MAX_MEMO = 500;
var MAX_BODY = 8000;

/* =========================================================
   セットアップ（最初に1回だけ実行します）
   ========================================================= */
function セットアップ() {
  var ss = SpreadsheetApp.getActive();

  var links = ss.getSheetByName(SHEET_LINKS);
  if (!links) {
    links = ss.insertSheet(SHEET_LINKS);
    links.getRange(1, 1, 1, LINK_HEADERS.length).setValues([LINK_HEADERS]).setFontWeight('bold');
    links.setFrozenRows(1);
    links.setColumnWidth(2, 260);
    links.setColumnWidth(3, 360);
    links.setColumnWidth(6, 300);
  }

  var subjects = ss.getSheetByName(SHEET_SUBJECTS);
  if (!subjects) {
    subjects = ss.insertSheet(SHEET_SUBJECTS);
    subjects.getRange(1, 1).setValue('教科').setFontWeight('bold');
    subjects.setFrozenRows(1);
    subjects.getRange(2, 1, DEFAULT_SUBJECTS.length, 1)
      .setValues(DEFAULT_SUBJECTS.map(function (s) { return [s]; }));
  }

  var log = ss.getSheetByName(SHEET_LOG);
  if (!log) {
    log = ss.insertSheet(SHEET_LOG);
    log.getRange(1, 1, 1, 5)
      .setValues([['日時', '操作', 'id', 'サイト名', '実行した人']])
      .setFontWeight('bold');
    log.setFrozenRows(1);
  }

  var p = PropertiesService.getScriptProperties();
  var missing = [];
  if (!p.getProperty('EDIT_PASSWORD'))  missing.push('EDIT_PASSWORD');
  if (!p.getProperty('ADMIN_PASSWORD')) missing.push('ADMIN_PASSWORD');

  var msg = missing.length
    ? 'シートを用意しました。\n\nつぎに「プロジェクトの設定 → スクリプト プロパティ」で\n'
      + missing.join(' と ') + ' を登録してください。'
    : 'シートを用意しました。パスワードも登録済みです。デプロイに進めます。';
  SpreadsheetApp.getUi().alert(msg);
}

/* =========================================================
   読み取り
   ========================================================= */
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'list';
    if (action !== 'list') return json_({ ok: false, error: '不明な操作です。' });
    return json_({ ok: true, subjects: readSubjects_(), links: readLinks_() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function readSubjects_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SUBJECTS);
  if (!sh || sh.getLastRow() < 2) return DEFAULT_SUBJECTS.slice();
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (s) { return s; });
}

function readLinks_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LINKS);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, LINK_HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    if (r[9] === false) continue;              // 表示 = FALSE は隠す
    out.push({
      id:      String(r[0]),
      name:    String(r[1]),
      url:     String(r[2]),
      subject: String(r[3]),
      grades:  parseGrades_(r[4]),
      memo:    String(r[5] || ''),
      addedBy: String(r[6] || '')
    });
  }
  return out;
}

function parseGrades_(v) {
  return String(v || '').split(',')
    .map(function (s) { return parseInt(s, 10); })
    .filter(function (n) { return n >= 1 && n <= 6; });
}

/* =========================================================
   書き込み
   ========================================================= */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: '中身がありません。' });
    }
    if (e.postData.contents.length > MAX_BODY) {
      return json_({ ok: false, error: '送られてきた内容が大きすぎます。' });
    }
    var req = JSON.parse(e.postData.contents);

    if (req.action === 'delete') {
      if (!checkPassword_('ADMIN_PASSWORD', req.adminPw)) {
        return json_({ ok: false, error: '管理者パスワードが違います。' });
      }
      return deleteLink_(req);
    }

    if (!checkPassword_('EDIT_PASSWORD', req.pw)) {
      return json_({ ok: false, error: '合い言葉が違います。' });
    }

    if (req.action === 'add')    return addLink_(req);
    if (req.action === 'update') return updateLink_(req);
    if (req.action === 'title')  return fetchTitle_(req);
    return json_({ ok: false, error: '不明な操作です。' });

  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/**
 * パスワード照合。総当たりへの簡易的な歯止めとして、
 * 失敗が続いたら一定時間すべての書き込みを止めます。
 */
function checkPassword_(propKey, given) {
  var cache = CacheService.getScriptCache();
  if (cache.get('lockout')) return false;

  var expected = PropertiesService.getScriptProperties().getProperty(propKey);
  if (!expected) throw new Error('サーバー側に ' + propKey + ' が設定されていません。');

  if (String(given || '') === expected) {
    cache.remove('failCount');
    return true;
  }
  var n = parseInt(cache.get('failCount') || '0', 10) + 1;
  cache.put('failCount', String(n), 3600);
  if (n >= 10) {
    cache.put('lockout', '1', 900);            // 15分止める
    cache.remove('failCount');
  }
  return false;
}

function validate_(req, subjects) {
  var url  = String(req.url || '').trim();
  var name = String(req.name || '').trim();
  var subject = String(req.subject || '').trim();
  var grades = (req.grades || [])
    .map(function (g) { return parseInt(g, 10); })
    .filter(function (g) { return g >= 1 && g <= 6; });
  grades = grades.filter(function (g, i) { return grades.indexOf(g) === i; }).sort();

  if (!/^https?:\/\//i.test(url)) throw new Error('URL は http:// か https:// で始めてください。');
  if (url.length > MAX_URL)   throw new Error('URL が長すぎます。');
  if (!name)                  throw new Error('サイト名を入れてください。');
  if (name.length > MAX_NAME) throw new Error('サイト名が長すぎます。');
  if (subjects.indexOf(subject) < 0) throw new Error('教科が正しくありません。');
  if (!grades.length)         throw new Error('学年を1つ以上えらんでください。');

  return {
    url: url, name: name, subject: subject, grades: grades,
    memo:    String(req.memo || '').trim().slice(0, MAX_MEMO),
    addedBy: String(req.addedBy || '').trim().slice(0, 100)
  };
}

function addLink_(req) {
  var v = validate_(req, readSubjects_());
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LINKS);
    var id = Utilities.getUuid().slice(0, 8);
    var now = new Date();
    sh.appendRow([id, v.name, v.url, v.subject, v.grades.join(','),
                  v.memo, v.addedBy, now, now, true]);
    writeLog_('追加', id, v.name, v.addedBy);
    return json_({ ok: true, id: id });
  } finally {
    lock.releaseLock();
  }
}

function updateLink_(req) {
  var v = validate_(req, readSubjects_());
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LINKS);
    var row = findRow_(sh, req.id);
    if (row < 0) return json_({ ok: false, error: 'そのリンクは見つかりませんでした。' });
    sh.getRange(row, 2, 1, 6).setValues([[
      v.name, v.url, v.subject, v.grades.join(','), v.memo, v.addedBy
    ]]);
    sh.getRange(row, 9).setValue(new Date());
    writeLog_('更新', req.id, v.name, v.addedBy);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/** 行そのものは消さず「表示」を FALSE にします（誤操作から戻せるように）。 */
function deleteLink_(req) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LINKS);
    var row = findRow_(sh, req.id);
    if (row < 0) return json_({ ok: false, error: 'そのリンクは見つかりませんでした。' });
    sh.getRange(row, 10).setValue(false);
    sh.getRange(row, 9).setValue(new Date());
    writeLog_('非表示', req.id, sh.getRange(row, 2).getValue(), '管理者');
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function findRow_(sh, id) {
  if (!id || sh.getLastRow() < 2) return -1;
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function writeLog_(action, id, name, by) {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
    if (sh) sh.appendRow([new Date(), action, id, name, by || '']);
  } catch (e) { /* ログの失敗で本体は止めない */ }
}

/* =========================================================
   URL からサイト名を取り込む
   ========================================================= */
function fetchTitle_(req) {
  var url = String(req.url || '').trim();
  if (!/^https?:\/\//i.test(url)) return json_({ ok: false, error: 'URL が正しくありません。' });

  var res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; SchoolLinkQR/1.0)' }
  });
  if (res.getResponseCode() >= 400) {
    return json_({ ok: false, error: 'ページを開けませんでした（' + res.getResponseCode() + '）。' });
  }

  var blob = res.getBlob();
  var html = blob.getDataAsString('UTF-8');
  var cs = html.match(/charset\s*=\s*["']?\s*([\w\-]+)/i);
  if (cs) {
    var name = cs[1].toUpperCase();
    if (name !== 'UTF-8' && name !== 'UTF8') {
      try { html = blob.getDataAsString(name); } catch (e) { /* そのまま使う */ }
    }
  }

  var m = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  if (!m) return json_({ ok: false, error: 'このページからは名前を取得できませんでした。' });

  var title = m[1]
    .replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>').replace(/&quot;/gi, '"').replace(/&#0*39;/g, "'")
    .replace(/\s+/g, ' ').trim().slice(0, 120);

  if (!title) return json_({ ok: false, error: 'このページからは名前を取得できませんでした。' });
  return json_({ ok: true, title: title });
}

/* =========================================================
   共通
   ========================================================= */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
