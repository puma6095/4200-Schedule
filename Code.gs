/**
 * 奇美醫院教學部 · 排班系統後端　v2
 * ---------------------------------------------------------------
 * v2 新增「局部更新（merge）」：網頁只送出異動的部分，由後端合併，
 *     存檔往返次數從 2 次降為 1 次，速度明顯變快，也不會覆蓋他人資料。
 *
 * 更新方式（已部署過 v1）：
 *   1. 開啟原本的 Apps Script 專案，全選舊程式碼刪除，貼上本檔內容，存檔。
 *   2. 部署 → 管理部署作業 → 鉛筆圖示 → 版本選「新版本」→ 部署。
 *      ※ 網址不會變，index.html 不用改。
 *   3. 回網頁 Ctrl+F5，存檔速度即提升。
 *
 * 首次部署：
 *   1. 開一份 Google 試算表 → 擴充功能 → Apps Script，貼上本檔內容。
 *   2. 部署 → 新增部署作業 → 網頁應用程式。
 *   3. 執行身分：我；具有存取權的使用者：任何人。
 *   4. 複製 /exec 結尾的網址，填入 index.html 的 REMOTE_URL。
 */

var SHEET_NAME = 'store';
var VERSION = 2;

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('SS_ID');
    if (!id) {
      var created = SpreadsheetApp.create('教學部排班系統資料');
      id = created.getId();
      props.setProperty('SS_ID', id);
    }
    ss = SpreadsheetApp.openById(id);
  }
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['key', 'value', 'updated']);
  }
  return sh;
}

function findRow_(sh, key) {
  var last = Math.max(sh.getLastRow(), 1);
  var keys = sh.getRange(1, 1, last, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) return i + 1;
  }
  return 0;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 深層合併：物件逐層併入，陣列與純值直接取代 */
function mergeInto_(base, patch) {
  for (var k in patch) {
    var v = patch[k];
    if (v && typeof v === 'object' && !(v instanceof Array)) {
      if (!base[k] || typeof base[k] !== 'object' || (base[k] instanceof Array)) base[k] = {};
      mergeInto_(base[k], v);
    } else {
      base[k] = v;
    }
  }
}

/** 讀取：GET ?key=xxx　／　版本查詢：GET ?op=ver */
function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (p.op === 'ver') return json_({ ver: VERSION });

    var out = { value: null };
    if (p.key) {
      var sh = getSheet_();
      var row = findRow_(sh, p.key);
      if (row) out.value = sh.getRange(row, 2).getValue();
    }
    return json_(out);
  } catch (err) {
    return json_({ error: String(err) });
  }
}

/**
 * 寫入：POST
 *   { key, value }  → 整筆覆蓋
 *   { key, merge }  → 局部合併（v2）
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);                      // 避免多人同時寫入互相覆蓋
    var body = JSON.parse(e.postData.contents);
    if (!body || !body.key) return json_({ ok: false, error: 'missing key' });

    var sh = getSheet_();
    var row = findRow_(sh, body.key);
    var now = new Date();
    var text;

    if (body.merge) {
      var cur = row ? sh.getRange(row, 2).getValue() : '';
      var base = {};
      if (cur) {
        try { base = JSON.parse(cur); } catch (x) { base = {}; }
      }
      mergeInto_(base, body.merge);
      text = JSON.stringify(base);
    } else if (typeof body.value === 'string') {
      text = body.value;
    } else {
      return json_({ ok: false, error: 'missing value' });   // 絕不寫入空值
    }

    if (row) {
      sh.getRange(row, 2).setValue(text);
      sh.getRange(row, 3).setValue(now);
    } else {
      sh.appendRow([body.key, text, now]);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
