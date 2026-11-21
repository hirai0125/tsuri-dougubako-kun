/**
 * 釣り道具箱くん 同期用 Apps Script
 * このコードを、同期に使うGoogleスプレッドシートの
 * 「拡張機能」→「Apps Script」に貼り付けて、ウェブアプリとしてデプロイします。
 *
 * ★ 下の PASSWORD の「ここに合言葉」を、自分で決めた合言葉に書き換えてから保存・デプロイしてください。
 *   アプリの「同期」画面にも同じ合言葉を入れます。
 *
 * 「同期データ」シート：アプリが使うデータ（触らないでください）
 * 「竿」「穂先」「リール」などのシート：確認用の一覧（同期のたびに書き直されます）
 */
// ▼▼▼ ここに合言葉を書く（' ' の内側だけ書き換える）▼▼▼
const PASSWORD = 'ここに合言葉';
// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

const DATA_SHEET = '同期データ';
const CHUNK = 40000; // 1セルに入れる文字数

function dataSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(DATA_SHEET) || ss.insertSheet(DATA_SHEET, 0);
}

function read_() {
  const sh = dataSheet_();
  const t = Number(sh.getRange('A1').getValue()) || 0;
  const last = sh.getLastRow();
  let json = '';
  if (last >= 2) {
    json = sh.getRange(2, 1, last - 1, 1).getValues()
      .map(function (r) { return String(r[0]).slice(1); }) // 先頭の「x」を外す
      .join('');
  }
  return { t: t, data: json ? JSON.parse(json) : null };
}

function okPw_(pw) {
  return PASSWORD && PASSWORD !== 'ここに合言葉' && String(pw || '') === PASSWORD;
}

function doGet(e) {
  if (!okPw_(e && e.parameter && e.parameter.pw)) return out_({ error: 'password' });
  return out_(read_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (!okPw_(body.pw)) return out_({ error: 'password' });
    const cur = read_();
    // サーバーの方が新しければ上書きしない（アプリ側で取り込み直す）
    if (cur.t && Number(body.t) < cur.t) return out_({ ok: false, t: cur.t });

    const sh = dataSheet_();
    sh.clear();
    const json = JSON.stringify(body.data);
    const rows = [];
    for (let i = 0; i < json.length; i += CHUNK) rows.push(['x' + json.slice(i, i + CHUNK)]);
    sh.getRange('A1:B1').setValues([[Number(body.t), new Date()]]);
    if (rows.length) sh.getRange(2, 1, rows.length, 1).setValues(rows);

    writeViews_(body.views || {});
    return out_({ ok: true, t: Number(body.t) });
  } finally {
    lock.releaseLock();
  }
}

/** 分類ごとの確認用シートを書き直す */
function writeViews_(views) {
  const ss = SpreadsheetApp.getActive();
  Object.keys(views).forEach(function (name) {
    const rows = views[name];
    if (!rows || !rows.length) return;
    const sheetName = String(name).replace(/[\[\]:*?\/\\]/g, '').slice(0, 90) || '分類';
    const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    sh.clear();
    const width = Math.max.apply(null, rows.map(function (r) { return r.length; }));
    const safe = rows.map(function (r) {
      const row = [];
      for (let i = 0; i < width; i++) {
        let v = r[i] === undefined || r[i] === null ? '' : r[i];
        if (typeof v === 'string' && /^[=+\-@]/.test(v)) v = "'" + v; // 数式として扱われないように
        row.push(v);
      }
      return row;
    });
    sh.getRange(1, 1, safe.length, width).setValues(safe);
    sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#e9f0f5');
    sh.setFrozenRows(1);
  });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
