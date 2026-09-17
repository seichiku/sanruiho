// 三類法 評価支援ツール ログ受信API
// バインド先スプレッドシート: 三類法③評価ログ_運用テスト (1nIUwuNJConN2kCH7Eap_9s319_5hVxn8q8wzTokur7s)
// Web版 https://seichiku.github.io/sanruiho/ の「スプレッドシートに送信」ボタンから doPost で受け取り、「ログ」シートに1行追記する。

var LOG_SHEET = "ログ";
var TZ = "Asia/Tokyo";
var HEADERS = ["受信日時","評価日時","記録者","部位","陰陽","六道","確定病位","推奨M","上焦点","中焦点","下焦点","失調","レッドフラッグ","陰陽の回答","六道の回答","三焦チェック","判定は妥当か","正しいと思う病位","正しいと思うM","コメント","対応状況","対応メモ"];

// 初回に1回だけ実行：シート名・見出し・固定行・入力規則・使い方タブを整える（何度実行しても壊れない）
function setup() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  var sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(LOG_SHEET); }
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold").setBackground("#e8f0fe");
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);
  var widths = {1:130, 2:110, 3:90, 4:50, 14:260, 15:260, 16:260, 17:90, 18:130, 19:90, 20:300, 21:90, 22:300};
  for (var c in widths) sh.setColumnWidth(+c, widths[c]);
  // 対応状況（U列）は竹中／Claudeが使うプルダウン
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(["未対応","確認中","反映済","見送り"], true).setAllowInvalid(true).build();
  sh.getRange(2, 21, sh.getMaxRows() - 1, 1).setDataValidation(rule);
  sh.getRange(2, 20, sh.getMaxRows() - 1, 3).setWrap(true);
  // 使い方タブ
  var guide = ss.getSheetByName("使い方");
  if (!guide) guide = ss.insertSheet("使い方");
  guide.clear();
  guide.getRange(1, 1, 8, 1).setValues([
    ["三類法 評価支援ツール 運用テストのログ"],
    ["Web版 https://seichiku.github.io/sanruiho/ で判定→「記録」欄に記録者名・判定への評価・コメントを入れて「スプレッドシートに送信」を押すと「ログ」タブに1行入ります。"],
    ["Q列「判定は妥当か」：妥当／違う／迷う（ツール側で選択）。R・S列は「違う」と思った時の正しい病位・M（ツール側で選択）。T列がコメント。"],
    ["N〜P列は判定の元になった回答（陰陽5項目・六道5項目・三焦チェック）。閾値の見直し（課題D）に使います。"],
    ["U列「対応状況」・V列「対応メモ」は竹中／Claudeが記入します（未対応→確認中→反映済／見送り）。"],
    ["行の削除はしないでください。間違えて送った行は V列に「誤送信」と書けばOKです。"],
    ["関連：①M施術たたき台_承認シート／②判定ロジック表_たたき台／Notion「三類法 方程式 一覧＋課題たたき台 v0.1」"],
    ["作成 2026-09-17"]
  ]);
  guide.getRange(1, 1).setFontWeight("bold").setFontSize(13);
  guide.setColumnWidth(1, 900);
  guide.getRange(1, 1, 8, 1).setWrap(true);
  Logger.log("setup done: rows=" + (sh.getLastRow() - 1));
}

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    var sh = getLogSheet_();
    var row = [
      Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss"),
      p.date || "", p.recorder || "", p.bui || "", p.inyo || "", p.rikudo || "", p.final || "", p.mRange || "",
      num_(p.sanshoU), num_(p.sanshoC), num_(p.sanshoG), p.shittyo || "", p.red || "",
      p.inyoAns || "", p.rikudoAns || "", p.sanshoChecked || "",
      p.fbValid || "", p.fbStage || "", p.fbM || "", p.comment || "",
      "未対応", ""
    ];
    sh.appendRow(row);
    return json_({ ok: true, count: sh.getLastRow() - 1 });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// 疎通確認用：件数を返す
function doGet(e) {
  var sh = getLogSheet_();
  return json_({ ok: true, count: sh.getLastRow() - 1 });
}

function getLogSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(LOG_SHEET); }
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}
function num_(v) { return (v === undefined || v === null || v === "") ? "" : Number(v); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
