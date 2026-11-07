// 三類法 評価支援ツール ログ受信API（v3対応）
// バインド先スプレッドシート: 三類法③評価ログ_運用テスト (1nIUwuNJConN2kCH7Eap_9s319_5hVxn8q8wzTokur7s)
// Web版 https://seichiku.github.io/sanruiho/ の「スプレッドシートに送信」から doPost で受け取る。
//   v0.3（payload.v === 3・2026-11-07 テスト運用〜）→「ログv3」タブ（43列＝v2の33列＋からだの傾向10列）
//   v0.2（payload.v === 2）→「ログv2」タブ（33列）※旧端末の未送信キュー用に残す
//   v0.1（v なし）→ 旧「ログ」タブ（22列）
// 更新手順：コードを差し替え→「デプロイを管理→編集→新バージョン」でURL固定のまま反映。初回は setupV3() を1回実行。

// v0.3 用：「ログv3」タブと使い方の整備（何度実行しても壊れない。ログv2・ログ タブには触れない）
function setupV3() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  var sh = ss.getSheetByName(LOG3_SHEET);
  if (!sh) sh = ss.insertSheet(LOG3_SHEET, 0);
  sh.getRange(1, 1, 1, HEADERS3.length).setValues([HEADERS3]).setFontWeight("bold").setBackground("#e8f0fe");
  sh.getRange(1, 34, 1, 10).setBackground("#fff3cd");   // v3で増えた10列は薄黄
  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);
  var widths = {1:130, 2:110, 3:70, 4:110, 5:90, 7:160, 11:150, 12:130, 15:260, 17:90, 18:90, 19:260, 23:300, 27:90, 31:300, 32:90, 33:300, 34:90, 35:90, 41:110, 42:90, 43:80};
  for (var c in widths) sh.setColumnWidth(+c, widths[c]);
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(["未対応","確認中","反映済","見送り"], true).setAllowInvalid(true).build();
  sh.getRange(2, 32, sh.getMaxRows() - 1, 1).setDataValidation(rule);
  sh.getRange(2, 31, sh.getMaxRows() - 1, 3).setWrap(true);
  var guide = ss.getSheetByName("使い方");
  if (!guide) guide = ss.insertSheet("使い方");
  guide.insertRowBefore(1);
  guide.getRange(1, 1).setValue("【2026-11-07〜 テスト運用】ツールからの送信は「ログv3」タブに入ります（v2の33列＋AH〜AQ列：顔色・欲しくなる味・肝・心・脾・肺・腎・からだの傾向・提案した色・QR表示。v0.3実装までは空欄）。「ログv2」は10/20〜11/7の分で保管用。").setFontWeight("bold").setWrap(true);
  Logger.log("setupV3 done: rows=" + (sh.getLastRow() - 1));
}

var LOG_SHEET = "ログ";
var LOG2_SHEET = "ログv2";
var LOG3_SHEET = "ログv3";
var TZ = "Asia/Tokyo";
var HEADERS = ["受信日時","評価日時","記録者","部位","陰陽","六道","確定病位","推奨M","上焦点","中焦点","下焦点","失調","レッドフラッグ","陰陽の回答","六道の回答","三焦チェック","判定は妥当か","正しいと思う病位","正しいと思うM","コメント","対応状況","対応メモ"];
var HEADERS2 = ["受信日時","評価日時","記録者","患者名","部位","左右","主訴","発症期間","きっかけ","経過",
  "対象組織","時期","陰のサイン","レッドフラッグ","陰陽の回答","推奨難度(陰陽)","分岐","分岐の推奨",
  "六道の回答","六道の乱れ数","最も大きい制限","推奨難度(六道)",
  "三焦の回答","三焦の警告","最も妨げている機能","推奨難度(三焦)",
  "判定は妥当か","正しいと思う組織","正しいと思う時期","正しいと思うM","コメント","対応状況","対応メモ"];
var HEADERS3 = HEADERS2.concat(["顔色","欲しくなる味","肝","心","脾","肺","腎","からだの傾向","提案した色","QR表示"]);

// v0.2 用：「ログv2」タブと使い方の整備（何度実行しても壊れない。旧「ログ」タブには触れない）
function setupV2() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  var sh = ss.getSheetByName(LOG2_SHEET);
  if (!sh) sh = ss.insertSheet(LOG2_SHEET, 0);
  sh.getRange(1, 1, 1, HEADERS2.length).setValues([HEADERS2]).setFontWeight("bold").setBackground("#e8f0fe");
  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);
  var widths = {1:130, 2:110, 3:70, 4:110, 5:90, 7:160, 11:150, 12:130, 15:260, 17:90, 18:90, 19:260, 23:300, 27:90, 31:300, 32:90, 33:300};
  for (var c in widths) sh.setColumnWidth(+c, widths[c]);
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(["未対応","確認中","反映済","見送り"], true).setAllowInvalid(true).build();
  sh.getRange(2, 32, sh.getMaxRows() - 1, 1).setDataValidation(rule);
  sh.getRange(2, 31, sh.getMaxRows() - 1, 3).setWrap(true);
  var guide = ss.getSheetByName("使い方");
  if (!guide) guide = ss.insertSheet("使い方");
  guide.clear();
  guide.getRange(1, 1, 10, 1).setValues([
    ["三類法 評価支援ツール 運用テストのログ（v0.2＝「ログv2」タブ／v0.1＝「ログ」タブは保管用）"],
    ["Web版 https://seichiku.github.io/sanruiho/ で評価→「記録・フィードバック」欄に記録者名・判定への評価・コメントを入れて「スプレッドシートに送信」を押すと「ログv2」タブに1行入ります。"],
    ["D列 患者名はツールで入力したもの。K〜P列が陰陽（対象組織・時期・陰のサイン・レッドフラッグ・時期の回答4項目・推奨難度）。Q〜R列が分岐（施術者の選択と、ツールの推奨）。"],
    ["S〜V列が六道（運動リズム5つの回答・乱れ数・最も大きい制限・推奨難度）、W〜Z列が三焦（該当項目・警告徴候・最も妨げている機能・推奨難度）。六道・三焦へ進まなかった時は空欄。"],
    ["AA列「判定は妥当か」：妥当／迷う／違う。AB〜AD列は「違う」と思った時の正しい組織・時期・難度M。AE列がコメント。"],
    ["AF列「対応状況」・AG列「対応メモ」は竹中／Claudeが記入します（未対応→確認中→反映済／見送り）。"],
    ["行の削除はしないでください。間違えて送った行は AG列に「誤送信」と書けばOKです。"],
    ["判定ルール（時期の多数決・難度の対応）はスプシ②「判定ロジック表」v0.2タブ。訓練プログラム枠はスプシ①の v0.2 タブ。正本は Notion「三類法」（スタッフ共有）。"],
    ["関連：①M施術たたき台_承認シート／②判定ロジック表_たたき台／Notion「三類法 方程式 一覧＋課題たたき台」"],
    ["v0.2 作成 2026-10-20（v0.1 2026-09-17）"]
  ]);
  guide.getRange(1, 1).setFontWeight("bold").setFontSize(13);
  guide.setColumnWidth(1, 900);
  guide.getRange(1, 1, 10, 1).setWrap(true);
  Logger.log("setupV2 done: rows=" + (sh.getLastRow() - 1));
}

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    if (p.v === 3) return appendV3_(p);
    if (p.v === 2) return appendV2_(p);
    var sh = getSheet_(LOG_SHEET, HEADERS);
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

function rowV2_(p) {
  var s = function (v) { return (v === undefined || v === null) ? "" : String(v); };
  return [
    Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss"),
    s(p.date), s(p.recorder), s(p.ptName), s(p.bui), s(p.side), s(p.complaints), s(p.period), s(p.trigger), s(p.course),
    s(p.tissue), s(p.phase), s(p.inSigns), s(p.red), s(p.inyoAns), s(p.inyoM), s(p.branch), s(p.branchRec),
    s(p.rikudoAns), s(p.rikudoN), s(p.rikudoMain), s(p.rikudoM),
    s(p.sanshoAns), s(p.sanshoWarn), s(p.sanshoMain), s(p.sanshoM),
    s(p.fbValid), s(p.fbTissue), s(p.fbPhase), s(p.fbM), s(p.comment),
    "未対応", ""
  ];
}

function appendV2_(p) {
  var sh = getSheet_(LOG2_SHEET, HEADERS2);
  sh.appendRow(rowV2_(p));
  return json_({ ok: true, count: sh.getLastRow() - 1, sheet: LOG2_SHEET });
}

// v0.3: v2の33列＋からだの傾向10列（ツール側が未実装の間は空欄）
function appendV3_(p) {
  var sh = getSheet_(LOG3_SHEET, HEADERS3);
  var s = function (v) { return (v === undefined || v === null) ? "" : String(v); };
  var g = p.gogyo || {};
  var row = rowV2_(p).concat([
    s(p.faceColor), s(p.taste), s(g.kan), s(g.shin), s(g.hi), s(g.hai), s(g.jin), s(p.gogyoTop), s(p.foodColor), s(p.qrShown)
  ]);
  sh.appendRow(row);
  return json_({ ok: true, count: sh.getLastRow() - 1, sheet: LOG3_SHEET });
}

// 疎通確認用：件数を返す
function doGet(e) {
  var sh = getSheet_(LOG3_SHEET, HEADERS3);
  return json_({ ok: true, count: sh.getLastRow() - 1, sheet: LOG3_SHEET });
}

function getSheet_(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name, 0); }
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}
function num_(v) { return (v === undefined || v === null || v === "") ? "" : Number(v); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
