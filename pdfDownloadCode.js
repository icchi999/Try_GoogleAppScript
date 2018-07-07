/*
＜対象＞スプレッドシート
＜概要＞
特定のシートの指定セルに1~31までの日付を順番に入力し、
それぞれの日付毎に該当シートをPDF化、それを指定フォルダに保存するGAS.

＜注意＞
・指定セルの日付を書き換える処理に続けてすぐにPDF化をかけると処理が間に合わず
うまくPDF保存できない（エラー）ため、日付セット後に7秒の遅延処理を挿入。

・上記エラーが起きても途中で処理を止めないように、
UrlFetchApp.fetch()で、muteHttpExceptions:trueをセット。

・7秒は、実際に使用したスプレッドシートで何度か試してエラーが起きないだいたいの秒数。
Gasでは実行時間が５分以上になると強制終了するため、長すぎてもNG

・function onOpenでスプレッドシートのメニューに「スクリプト」を追加し、
そこから実行できるようにした
*/

function inputDate() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssid = ss.getId();

  var sheet = SpreadsheetApp.getActive().getSheetByName("業務日報");
  var sheetid = sheet.getSheetId();

  var folderid = "15RGT36Nd0elJpIby0QhFTweQPjm2bpl9";

  for(var i=1; i < 32; i++) {
    sheet.getRange(5, 4).setValue(i);
    Logger.log(sheet.getRange(5, 4).getValue());

    Utilities.sleep(7000);

    var folder = DriveApp.getFolderById(folderid);
    var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);
    var opts = {
      exportFormat: "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
      format:       "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
      size:         "A4",     // 用紙サイズの指定 legal / letter / A4
      portrait:     "true",   // true → 縦向き、false → 横向き
      fitw:         "true",   // 幅を用紙に合わせるか
      sheetnames:   "false",  // シート名をPDF上部に表示するか
      printtitle:   "false",  // スプレッドシート名をPDF上部に表示するか
      pagenumbers:  "false",  // ページ番号の有無
      gridlines:    "false",  // グリッドラインの表示有無
      fzr:          "false",  // 固定行の表示有無
      gid:          sheetid   // シートIDを指定 sheetidは引数で取得
    };

    var url_ext = [];
    for( optName in opts ){
      url_ext.push( optName + "=" + opts[optName] );
    }
    var options = url_ext.join("&");
    Logger.log(options);
    var token = ScriptApp.getOAuthToken();

    var response = UrlFetchApp.fetch(url + options, {
       headers: {
       'Authorization': 'Bearer ' +  token
       }
       ,muteHttpExceptions:true
    });

    var sheet_yyyymm = SpreadsheetApp.getActive().getSheetByName("勤務表");
    var yyyymm = sheet_yyyymm.getRange(1, 1).getValue();
    var yyyymmString = Utilities.formatDate(yyyymm,"JST","yyyy.MM");
    var filename = yyyymmString + "." + i + "_業務日報";

    var blob = response.getBlob().setName(filename + '.pdf');
    folder.createFile(blob);
  }
}

function onOpen() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var entries = [
        {
            name : "業務日報をドライブフォルダにPDF保存",
            functionName : "inputDate"
        }
        ];
    sheet.addMenu("スクリプト", entries);
};

function test() {
  var sheet_yyyymm = SpreadsheetApp.getActive().getSheetByName("勤務表");
  var yyyymm = sheet_yyyymm.getRange(1, 1).getValue();
  var yyyymmString = Utilities.formatDate(yyyymm,"JST","yyyy.MM");
  Logger.log(yyyymmString);
}
