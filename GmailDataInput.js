//シート名「input」を指定。
var sheet = SpreadsheetApp.getActive().getSheetByName("input");

//データを保存するフォルダ先のIDを取得。
var destFolderId = sheet.getRange("XX").getValue(); //保存するフォルダのフォルダIDをシートから取得

//情報取得
function entryInput() {

  //指定のメールラベル登録されたメールから最新のスレッド〇件までを抽出。
  var query = sheet.getRange("XX").getValue();//ラベルクエリをシートから取得
  var maxThreads = sheet.getRange("XX").getValue();//読み込むスレッドの最大数をシートから取得
  var threads = GmailApp.search(query, 0, maxThreads);//クエリで抽出されるスレッドから最大数個のスレッド

  //フォルダIDからフォルダオブジェクトを取得。
  var destFolder = DriveApp.getFolderById(destFolderId); //フォルダを取得

  //抽出したスレッドからメール内容を読み込み。配列messages = [スレッド数i , スレッド内のメール数j]
  var messages = GmailApp.getMessagesForThreads(threads);//

  //スレッド数iとスレッド内のメール数ｊの繰り返し処理でメール文面を全て抽出。
  var arrMessages = messages.reverse();
  Logger.log(arrMessages);
  for(var i=0; i < arrMessages.length; i++) {
    Logger.log(arrMessages);
    for(var j=0; j < arrMessages[i].length; j++) {
        var body = arrMessages[i][j].getPlainBody();

        //メッセージ毎にメッセージIDを取得
        var id = arrMessages[i][j].getId();
        //新しいメッセージIDかどうか判定
        if(!hasId(id)) {
            //メッセージ毎に受信日時を取得
            var massageDate = arrMessages[i][j].getDate();

            //氏名だけ先に取り出す。
            var entryName = fetchData(body,"氏名 ] ","¥r");

            //destフォルダの子フォルダを全て取得
            var destFolder_childs = destFolder.getFolders();

            //「”氏名”_書類」のフォルダを作成
            var folderName = entryName + "_" + "書類";
            var newFolder = destFolder.createFolder(folderName);

            //「”氏名”_書類」に添付ファイルを取得。
            var attachments = arrMessages[i][j].getAttachments();
            for(var k in attachments){
              newFolder.createFile(attachments[k]);
            }

            //該当フォルダのURLを取得
            var folderUrl = newFolder.getUrl();
              //抽出したメール文面から、項目名を省いて必要な個人情報のみを抽出し、「input」シートの最終行に追加していく。
              sheet.appendRow([
                entryName,
                fetchData(body, "〇〇〇〇 ] ", "¥r"),
                fetchData(body, "〇〇〇〇 ] ", "¥r"),
                fetchData(body, "〇〇〇〇 ] ", "¥r"),
                fetchData(body, "〇〇〇〇 ] ", "¥r"),
                fetchData(body, "〇〇〇〇 ] ", "¥r"),
                id,
                folderUrl,
                massageDate
                //fetchData(body, "〇〇〇〇 ]", "¥r"));
              ]);
        }

    //messages[i][j].markRead();//メッセージを既読にする。--未読判定時
    }
  }
}


//文章から特定部分を削除する関数
function fetchData(str, pre, suf) {

  //preで始まりsufで終わる文を正規表現でregとして宣言。
  var reg = new RegExp(pre + '.*?' + suf);

  //
  var matchtxt = str.match(reg);

  if(matchtxt === null){
    return null;
  } else {
    //文章strから文regに完全一致する箇所を探しだし、preとsufに指定した箇所を空置換によって削除。
    return matchtxt[0].replace(pre,"").replace(suf,"");
  }
}

//既に取得済のメッセージIDと比較し、同じIDが一つでもあればtrueを返す関数
function hasId(id) {
  var rowCnt = sheet.getLastRow();
  if(rowCnt > 4) {
    var idData = sheet.getRange(5, 7, rowCnt - 4).getValues();//5行目7列目のセルから、（rowCnt - 4）行下のセルまでを選択
    //Logger.log(sheet.getLastRow() - 4);
    //Logger.log(idData);
    var hasId = idData.some(function(array,i,idData) {
      return (array[0]===id);
    });
    return hasId;
  } else {
    return false;
    }
}
