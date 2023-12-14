// 連絡事項シート用
function insertRowBeforeWithDate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var today　=  new Date(); 　
  var formatDate = Utilities.formatDate(today, "JST", "yyyy/MM/dd"); 
  sheet.getRange("B2").setValue(formatDate);
  sheet.getRange("A2").setValue('未発表');
  // これは最初の行位置の前に行を挿入します
  sheet.insertRowBefore(2);
}

// 問題解決共有用
function insertRowBeforeWithDateAndStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var today　=  new Date(); 　
  var formatDate = Utilities.formatDate(today, "JST", "yyyy/MM/dd"); 
  sheet.getRange("B2").setValue(formatDate);
  sheet.getRange("A2").setValue('未発表');
  // これは最初の行位置の前に行を挿入します
  sheet.insertRowBefore(2);
}

// リリースコーナー用
function insertRowBeforeWithStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.getRange("A2").setValue('未発表');
  // これは最初の行位置の前に行を挿入します
  sheet.insertRowBefore(2);
}

// 持ち込み用
function insertRowBeforeWithStatus5() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.getRange("A5").setValue('未発表');
  // これは最初の行位置の前に行を挿入します
  sheet.insertRowBefore(5);
}

// 今週の感じシート作成
function createGoodBadSheet() {
  // 日付の取得
  const formatDate = Utilities.formatDate(new Date(), "JST","yyMMdd");

  // スクリプトに紐付いたスプレッドシートを読み込む
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // シートの有無をチェック
  if (spreadsheet.getSheetByName("今週の感じ(" + formatDate + ")") !== null) {
    return;
  }
  // コピー元になるシートを取得
  const baseSheet = spreadsheet.getSheetByName("★今週の感じ(yyMMdd)");

  //コピー対象シートを同一のスプレッドシートにコピー
  let newsheet = baseSheet.copyTo(spreadsheet);

  //シートのリネーム
  newsheet.setName("今週の感じ(" + formatDate + ")");
}



