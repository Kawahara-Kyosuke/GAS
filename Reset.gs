//手順3：シート及び入力内容の削除
function myFunctionReset() {
  //変数宣言
  var ss = SpreadsheetApp.getActive();// 現在アクティブなスプレッドシートを取得
  var counts = ss.getSheetByName("count"); //確認用シートを取得


  counts.getRange('B6:B9').clearContent();
  counts.getRange('D2:F52').clearContent();
  
  // そのスプレッドシートにある シート名:テスト用シート のシートを取得 
  var deletesheet1 = ss.getSheetByName('プロジェクトNoチェックリスト');
  var deletesheet2 = ss.getSheetByName('勤怠票（備考)'); 

  // そのシートを削除
  ss.deleteSheet(deletesheet1);
  ss.deleteSheet(deletesheet2);  
}