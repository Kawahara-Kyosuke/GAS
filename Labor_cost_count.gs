//手順1:インポートしたExcelから値読み取り
function myFunctionAggregate() {
  //　変数宣言
  var ss = SpreadsheetApp.getActiveSpreadsheet();//現在アクティブなスプレッドシートを取得
  var checks = ss.getSheetByName("プロジェクトNoチェックリスト"); //読み取り用シート(サンプル)
  var counts = ss.getSheetByName("count"); //確認用シート

  //確認事項整理
  checks.getRange('E5').copyTo(counts.getRange(3, 2),{contentsOnly:true}) //名前
  checks.getRange(8, 5).copyTo(counts.getRange(4, 2),{contentsOnly:true}) //月度
  checks.getRange(13, 9).copyTo(counts.getRange(5, 2),{contentsOnly:true}) //支払時間
  checks.getRange(18, 4).copyTo(counts.getRange(6, 2),{contentsOnly:true}) //時間外
  checks.getRange('B34:B84').copyTo(counts.getRange('D2:D52'),{contentsOnly:true}) //日付
  checks.getRange('J34:J84').copyTo(counts.getRange('E2:E52'),{contentsOnly:true}) //プロジェクト名
  checks.getRange('AI34:AI84').copyTo(counts.getRange('F2:F52'),{contentsOnly:true}) //プロジェクト時間 

  //表示形式変更
  var range1=counts.getRange("B3:B14");
  var range2=counts.getRange("F2:F52");
  var range3=counts.getRange("A1:F52");
  range1.setNumberFormat('#,##0.00');
  range2.setNumberFormat('#,##0.00');
  range3.setFontFamily("Calibri");

}

//手順2.読み取った内容から労務費管理シートに入力
//入力セル検索
function search() {
  //変数
  const ss = SpreadsheetApp.getActiveSpreadsheet(); //現在アクティブなスプレッドシートを取得
  var ss_copy = SpreadsheetApp.openById('1RfVjjTiKCawDfir-bjYab2T-s1c7YFwHXPxUUAaghgc'); //コピー先のスプレッドシートを取得
  var counts = ss.getSheetByName('count'); //確認用シートを取得
  const count_copy = ss_copy.getSheetByName('KC労務費管理表_最新'); //確認用シートを取得
  const values = count_copy.getDataRange().getValues();
  const values2 = count_copy.getRange(1, 3, count_copy.getLastRow() - 1).getValues(); 
  const values3 = count_copy.getRange(1, 2, count_copy.getLastRow() - 1).getValues(); 
  var name = counts.getRange('B3').getValue();  //名前指定
  var month = counts.getRange('B4').getValue(); //月指定
　var test = counts.getRange('A8').getValue();
  var ColA = values3.flat().indexOf(name); //名前記載先頭行
  var valuesA = count_copy.getRange(ColA+1, 3, 10 ,1).getValues(); //名前準拠の検索範囲
  var Row1 = values[2].indexOf(month);　// 月(列)  
          
  for(let i = 8; i <= 13; i++) {
    var KC1 = counts.getRange(i, 1).getValue(); 　//ツール上のKC番号

    if(counts.getRange(i, 1).getValue() === '' ){      
      console.log(i, "NO")　　　　　　　　　　　　　　　//空白行の時ログ出力(確認用)
     
    }else if ( valuesA.flat().indexOf(KC1)!= -1) {
      Browser.msgBox(KC1 + "の稼働時間を労務費管理表に入力しました。");　　//一致するKC番号が管理表上に存在するとき
      
      var Col1 = valuesA.flat().indexOf(KC1); 　//KC(行)
      var KC_time = counts.getRange(i, 2).getValues(); //時間
      count_copy.getRange(ColA+Col1+1, Row1+1).setValue(KC_time); //時間入力  
      
    } else {
          Browser.msgBox(KC1 + "が労務費管理表上に存在しません。");　　　 //一致するKC番号が存在しないとき
        }
  }
}


//手順3：シート及び入力内容の削除
function myFunctionReset() {
  //変数宣言
  var ss = SpreadsheetApp.getActive();// 現在アクティブなスプレッドシートを取得
  var counts = ss.getSheetByName("count"); //確認用シートを取得


  counts.getRange('B3:B13').clearContent();
  counts.getRange('D2:F52').clearContent();
  
  // そのスプレッドシートにある シート名:テスト用シート のシートを取得 
  var deletesheet1 = ss.getSheetByName('プロジェクトNoチェックリスト');
  var deletesheet2 = ss.getSheetByName('勤怠票（備考)'); 

  // そのシートを削除
  ss.deleteSheet(deletesheet1);
  ss.deleteSheet(deletesheet2);  
}
