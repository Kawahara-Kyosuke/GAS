//手順1:インポートしたExcelから値読み取り
function myFunctionAggregate() {
  //変数宣言
  var ss = SpreadsheetApp.getActiveSpreadsheet();//現在アクティブなスプレッドシートを取得
  var checks = ss.getSheetByName("プロジェクトNoチェックリスト"); //読み取り用シート(サンプル)
  var counts = ss.getSheetByName("count"); //確認用シート
  var range1=counts.getRange("B3:B15");
  var range2=counts.getRange("G2:G52");
  var range3=counts.getRange("B5:B6");
  var range4=counts.getRange("B9:B14");
  var range5=counts.getRange("A1:G52");
  
  //確認事項整理
  checks.getRange('E5').copyTo(counts.getRange(3, 2),{contentsOnly:true}) //名前
  checks.getRange(8, 5).copyTo(counts.getRange(4, 2),{contentsOnly:true}) //月度
  checks.getRange(13, 9).copyTo(counts.getRange(5, 2),{contentsOnly:true}) //支払時間
  checks.getRange(18, 4).copyTo(counts.getRange(6, 2),{contentsOnly:true}) //時間外
  checks.getRange('B34:B84').copyTo(counts.getRange('E2:E52'),{contentsOnly:true}) //日付
  checks.getRange('J34:J84').copyTo(counts.getRange('F2:F52'),{contentsOnly:true}) //プロジェクト名
  checks.getRange('AI34:AI84').copyTo(counts.getRange('G2:G52'),{contentsOnly:true}) //プロジェクト時間 

  if(!counts.getRange('B6').getValue()){    //時間外労働時間なしの場合、空白行に0時間を追記
    counts.getRange('B6').setValue('0')
  }
  
  //シートの表示形式変更
  range1.setNumberFormat('#,##0');
  range2.setNumberFormat('#,##0.00');
  range3.setNumberFormat('#,##0.00');  
  range4.setNumberFormat('#,##0.00');  
  range5.setFontFamily("Calibri"); 
  
}

//手順2.読み取った内容から労務費管理シートに入力
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
  var ColA = values3.flat().indexOf(name); //名前記載先頭行
  var valuesA = count_copy.getRange(ColA+1, 3, 14 ,1).getValues(); //名前準拠の検索範囲
  var Row1 = values[2].indexOf(month);　// 月(列)  
  var time1 = counts.getRange('B5').getValue(); //就業時間
  var time2 = counts.getRange('B6').getValue(); //時間外
  var Col1 = valuesA.flat().indexOf('就業時間（H）'); //就業時間の入力行検索
  var Col2 = valuesA.flat().indexOf('時間外労働時間（H）'); //時間外の入力行検索
  
  //手順2-1.就業時間入力
  var msg1 = Browser.msgBox(month + "月の勤務時間入力","就業時間（H）:" + time1 ,Browser.Buttons.OK_CANCEL);
  if(msg1 == "cancel"){
    Browser.msgBox("入力を中止しました。")
    return;
  }else{
    count_copy.getRange(ColA+Col1+1,Row1+1).setValue(time1);//実労働
  }

  //手順2-2.時間外労働時間入力
  var msg2 = Browser.msgBox(month + "月の勤務時間入力","時間外労働時間（H）:" + time2 ,Browser.Buttons.OK_CANCEL);
  if(msg2 == "cancel"){
    Browser.msgBox("入力を中止しました。")
    return;
  }else{
    count_copy.getRange(ColA+Col2+1,Row1+1).setValue(time2);//時間外
  }
  
  //手順2-3.労務費入力
  for(let i = 9; i <= 14; i++) {
    var KC1 = counts.getRange(i, 1).getValue(); 　//ツール上のKC番号
    var KC_Cost = counts.getRange(i, 3).getValues(); //労務費

    if(counts.getRange(i, 1).getValue() === '' ){      
      console.log(i, "NO")　　　　　　　　　　　　　　　//空白行の時ログ出力(確認用)    
    }else if ( valuesA.flat().indexOf(KC1)!= -1) {
      var msg3 = Browser.msgBox(month + "月労務費入力" ,"KC番号" + KC1 + "の労務費" + KC_Cost +"（千円）を管理表に入力します。" ,Browser.Buttons.OK);　　//一致するKC番号が管理表上に存在するとき
      var Col3 = valuesA.flat().indexOf(KC1); 　//KC(行)
      count_copy.getRange(ColA+Col3+1, Row1+1).setValue(KC_Cost); //時間入力        
    } else {
      var msg4 = Browser.msgBox(month + "月労務費入力"  ,KC1 + "が労務費管理表上に存在しません。",Browser.Buttons.OK);　　　 //一致するKC番号が存在しないとき
        }
  }
}


//手順3：シート及び入力内容の削除
function myFunctionReset() {
  //変数宣言
  var ss = SpreadsheetApp.getActive();// 現在アクティブなスプレッドシートを取得
  var counts = ss.getSheetByName("count"); //確認用シートを取得
  var deletesheet1 = ss.getSheetByName('プロジェクトNoチェックリスト');
  var deletesheet2 = ss.getSheetByName('勤怠票（備考)'); 

  var msg5 = Browser.msgBox("削除","入力内容、入力シートを削除します。" ,Browser.Buttons.OK_CANCEL);
  
  if(msg5 == "cancel"){
    var msg6 = Browser.msgBox("削除","削除を中止しました。" ,Browser.Buttons.OK)
    return;
  }else{
    counts.getRange('B3:B6').clearContent();  //集計シート上の入力値を削除
    counts.getRange('B9:B14').clearContent();  
    counts.getRange('E2:G52').clearContent();
    ss.deleteSheet(deletesheet1);  // 追加したシートを削除
    ss.deleteSheet(deletesheet2);  
    var msg7 = Browser.msgBox("削除","削除が完了しました。" ,Browser.Buttons.OK)
  } 
}
