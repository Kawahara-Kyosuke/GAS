//手順1:インポートしたExcelから値読み取り
function myFunctionAggregate() {
  //　変数宣言
  var ss = SpreadsheetApp.getActiveSpreadsheet();//現在アクティブなスプレッドシートを取得
  var checks = ss.getSheetByName("プロジェクトNoチェックリスト"); //読み取り用シート(サンプル)
  var counts = ss.getSheetByName("count"); //確認用シート

  //確認事項整理
  checks.getRange('E5').copyTo(counts.getRange(6, 2),{contentsOnly:true}) //名前
  checks.getRange(8, 5).copyTo(counts.getRange(7, 2),{contentsOnly:true}) //月度
  checks.getRange(13, 9).copyTo(counts.getRange(8, 2),{contentsOnly:true}) //支払時間
  checks.getRange(18, 4).copyTo(counts.getRange(9, 2),{contentsOnly:true}) //時間外
  checks.getRange('B34:B84').copyTo(counts.getRange('D2:D52'),{contentsOnly:true}) //日付
  checks.getRange('J34:J84').copyTo(counts.getRange('E2:E52'),{contentsOnly:true}) //プロジェクト名
  checks.getRange('AI34:AI84').copyTo(counts.getRange('F2:F52'),{contentsOnly:true}) //プロジェクト時間 

  //表示形式変更
  var range1=counts.getRange("B7:B15");
  var range2=counts.getRange("F2:F32");
  var range3=counts.getRange("A1:F51");
  range1.setNumberFormat('#,##0.00');
  range2.setNumberFormat('#,##0.00');
  range3.setFontFamily("Calibri");

}
