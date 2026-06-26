function addNewRowFromTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var configSheet = ss.getSheetByName("設定");
  
  // エラー時のポップアップ（メッセージボックス）も画面を止める原因になるため削除し、静かに終了させます
  if (!configSheet) return; 
  
  var lastRow = sheet.getLastRow();
  var nextRow = lastRow + 1;
  
  // 物理的に行を追加（下のボタンを押し下げるため）
  sheet.insertRowAfter(lastRow);
  
  // 「設定」シートの2行目を、本番シートの新行へ「行ごと」丸ごとコピー
  configSheet.getRange("2:2").copyTo(sheet.getRange(nextRow + ":" + nextRow));
}