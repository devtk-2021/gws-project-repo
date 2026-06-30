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

  // 追加した行の高さをデフォルトの2行分（34px）に設定
  sheet.setRowHeight(nextRow, 34);

  // B列（日付）とH列（次回作業予定日）の数式を値（実数）に変換して、動的に日付が変わらないようにする
  var cellB = sheet.getRange(nextRow, 2);
  cellB.setValue(cellB.getValue());

  var cellH = sheet.getRange(nextRow, 8);
  cellH.setValue(cellH.getValue());
}