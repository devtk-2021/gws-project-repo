function onEdit(e) {
  // 【高速化の工夫1】関係のない列（F列以外）や1行目が編集されたら、1ミリ秒も無駄にせず即座に終了する
  if (e.range.columnStart !== 6 || e.range.rowStart < 2) return;
  
  // 【高速化の工夫2】スプレッドシートとの通信回数を減らすため、イベントデータ(e)から直接値を取得
  var sheet = e.range.getSheet();
  var status = e.value; // 編集されたセルの値
  var dateCell = sheet.getRange(e.range.rowStart, 7); // G列のセル
  
  if (status) {
    // ステータスが入力・変更された場合
    dateCell.setValue(new Date());
  } else {
    // ステータスが削除されて空欄になった場合
    dateCell.clearContent();
  }
}