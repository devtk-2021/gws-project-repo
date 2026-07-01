function handleEdit(e) {
  try {
    if (!e) return;
    
    var range = e.range;
    var sheet = range.getSheet();
    var sheetName = sheet.getName();
    var a1Notation = range.getA1Notation();
    
    // 対象のシート名が「収支（実績）」で、セルが「F3」の場合のみ実行
    if (sheetName === "収支（実績）" && a1Notation === "F3") {
      var facilityName = range.getValue(); // 選択された施設名
      
      // 施設名が空にされた場合は、9行目以下のA〜P列をクリアして終了
      if (!facilityName) {
        var lastRow = sheet.getLastRow();
        if (lastRow >= 9) {
          sheet.getRange(9, 1, lastRow - 8, 16).clear(); // A9:Pの範囲をクリア
        }
        return;
      }
      
      // データインポートを実行
      importData(sheet, facilityName);
    }
  } catch (error) {
    console.error("handleEditでエラーが発生しました: " + error.toString());
  }
}

// データインポートの実処理関数
function importData(sheet, facilityName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = null;
  var loadingCell = sheet.getRange("A9");
  
  try {
    // 1. 画面右下に処理中のトースト表示（時間を-1にして処理が終わるまで残す）
    ss.toast(facilityName + " のデータを取得しています...", "データ読み込み中", -1);
    
    // 2. 9行目から下のA〜P列（16列分）をクリア
    var lastRow = sheet.getLastRow();
    if (lastRow >= 9) {
      sheet.getRange(9, 1, lastRow - 8, 16).clear();
    }
    
    // 3. A9セルに一時的な処理中メッセージを表示（結合はエラーの原因になるため行わない）
    loadingCell.setValue("【データ取得中...】 " + facilityName + " のデータを読み込んでいます。このままお待ちください。");
    loadingCell.setFontColor("#ef4444"); // 赤字
    loadingCell.setFontWeight("bold"); // 太字
    
    // スプレッドシートに一時的な表示を即座に反映させる
    SpreadsheetApp.flush();
    
    // 4. URL一覧シートから対象施設のURLを検索
    var masterSsUrl = "https://docs.google.com/spreadsheets/d/1PvsAhiLcpI8174QVLBelJFulMPxXvnSCDOJ8B4hZTlU/edit";
    var masterSs = SpreadsheetApp.openByUrl(masterSsUrl);
    var masterSheet = masterSs.getSheetByName("各施設リンク先一覧 のコピー");
    if (!masterSheet) {
      throw new Error("マスタシート「各施設リンク先一覧 のコピー」が見つかりませんでした。");
    }
    
    var masterLastRow = masterSheet.getLastRow();
    var masterData = masterSheet.getRange(1, 1, masterLastRow, 4).getValues(); // A列〜D列を取得
    var targetUrl = "";
    
    // ループで施設名（B列 = インデックス1）を探す
    for (var i = 0; i < masterData.length; i++) {
      if (masterData[i][1] === facilityName) {
        // スマートチップ形式と生URLテキストの両方に対応させてURLを抽出
        var cell = masterSheet.getRange(i + 1, 4); // D列（4列目）
        var richText = cell.getRichTextValue();
        if (richText) {
          targetUrl = richText.getLinkUrl() || masterData[i][3];
        } else {
          targetUrl = masterData[i][3];
        }
        break;
      }
    }
    
    // URLが見つからない場合はエラー
    if (!targetUrl || targetUrl === "") {
      throw new Error("「" + facilityName + "」のURLが一覧シートから見つかりませんでした。");
    }
    
    // 5. データソースからデータを丸ごとコピー＆ペースト
    var srcSs = SpreadsheetApp.openByUrl(targetUrl);
    var srcSheet = srcSs.getSheetByName("シート1") || srcSs.getSheets()[0];
    
    // 別スプレッドシートのシートを現在のスプレッドシートに一時コピー（同じファイル内にする）
    tempSheet = srcSheet.copyTo(ss);
    
    var srcLastRow = tempSheet.getLastRow();
    var srcRange = tempSheet.getRange("A1:P" + srcLastRow);
    
    // 貼り付け前にA9セルの一時メッセージ用書式を元に戻す
    try {
      loadingCell.setFontColor(null);
      loadingCell.setFontWeight(null);
    } catch (e) {
      console.error("A9書式のリセットに失敗しました: " + e.toString());
    }
    
    // 出力先シートの9行目、1列目（A9セルの位置）を貼り付けの起点にする
    var destRange = sheet.getRange(9, 1);
    
    // コピー元の数式・書式・値をすべて丸ごと上書きコピー
    srcRange.copyTo(destRange);
    
    // 6. 完了トーストを表示
    ss.toast(facilityName + " のデータを反映しました。", "読込完了", 3);
    
  } catch (error) {
    // エラー発生時もA9セルの書式を元に戻す
    try {
      loadingCell.setFontColor(null);
      loadingCell.setFontWeight(null);
    } catch (e) {}
    
    ss.toast("エラーが発生しました。", "読込失敗", 3);
    SpreadsheetApp.getUi().alert("エラーが発生しました:\n" + error.toString());
    
  } finally {
    // 一時シートは確実に削除する
    if (tempSheet) {
      try {
        ss.deleteSheet(tempSheet);
      } catch (e) {
        console.error("一時シートの削除に失敗しました: " + e.toString());
      }
    }
  }
}