// スプレッドシート上のボタンから起動される関数
function runImportFromButton() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("収支（実績）");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("「収支（実績）」シートが見つかりませんでした。");
      return;
    }
    
    var facilityName = sheet.getRange("F3").getValue(); // 選択された施設名
    
    if (!facilityName) {
      SpreadsheetApp.getUi().alert("F3セルに施設名が選択されていません。");
      return;
    }
    
    // ローディングモーダルを表示
    var template = HtmlService.createTemplateFromFile('loading');
    template.facilityName = facilityName;
    
    var htmlOutput = template.evaluate()
        .setWidth(360)
        .setHeight(240);
        
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("実行中にエラーが発生しました:\n" + error.toString());
  }
}

// モーダル（loading.html）から非同期で呼び出されるデータインポートの実処理関数
function runImport(facilityName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("収支（実績）");
  if (!sheet) {
    throw new Error("「収支（実績）」シートが見つかりませんでした。");
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = null;
  
  try {
    // 1. 9行目から下のA〜P列（16列分）のみをクリア
    var lastRow = sheet.getLastRow();
    if (lastRow >= 9) {
      sheet.getRange(9, 1, lastRow - 8, 16).clear();
    }
    
    // 2. URL一覧シートから対象施設のURLを検索
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
    
    // 3. データソースからデータを丸ごとコピー＆ペースト
    var srcSs = SpreadsheetApp.openByUrl(targetUrl);
    var srcSheet = srcSs.getSheetByName("シート1") || srcSs.getSheets()[0];
    
    // 別スプレッドシートのシートを現在のスプレッドシートに一時コピー（同じファイル内にする）
    tempSheet = srcSheet.copyTo(ss);
    
    var srcLastRow = tempSheet.getLastRow();
    var srcRange = tempSheet.getRange("A1:P" + srcLastRow);
    
    // 出力先シートの9行目、1列目（A9セルの位置）を貼り付けの起点にする
    var destRange = sheet.getRange(9, 1);
    
    // コピー元の数式・書式・値をすべて丸ごと上書きコピー
    srcRange.copyTo(destRange);
    
    // 4. 完了トーストを表示
    ss.toast(facilityName + " のデータを反映しました。", "読込完了", 3);
    
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