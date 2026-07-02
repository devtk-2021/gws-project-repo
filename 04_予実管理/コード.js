// スプレッドシートを開いたときに起動する関数
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var ui = SpreadsheetApp.getUi();
  
  // カスタムメニューを追加
  ui.createMenu('カスタム機能')
      .addItem('F3セルの施設データを取り込む', 'runImportFromButton')
      .addItem('マスタデータを最新化（同期）', 'syncMasterSheet')
      .addItem('キャッシュ（URL）をすべてクリア', 'clearUrlCache')
      .addToUi();
      
  // バックグラウンドで非表示の一時コピーシート（temp_import_）を一括削除（クリーンアップ）
  console.log("一時コピーシートのクリーンアップを開始します...");
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf("temp_import_") === 0) {
      try {
        ss.deleteSheet(sheets[i]);
        console.log("一時シートを削除しました: " + name);
      } catch (e) {
        console.error("一時シートのクリーンアップに失敗しました: " + name + " -> " + e.toString());
      }
    }
  }
}

// 保存されたURLキャッシュをクリアする関数
function clearUrlCache() {
  try {
    var properties = PropertiesService.getScriptProperties();
    var keys = properties.getKeys();
    var clearedCount = 0;
    
    for (var i = 0; i < keys.length; i++) {
      if (keys[i].indexOf("url_") === 0) {
        properties.deleteProperty(keys[i]);
        clearedCount++;
      }
    }
    SpreadsheetApp.getUi().alert("キャッシュをクリアしました。（クリア件数: " + clearedCount + "件）\n次回取り込み時に最新のマスタデータからURLを再取得します。");
  } catch (e) {
    SpreadsheetApp.getUi().alert("キャッシュのクリア中にエラーが発生しました:\n" + e.toString());
  }
}

// マスタデータを最新化（同期）するユーザー用関数
function syncMasterSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var success = syncMasterSheetInternal(ss);
    if (success) {
      SpreadsheetApp.getUi().alert("マスタデータの同期が完了しました。");
    } else {
      SpreadsheetApp.getUi().alert("マスタデータの同期に失敗しました。詳細はログをご確認ください。");
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("同期中にエラーが発生しました:\n" + e.toString());
  }
}

// 内部用同期関数
function syncMasterSheetInternal(ss) {
  try {
    console.log("外部からマスタデータの取得を開始します...");
    var masterSsUrl = "https://docs.google.com/spreadsheets/d/1PvsAhiLcpI8174QVLBelJFulMPxXvnSCDOJ8B4hZTlU/edit";
    var masterSs = SpreadsheetApp.openByUrl(masterSsUrl);
    var srcMasterSheet = masterSs.getSheetByName("各施設リンク先一覧 のコピー");
    if (!srcMasterSheet) {
      console.error("外部のマスタシートが見つかりませんでした。");
      return false;
    }
    
    var destMasterSheet = ss.getSheetByName("マスタ_施設リンク先");
    if (destMasterSheet) {
      ss.deleteSheet(destMasterSheet);
    }
    
    destMasterSheet = srcMasterSheet.copyTo(ss);
    destMasterSheet.setName("マスタ_施設リンク先");
    destMasterSheet.hideSheet(); // 非表示にする
    console.log("マスタの同期に成功しました。");
    return true;
  } catch (e) {
    console.error("マスタ同期エラー: " + e.toString());
    return false;
  }
}

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
    
    // 軽量ローディングモーダルを表示（画面センターに表示されます）
    var template = HtmlService.createTemplateFromFile('loading');
    template.facilityName = facilityName;
    
    var htmlOutput = template.evaluate()
        .setWidth(360)
        .setHeight(220);
        
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
  var success = false;
  
  try {
    // 1. 9行目から下のすべての行を丸ごと削除（値、書式、古い行グループを一発で消去）
    var lastRow = sheet.getLastRow();
    if (lastRow >= 9) {
      var maxRows = sheet.getMaxRows();
      var freezeRows = sheet.getFrozenRows();
      var rowsToDelete = lastRow - 8;
      
      // 削除後の残りの行数が固定行数（通常8行）以下になってしまうのを防ぐため、
      // 削除する行数と同等以上の空行をあらかじめ末尾に挿入して安全マージンを確保します。
      if (maxRows - rowsToDelete <= freezeRows) {
        sheet.insertRowsAfter(maxRows, rowsToDelete + 1);
      }
      
      sheet.deleteRows(9, rowsToDelete);
    }
    
    // 2. キャッシュ（PropertiesService）から対象施設のURLを取得（24時間有効）
    var properties = PropertiesService.getScriptProperties();
    var cacheKey = "url_" + facilityName.replace(/\s+/g, "_"); // 安全なキー名に変換
    var targetUrl = "";
    
    var cachedData = properties.getProperty(cacheKey); // 形式: "URL|タイムスタンプ"
    var isCacheValid = false;
    
    if (cachedData) {
      var parts = cachedData.split("|");
      var url = parts[0];
      var timestamp = parseInt(parts[1], 10);
      var now = new Date().getTime();
      
      // 24時間（86,400,000ミリ秒）以内か判定
      if (now - timestamp < 86400000) {
        targetUrl = url;
        isCacheValid = true;
        console.log("24時間以内の有効なキャッシュからURLを取得しました: " + targetUrl);
      }
    }
    
    if (!isCacheValid) {
      console.log("キャッシュが無効または期限切れのため、マスタシートから検索します...");
      var masterSheet = ss.getSheetByName("マスタ_施設リンク先");
      
      // ローカルシートが存在しない場合は、外部から同期する
      if (!masterSheet) {
        console.log("ローカルマスタシートが存在しないため、外部から同期します...");
        syncMasterSheetInternal(ss);
        masterSheet = ss.getSheetByName("マスタ_施設リンク先");
      }
      
      if (!masterSheet) {
        throw new Error("マスタシート「マスタ_施設リンク先」の作成に失敗しました。");
      }
      
      var masterLastRow = masterSheet.getLastRow();
      // A列〜D列の値と、D列のRichText（リンク情報）を一括取得
      var masterData = masterSheet.getRange(1, 1, masterLastRow, 4).getValues();
      var richTextValues = masterSheet.getRange(1, 4, masterLastRow, 1).getRichTextValues();
      
      // ループで施設名を探す
      for (var i = 0; i < masterData.length; i++) {
        if (masterData[i][1] === facilityName) {
          var richText = richTextValues[i][0];
          if (richText) {
            targetUrl = richText.getLinkUrl() || masterData[i][3];
          } else {
            targetUrl = masterData[i][3];
          }
          break;
        }
      }
      
      // もし施設名が見つからなかった場合、マスタデータが古い可能性があるため、もう一度だけ外部から最新化して再検索する
      if (!targetUrl || targetUrl === "") {
        console.log("施設名が見つからなかったため、マスタを強制同期して再検索します...");
        syncMasterSheetInternal(ss);
        masterSheet = ss.getSheetByName("マスタ_施設リンク先");
        masterLastRow = masterSheet.getLastRow();
        masterData = masterSheet.getRange(1, 1, masterLastRow, 4).getValues();
        richTextValues = masterSheet.getRange(1, 4, masterLastRow, 1).getRichTextValues();
        
        for (var i = 0; i < masterData.length; i++) {
          if (masterData[i][1] === facilityName) {
            var richText = richTextValues[i][0];
            if (richText) {
              targetUrl = richText.getLinkUrl() || masterData[i][3];
            } else {
              targetUrl = masterData[i][3];
            }
            break;
          }
        }
      }
      
      if (targetUrl && targetUrl !== "") {
        // 取得したURLとタイムスタンプを保存
        properties.setProperty(cacheKey, targetUrl + "|" + new Date().getTime());
        console.log("取得したURLをプロパティにキャッシュ保存しました: " + targetUrl);
      }
    }
    
    // URLが見つからない場合はエラー
    if (!targetUrl || targetUrl === "") {
      throw new Error("「" + facilityName + "」のURLが一覧シートから見つかりませんでした。");
    }
    
    // 3. データソースからデータを丸ごとコピー＆ペースト
    var srcSs = SpreadsheetApp.openByUrl(targetUrl);
    var srcSheet = srcSs.getSheetByName("シート1") || srcSs.getSheets()[0];
    
    // 一時コピー用の一意なシート名
    var tempSheetName = "temp_import_" + new Date().getTime();
    
    // 別スプレッドシートのシートを一時コピーし、即座に非表示にする
    tempSheet = srcSheet.copyTo(ss);
    tempSheet.setName(tempSheetName);
    tempSheet.hideSheet(); // 画面上見えないようにする
    
    var srcLastRow = tempSheet.getLastRow();
    
    // 一時シートの9行目から最終行までの「行全体」の範囲を取得
    var srcRowRange = tempSheet.getRange("9:" + srcLastRow);
    
    // 本来のシートの9行目の「行全体」を貼り付け先に指定
    var destRowRange = sheet.getRange("9:9");
    
    // 行全体のコピー＆ペースト（値・数式・書式・行高がコピーされます）
    srcRowRange.copyTo(destRowRange);
    
    // スプレッドシートの変更（行の自動追加など）を強制同期させ、Sheets APIに反映します
    SpreadsheetApp.flush();
    
    // Sheets API（v4）を使って、一時シートから本番シートに行グループを一括コピー
    console.log("Sheets APIを使用して行グループを一括コピーします...");
    var spreadsheetId = ss.getId();
    
    // コピー元（tempSheet）の行グループを取得
    var response = Sheets.Spreadsheets.get(spreadsheetId, {
      ranges: [tempSheet.getName()],
      fields: "sheets(properties(sheetId),rowGroups)"
    });
    
    var tempSheetRowGroups = response.sheets[0].rowGroups || [];
    var destSheetId = sheet.getSheetId();
    
    var requests = [];
    
    // 行グループがある場合、本番シートに適用するためのリクエストを作成
    if (tempSheetRowGroups.length > 0) {
      for (var i = 0; i < tempSheetRowGroups.length; i++) {
        var group = tempSheetRowGroups[i];
        
        // 9行目（index: 8）以上のグループのみを対象とする
        if (group.range.startIndex >= 8) {
          requests.push({
            addDimensionGroup: {
              range: {
                sheetId: destSheetId,
                dimension: "ROWS",
                startIndex: group.range.startIndex,
                endIndex: group.range.endIndex
              }
            }
          });
        }
      }
    }
    
    // 一括でリクエストを送信
    if (requests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({
        requests: requests
      }, spreadsheetId);
      console.log("行グループの一括適用が完了しました。適用件数: " + requests.length);
    }
    
    // 完了フラグをセット
    success = true;
    
    // 4. 完了トーストを表示
    ss.toast(facilityName + " のデータを反映しました。", "読込完了", 3);
    
  } finally {
    // ユーザーのロード待ち時間を短縮するため、正常完了時はここでは削除（deleteSheet）しません。
    // 代わりに非表示のまま残しておき、次回スプレッドシートを開いたとき（onOpen）に裏側で一括削除します。
    // コピー中に失敗（success === false）した場合のみ、即座に削除します。
    if (tempSheet && !success) {
      try {
        ss.deleteSheet(tempSheet);
      } catch (e) {
        console.error("エラー時の一時シート削除に失敗しました: " + e.toString());
      }
    }
  }
}