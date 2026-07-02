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
  var success = false;
  
  try {
    // 1. 9行目から下のすべてのデータが存在する列をクリア
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow >= 9) {
      sheet.getRange(9, 1, lastRow - 8, lastCol > 0 ? lastCol : 1).clear();
      
      // 既存の古い行グループをすべて解除する（最大8階層分をフラット化）
      var clearRange = sheet.getRange(9, 1, lastRow - 8, 1);
      for (var d = 0; d < 8; d++) {
        try {
          clearRange.shiftRowGroupDepth(-1);
        } catch (e) {
          // これ以上解除できるグループが無い場合はループを抜ける
          break;
        }
      }
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
    var srcLastCol = tempSheet.getLastColumn();
    
    // A列の9行目から、データが存在する最終行・最終列までを取得
    var srcRange = tempSheet.getRange(9, 1, Math.max(1, srcLastRow - 8), Math.max(1, srcLastCol));
    
    // 出力先シートの9行目、1列目（A9セルの位置）を貼り付けの起点にする
    var destRange = sheet.getRange(9, 1);
    
    // コピー元の数式・書式・値をすべて丸ごと上書きコピー
    srcRange.copyTo(destRange);
    
    // 4. 行のグループ化情報をコピー元（tempSheet）から取得してコピー先に再現
    console.log("行グループの再現処理を開始します...");
    var maxDepth = 0;
    var depths = [];
    var useSheetsApi = (typeof Sheets !== 'undefined');
    
    if (useSheetsApi) {
      try {
        console.log("Sheets APIを使用して行グループ情報を一括取得します...");
        var spreadsheetId = ss.getId();
        var sheetName = tempSheet.getName();
        // Sheets APIでtempSheetの9行目〜最終行の行メタデータ（depth含む）を一括取得
        var response = Sheets.Spreadsheets.get(spreadsheetId, {
          ranges: [sheetName + "!A9:A" + srcLastRow],
          fields: "sheets(data(rowMetadata(depth)))"
        });
        
        var rowMetadata = response.sheets[0].data[0].rowMetadata;
        if (rowMetadata) {
          for (var i = 0; i < rowMetadata.length; i++) {
            var d = rowMetadata[i].depth || 0;
            depths.push(d);
            if (d > maxDepth) maxDepth = d;
          }
        }
      } catch (apiError) {
        console.error("Sheets APIの実行中にエラーが発生したため、従来方式にフォールバックします: " + apiError.toString());
        useSheetsApi = false;
      }
    }
    
    if (!useSheetsApi) {
      console.warn("【注意】Google Sheets APIが有効化されていないため、低速な従来方式（1行ずつのスキャン）を実行します。さらに高速化するには、GASの「サービス」から「Google Sheets API」を追加してください。");
      for (var r = 9; r <= srcLastRow; r++) {
        var d = tempSheet.getRowGroupDepth(r);
        depths.push(d);
        if (d > maxDepth) maxDepth = d;
      }
    }
    
    // 深度1から最大深度まで順番に、連続するグループを検出して適用
    for (var d = 1; d <= maxDepth; d++) {
      var startIdx = -1;
      for (var i = 0; i < depths.length; i++) {
        if (depths[i] >= d) {
          if (startIdx === -1) {
            startIdx = i; // グループの始まり
          }
        } else {
          if (startIdx !== -1) {
            // グループの終了、コピー先に一括適用
            var startRow = startIdx + 9;
            var endRow = i - 1 + 9;
            sheet.getRange(startRow + ":" + endRow).shiftRowGroupDepth(1);
            startIdx = -1;
          }
        }
      }
      // ループ終了時にグループが開いたままの場合の適用
      if (startIdx !== -1) {
        var startRow = startIdx + 9;
        var endRow = depths.length - 1 + 9;
        sheet.getRange(startRow + ":" + endRow).shiftRowGroupDepth(1);
      }
    }
    console.log("行グループの再現処理が完了しました。最大深度: " + maxDepth);
    
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