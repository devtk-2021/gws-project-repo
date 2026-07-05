// 対象シート名の定数
const TARGET_SHEET_NAME = "収支（実績）";
const TARGET_SHEET_NAME_SIM = "収支（シミュレーション）";

// スプレッドシートを開いたときに起動する関数
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var ui = SpreadsheetApp.getUi();

  // カスタムメニューを追加
  ui.createMenu('カスタム機能')
    .addItem('F3セルの施設データを取り込む（実績）', 'runImportFromButton')
    .addItem('シミュレーションインポート（フォーム）', 'showSimulationForm')
    .addItem('マスタデータを最新化（同期）', 'syncMasterSheet')
    .addItem('キャッシュ（URL）をすべてクリア', 'clearUrlCache')
    .addToUi();

  // バックグラウンドで非表示の一時コピーシート（temp_import_、temp_sim_、backup_sim_）を一括削除（クリーンアップ）
  console.log("一時コピーシートのクリーンアップを開始します...");
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf("temp_import_") === 0 || name.indexOf("temp_sim_") === 0 || name.indexOf("backup_sim_") === 0) {
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

// 実績取り込みボタンから起動される関数
function runImportFromButton() {
  runImportFromButtonCommon("実績");
}

// シミュレーション取り込みボタンから起動される関数
function runImportSimulationFromButton() {
  runImportFromButtonCommon("シュミレーション");
}

// ボタンから起動される共通関数
function runImportFromButtonCommon(type) {
  try {
    var sheetName = type === "実績" ? TARGET_SHEET_NAME : TARGET_SHEET_NAME_SIM;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      SpreadsheetApp.getUi().alert("「" + sheetName + "」シートが見つかりませんでした。");
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
    template.functionName = type === "実績" ? "runImport" : "runImportSimulation";

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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 引数が空の場合のフォールバック（ボタン直接実行など）
  if (!facilityName) {
    var activeSheet = ss.getActiveSheet();
    facilityName = activeSheet.getRange("F3").getValue();
    console.log("引数 facilityName が空のため、アクティブシートのF3から取得しました: " + facilityName);
  }
  
  if (!facilityName) {
    throw new Error("施設名が選択されていません。F3セルに施設名を入力してください。");
  }

  var sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) {
    throw new Error("「" + TARGET_SHEET_NAME + "」シートが見つかりませんでした。");
  }
  var tempSheet = null;
  var success = false;

  try {
    // A1:Z1の既存リンクをクリア
    sheet.getRange("A1:Z1").clearContent();

    // 0. 本番シートの既存の行グループ（9行目以降）を Sheets API で一括削除して全クリアする
    console.log("既存の行グループの一括削除を開始します...");
    var spreadsheetId = ss.getId();
    var destSheetId = sheet.getSheetId();

    var destSheetResponse = Sheets.Spreadsheets.get(spreadsheetId, {
      ranges: [sheet.getName()],
      fields: "sheets(rowGroups)"
    });

    var existingRowGroups = destSheetResponse.sheets[0].rowGroups || [];
    var deleteRequests = [];

    if (existingRowGroups.length > 0) {
      for (var i = 0; i < existingRowGroups.length; i++) {
        var group = existingRowGroups[i];
        // すべての行グループを例外なく全削除（全クリア）する
        deleteRequests.push({
          deleteDimensionGroup: {
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

    if (deleteRequests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({
        requests: deleteRequests
      }, spreadsheetId);
      console.log("既存の行グループを一括全削除しました。件数: " + deleteRequests.length);
    }

    // 1. 9行目から下のすべての行を丸ごと削除（値、書式、古い行グループを一発で消去）
    var maxRows = sheet.getMaxRows();
    if (maxRows >= 9) {
      var freezeRows = sheet.getFrozenRows();
      var rowsToDelete = maxRows - 8; // lastRowではなく、シートの最大行数（maxRows）まで全削除して完全にクリーンアップします

      // 削除後の残りの行数が固定行数（通常8行）以下になってしまうのを防ぐため、
      // 削除する行数と同等以上の空行をあらかじめ末尾に挿入して安全マージンを確保します。
      if (maxRows - rowsToDelete <= freezeRows) {
        sheet.insertRowsAfter(maxRows, rowsToDelete + 1);
      }

      sheet.deleteRows(9, rowsToDelete);
    }

    // 2. キャッシュ（PropertiesService）から対象施設のURLを取得
    var properties = PropertiesService.getScriptProperties();
    var cacheKey = "url_" + facilityName.replace(/\s+/g, "_"); // 安全なキー名に変換
    var targetUrl = "";

    var cachedData = properties.getProperty(cacheKey); // 形式: "URL|タイムスタンプ"

    if (cachedData) {
      var parts = cachedData.split("|");
      targetUrl = parts[0];
      console.log("キャッシュからURLを取得しました: " + targetUrl);
    }

    // キャッシュにない場合は、ローカルのマスタシートから直接検索する（日中は外部マスタの同期は行わない）
    if (!targetUrl || targetUrl === "") {
      console.log("キャッシュにないため、ローカルのマスタシートから検索します...");
      var masterSheet = ss.getSheetByName("マスタ_施設リンク先");

      if (masterSheet) {
        var masterLastRow = masterSheet.getLastRow();
        if (masterLastRow > 0) {
          var masterRange = masterSheet.getRange(1, 1, masterLastRow, 4);
          var masterValues = masterRange.getValues();
          var masterFormulas = masterRange.getFormulas();
          var masterRichTexts = masterRange.getRichTextValues();

          for (var i = 0; i < masterValues.length; i++) {
            if (masterValues[i][1] === facilityName) {
              targetUrl = extractUrlFromRow(masterValues, masterFormulas, masterRichTexts, i, 3);
              break;
            }
          }
        }
      }
    }

    // URLが見つからない場合はエラー
    if (!targetUrl || targetUrl === "") {
      throw new Error("「" + facilityName + "」のURLが見つかりませんでした。夜間の自動同期をお待ちいただくか、管理機能からマスタデータの手動更新を実行してください。");
    }

    // 3. データソースからデータを丸ごとコピー＆ペースト
    var srcSs = SpreadsheetApp.openByUrl(targetUrl);
    var srcSheet = srcSs.getSheetByName(TARGET_SHEET_NAME) || srcSs.getSheets()[0];

    // 【最重要】コピー元シート自体の最終行（値の存在する本当の最終行）をこの時点で正確に取得します
    // これにより、一時シートにコピーした後のタイムラグによる行数未同期バグを防ぎます
    var srcLastRow = srcSheet.getLastRow();
    console.log("コピー元シートのデータ最終行: " + srcLastRow);

    // 一時コピー用の一意なシート名
    var tempSheetName = "temp_import_" + new Date().getTime();

    // 別スプレッドシートのシートを一時コピーし、即座に非表示にする
    tempSheet = srcSheet.copyTo(ss);
    tempSheet.setName(tempSheetName);
    tempSheet.hideSheet(); // 画面上見えないようにする

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

    // 3. 今後の手動編集や拡張を容易にするため、シートの末尾に十分な空行（1000行まで）を確保します
    var currentMaxRows = sheet.getMaxRows();
    if (currentMaxRows < 1000) {
      var rowsToAdd = 1000 - currentMaxRows;
      sheet.insertRowsAfter(currentMaxRows, rowsToAdd);

      // 挿入された行が上の行の書式（左端の赤・オレンジの帯など）を引き継いでしまうのを防ぐため、
      // 追加された空行エリアの書式をすべてクリアして真っ白にします
      var addedRange = sheet.getRange(currentMaxRows + 1, 1, rowsToAdd, sheet.getMaxColumns());
      addedRange.clearFormat();
    }

    // 完了フラグをセット
    success = true;

    // 取得元のファイルをスマートチップ（ファイル）としてD1セルに書き出し
    try {
      var spreadsheetId = ss.getId();
      var sheetId = sheet.getSheetId();
      var updateChipRequest = [{
        updateCells: {
          range: {
            sheetId: sheetId,
            startRowIndex: 0, // 1行目 (index 0)
            endRowIndex: 1,
            startColumnIndex: 3, // D列 (index 3)
            endColumnIndex: 4
          },
          rows: [{
            values: [{
              userEnteredValue: { stringValue: "@" },
              chipRuns: [{
                startIndex: 0,
                chip: {
                  richLinkProperties: {
                    uri: targetUrl
                  }
                }
              }]
            }]
          }],
          fields: "userEnteredValue,chipRuns"
        }
      }];
      
      Sheets.Spreadsheets.batchUpdate({
        requests: updateChipRequest
      }, spreadsheetId);
      console.log("D1セルにスマートチップを挿入しました: " + targetUrl);
    } catch (linkError) {
      console.error("スマートチップ書き出しエラー: " + linkError.toString());
      // エラー発生時のフォールバックとして通常のハイパーリンクを設定
      try {
        var srcSsName = srcSs.getName();
        sheet.getRange("D1").setValue('=HYPERLINK("' + targetUrl + '", "' + srcSsName + '")');
      } catch (fallbackError) {
        console.error("フォールバックのハイパーリンク書き出しも失敗しました: " + fallbackError.toString());
      }
    }

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

/**
 * 与えられたマスタ行データからURLを抽出するヘルパー関数
 */
function extractUrlFromRow(values, formulas, richTexts, rowIndex, colIndex) {
  var url = "";

  // 1. RichTextからURLを取得
  var richText = richTexts[rowIndex][colIndex];
  if (richText) {
    url = richText.getLinkUrl();
    if (url) return url;
  }

  // 2. 数式（HYPERLINK）からURLを取得
  var formula = formulas[rowIndex][colIndex];
  if (formula) {
    var match = formula.match(/=HYPERLINK\(\s*["']([^"']+)["']/i);
    if (match && match[1]) {
      return match[1];
    }
  }

  // 3. セルの値（文字列）からURLを取得
  var value = values[rowIndex][colIndex];
  if (typeof value === 'string' && value.match(/^https?:\/\/[^\s]+/)) {
    return value;
  }

  return "";
}

/**
 * 夜間バッチ処理: マスタシートの同期と全URLの事前キャッシュ登録
 */
function dailyNightlyTask() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log("【夜間バッチ】処理を開始します。");

  // 1. 外部マスタから最新データを同期
  var syncSuccess = syncMasterSheetInternal(ss);
  if (!syncSuccess) {
    console.error("【夜間バッチ】マスタの同期に失敗しました。処理を中断します。");
    return;
  }

  var masterSheet = ss.getSheetByName("マスタ_施設リンク先");
  if (!masterSheet) {
    console.error("【夜間バッチ】マスタシート「マスタ_施設リンク先」が見つかりません。");
    return;
  }

  var masterLastRow = masterSheet.getLastRow();
  if (masterLastRow < 2) {
    console.log("【夜間バッチ】マスタデータが存在しません（ヘッダー行のみ、または空）。");
    return;
  }

  // 2. 全行のデータを一括取得してキャッシュ化
  var masterRange = masterSheet.getRange(1, 1, masterLastRow, 4);
  var masterValues = masterRange.getValues();
  var masterFormulas = masterRange.getFormulas();
  var masterRichTexts = masterRange.getRichTextValues();

  var properties = PropertiesService.getScriptProperties();
  var now = new Date().getTime();
  var cachedCount = 0;

  // 2行目（インデックス1）からループ
  for (var i = 1; i < masterValues.length; i++) {
    var facilityName = masterValues[i][1]; // B列: 施設名
    if (!facilityName) continue;

    var url = extractUrlFromRow(masterValues, masterFormulas, masterRichTexts, i, 3); // D列: URL
    if (url) {
      var cacheKey = "url_" + facilityName.replace(/\s+/g, "_");
      properties.setProperty(cacheKey, url + "|" + now);
      cachedCount++;
      console.log("キャッシュ登録: " + facilityName + " -> " + url);
    } else {
      console.warn("URL未設定の施設をスキップしました: " + facilityName);
    }
  }

  console.log("【夜間バッチ】キャッシュ登録完了。合計: " + cachedCount + " 件");
}

/**
 * 夜間バッチ実行用の時間主導型トリガーを登録する関数
 * ※管理者が手動で一度だけ実行すれば、以降は毎日午前2〜3時に自動実行されます。
 */
function setNightlyTrigger() {
  var functionName = "dailyNightlyTask";
  var triggers = ScriptApp.getProjectTriggers();

  // 重複登録防止のため既存のトリガーを削除
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
      console.log("既存の夜間バッチトリガーを削除しました。");
    }
  }

  // 毎日午前2時〜3時に実行
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  console.log("夜間バッチ（毎日午前2:00〜3:00）のトリガーを新規登録しました。");
}

/**
 * マスタシートのB列3行目から最終行までの施設名リストを取得する関数
 * @return {Array<string>} 施設名リスト
 */
function getFacilityList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName("マスタ_施設リンク先");
    if (!masterSheet) {
      console.error("マスタシート「マスタ_施設リンク先」が見つかりませんでした。");
      return [];
    }
    
    var lastRow = masterSheet.getLastRow();
    if (lastRow < 3) {
      console.warn("マスタデータが3行目未満です。");
      return [];
    }
    
    // B3からB列最終行までのデータを一括取得
    var values = masterSheet.getRange(3, 2, lastRow - 2, 1).getValues();
    var list = [];
    for (var i = 0; i < values.length; i++) {
      var val = values[i][0];
      if (val !== undefined && val !== null) {
        var strVal = String(val).trim();
        if (strVal !== "" && list.indexOf(strVal) === -1) {
          list.push(strVal);
        }
      }
    }
    
    // あいうえお順（辞書順）でソート
    list.sort();
    return list;
  } catch (e) {
    console.error("getFacilityListエラー: " + e.toString());
    return [];
  }
}

/**
 * シミュレーションデータ取り込み用入力フォームダイアログを表示する関数
 */
function showSimulationForm() {
  try {
    var template = HtmlService.createTemplateFromFile('inputForm');
    
    // 事前に施設リストを読み込んでHTMLテンプレートに埋め込む（初動の遅延防止）
    var facilityList = getFacilityList();
    template.facilityListJson = JSON.stringify(facilityList);
    
    var htmlOutput = template.evaluate()
        .setWidth(600)
        .setHeight(260) // サジェスト候補リストが表示されるため、少し高さを広げます
        .setTitle(' '); // Google側の標準ヘッダーを極力目立たせないため空白に
        
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
  } catch (e) {
    SpreadsheetApp.getUi().alert("入力フォーム起動中にエラーが発生しました:\n" + e.toString());
  }
}