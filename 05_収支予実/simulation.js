/**
 * シミュレーションインポート フェーズ1 (データ準備、反映先へのコピー完了まで)
 * クライアント側から呼ばれ、データがスプレッドシートの画面に書き込まれた（値貼り付けされた）瞬間に即座に応答します。
 * 
 * @param {string} facilityName 対象施設名
 * @return {Object} フェーズ2で必要なパラメータ
 */
function runImportSimulationPhase1(facilityName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 引数が空の場合のフォールバック（ボタン直接実行など）
  if (!facilityName) {
    var activeSheet = ss.getActiveSheet();
    facilityName = activeSheet.getRange("F3").getValue();
  }
  
  console.log("【シミュレーション Phase1】インポート開始: facilityName = " + facilityName);
  
  if (!facilityName) {
    throw new Error("施設名が選択されていません。F3セルに施設名を入力してください。");
  }
  
  var tempSheet = null;
  
  try {
    // 1. キャッシュ・ローカルマスタから対象施設のURLを取得
    var properties = PropertiesService.getScriptProperties();
    var cacheKey = "url_" + facilityName.replace(/\s+/g, "_");
    var targetUrl = "";
    
    var cachedData = properties.getProperty(cacheKey);
    if (cachedData) {
      var parts = cachedData.split("|");
      targetUrl = parts[0];
      console.log("キャッシュからURLを取得しました: " + targetUrl);
    }
    
    // キャッシュにない場合はマスタから検索
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
    
    if (!targetUrl || targetUrl === "") {
      throw new Error("「" + facilityName + "」のURLが見つかりませんでした。");
    }
    
    // 2. データソース（コピー元）から「収支（シミュレーション）」シートを取得
    var srcSs = SpreadsheetApp.openByUrl(targetUrl);
    var srcSheet = srcSs.getSheetByName(TARGET_SHEET_NAME_SIM) || srcSs.getSheets()[0];
    var srcSsName = srcSs.getName();
    
    // 3. コピー元シートを反映先スプレッドシートへ一時的にコピー（値化と書式調整用）
    var tempSheetName = "temp_sim_" + new Date().getTime();
    tempSheet = srcSheet.copyTo(ss);
    tempSheet.setName(tempSheetName);
    tempSheet.hideSheet(); // 画面上は見えないようにしておく
    
    // 4. 新しくコピーしたシートのすべての数式を「値」に変換（値貼り付けで上書き）
    console.log("数式を値（実数）に変換しています...");
    var dataRange = tempSheet.getDataRange();
    dataRange.copyTo(dataRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    // 5. 既存の「収支（シミュレーション）」シートを取得（存在しなければ新規作成）
    var destSheet = ss.getSheetByName(TARGET_SHEET_NAME_SIM);
    if (!destSheet) {
      destSheet = ss.insertSheet(TARGET_SHEET_NAME_SIM);
    }
    
    // 6. 既存シートの中身と行グループを完全にクリア（シートオブジェクト自体は削除しない）
    console.log("既存シートの中身をクリアしています...");
    destSheet.clear(); // 値、書式、入力規則等を全てクリア
    
    // 既存シートの行グループを一括削除（Sheets API使用）
    var spreadsheetId = ss.getId();
    var destSheetId = destSheet.getSheetId();
    var destSheetResponse = Sheets.Spreadsheets.get(spreadsheetId, {
      ranges: [destSheet.getName()],
      fields: "sheets(rowGroups)"
    });
    var existingRowGroups = destSheetResponse.sheets[0].rowGroups || [];
    var deleteRequests = [];
    if (existingRowGroups.length > 0) {
      for (var i = 0; i < existingRowGroups.length; i++) {
        var group = existingRowGroups[i];
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
      console.log("既存の行グループを削除しました。");
    }
    
    // 7. 反映先シートの行数・列数をデータがある範囲に合わせて必要に応じて拡張（削除処理は著しく遅いため廃止）
    var srcLastRow = tempSheet.getLastRow();
    var srcLastCol = tempSheet.getLastColumn();
    
    // 行数が足りない場合のみ追加
    var destMaxRows = destSheet.getMaxRows();
    if (destMaxRows < srcLastRow) {
      destSheet.insertRowsAfter(destMaxRows, srcLastRow - destMaxRows);
      console.log("反映先シートに行数を追加しました: " + (srcLastRow - destMaxRows) + "行");
    }
    
    // 列数が足りない場合のみ追加
    var destMaxCols = destSheet.getMaxColumns();
    if (destMaxCols < srcLastCol) {
      destSheet.insertColumnsAfter(destMaxCols, srcLastCol - destMaxCols);
      console.log("反映先シートに列数を追加しました: " + (srcLastCol - destMaxCols) + "列");
    }
    
    // 8. 一時コピーシートから既存シートへデータを上書きコピペ（反映先への値貼り付け）
    console.log("データを反映先シートへ上書きコピーしています...");
    var srcRowRange = tempSheet.getRange("1:" + srcLastRow);
    var destRowRange = destSheet.getRange("1:1");
    srcRowRange.copyTo(destRowRange); // ★この瞬間に、データ（値）が反映先シートの画面にパッと表示されます！
    
    // 反映先のシートを前面にアクティブ化
    destSheet.activate();
    
    console.log("【シミュレーション Phase1】反映先への値貼り付けが完了しました。ダイアログを閉じます。");
    
    // フェーズ2（遅い列幅・行グループ同期、シート削除、スマートチップ、トースト表示）のためのパラメータを返却
    return {
      tempSheetName: tempSheetName,
      targetUrl: targetUrl,
      facilityName: facilityName,
      srcSsName: srcSsName,
      destSheetId: destSheetId,
      srcLastCol: srcLastCol
    };
    
  } catch (error) {
    console.error("シミュレーションインポートPhase1中にエラーが発生しました: " + error.toString());
    if (tempSheet) {
      try {
        ss.deleteSheet(tempSheet);
      } catch (e) {
        console.error("一時シートの削除失敗: " + e.toString());
      }
    }
    throw error;
  }
}

/**
 * シミュレーションインポート フェーズ2 (列幅同期、行グループ適用、一時シート削除、スマートチップ、トースト表示)
 * クライアントのモーダルが閉じた後、バックグラウンドで非同期に実行されます。
 * 
 * @param {Object} params Phase1から返されたパラメータ
 */
function runImportSimulationPhase2(params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tempSheet = ss.getSheetByName(params.tempSheetName);
    var destSheet = ss.getSheetByName(TARGET_SHEET_NAME_SIM);
    var spreadsheetId = ss.getId();
    
    console.log("【シミュレーション Phase2】バックグラウンド反映（レイアウト・後処理）を開始します...");
    
    // 1. 列幅の同期 (ループ処理で少し時間がかかるため、バックグラウンドに回す)
    if (tempSheet && destSheet) {
      console.log("列幅を同期しています...");
      for (var col = 1; col <= params.srcLastCol; col++) {
        destSheet.setColumnWidth(col, tempSheet.getColumnWidth(col));
      }
    }
    
    // 2. 一時コピーシート（tempSheet）から既存シート（destSheet）に行グループを一括コピー（Sheets API）
    if (tempSheet) {
      console.log("行グループを適用しています...");
      var response = Sheets.Spreadsheets.get(spreadsheetId, {
        ranges: [tempSheet.getName()],
        fields: "sheets(properties(sheetId),rowGroups)"
      });
      var tempSheetRowGroups = response.sheets[0].rowGroups || [];
      var requests = [];
      if (tempSheetRowGroups.length > 0) {
        for (var i = 0; i < tempSheetRowGroups.length; i++) {
          var group = tempSheetRowGroups[i];
          requests.push({
            addDimensionGroup: {
              range: {
                sheetId: params.destSheetId,
                dimension: "ROWS",
                startIndex: group.range.startIndex,
                endIndex: group.range.endIndex
              }
            }
          });
        }
      }
      if (requests.length > 0) {
        Sheets.Spreadsheets.batchUpdate({
          requests: requests
        }, spreadsheetId);
        console.log("行グループの適用が完了しました。");
      }
    }
    
    // 3. 一時コピーシートの削除 (同期ブロックが著しく重いため、バックグラウンドで処理します)
    if (tempSheet) {
      try {
        ss.deleteSheet(tempSheet);
        console.log("一時コピーシートを削除しました。");
      } catch (e) {
        console.error("一時コピーシートの削除に失敗しました: " + e.toString());
      }
    }
    
    // 4. 取得元のファイルをスマートチップ（ファイル）としてJ1セルに書き出し
    try {
      var updateChipRequest = [{
        updateCells: {
          range: {
            sheetId: params.destSheetId,
            startRowIndex: 0,
            endRowIndex: 1,
            startColumnIndex: 9, // J列 (index 9)
            endColumnIndex: 10
          },
          rows: [{
            values: [{
              userEnteredValue: { stringValue: "@" },
              chipRuns: [{
                startIndex: 0,
                chip: {
                  richLinkProperties: {
                    uri: params.targetUrl
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
      console.log("J1セルにスマートチップを挿入しました: " + params.targetUrl);
    } catch (linkError) {
      console.error("スマートチップ書き出しエラー: " + linkError.toString());
      try {
        if (destSheet) {
          destSheet.getRange("J1").setValue('=HYPERLINK("' + params.targetUrl + '", "' + params.srcSsName + '")');
        }
      } catch (fallbackError) {
        console.error("フォールバックのハイパーリンク書き出しも失敗しました: " + fallbackError.toString());
      }
    }
    
    // 5. 完全に後処理が終わったタイミングで、トースト通知を表示
    ss.toast(params.facilityName + " のシミュレーションデータ反映を完了しました。", "反映完了", 3);
    console.log("【シミュレーション Phase2】すべてのバックグラウンド処理が完了しました。");
    
  } catch (error) {
    console.error("【シミュレーション Phase2】実行中にエラーが発生しました: " + error.toString());
  }
}

/**
 * 互換用：シミュレーションデータ取り込み（同期実行）
 * 物理ボタンなどから引数なしで直接起動された場合でも、従来通り動作するようにサポートします。
 * 
 * @param {string} facilityName 対象施設名
 */
function runImportSimulation(facilityName) {
  var result = runImportSimulationPhase1(facilityName);
  runImportSimulationPhase2(result);
}
