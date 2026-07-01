function handleEdit(e) {
  try {
    console.log("handleEditがトリガーされました");
    if (!e) {
      console.log("引数 e が定義されていません（GASエディタから直接実行された可能性があります）");
      return;
    }

    // 1. 編集されたセルとシートのチェック
    var range = e.range;
    var sheet = range.getSheet();
    var sheetName = sheet.getName();
    var a1Notation = range.getA1Notation();

    console.log("編集されたシート: " + sheetName + ", セル: " + a1Notation);

    // 対象のシート名が「収支（実績）」で、セルが「F3」の場合のみ実行
    if (sheetName === "収支（実績）" && a1Notation === "F3") {
      var facilityName = range.getValue(); // 選択された施設名
      console.log("選択された施設名: " + facilityName);

      // 9行目から下の既存データを一旦すべてクリア（古い書式やデータを残さないため）
      var lastRow = sheet.getLastRow();
      console.log("現在のシートの最終行: " + lastRow);
      if (lastRow >= 9) {
        console.log("9行目以下をクリアします");
        sheet.getRange(9, 1, lastRow - 8, sheet.getLastColumn()).clear();
      }

      // プルダウンが「空」にされた場合はここで処理を終了
      if (!facilityName) {
        console.log("施設名が空のため処理を終了します");
        return;
      }

      // 2. URL一覧シートから対象施設のURLを検索
      var masterSsUrl = "https://docs.google.com/spreadsheets/d/1PvsAhiLcpI8174QVLBelJFulMPxXvnSCDOJ8B4hZTlU/edit";
      console.log("URL一覧シートを開きます: " + masterSsUrl);
      var masterSs = SpreadsheetApp.openByUrl(masterSsUrl);
      var masterSheet = masterSs.getSheetByName("各施設リンク先一覧 のコピー");

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

      console.log("取得したURL: " + targetUrl);

      // URLが見つからない場合はエラーを画面に出して終了
      if (!targetUrl || targetUrl === "") {
        throw new Error("「" + facilityName + "」のURLが一覧シートから見つかりませんでした。");
      }

      // 3. データソースからデータを丸ごとコピー＆ペースト
      console.log("データソースを開きます...");
      var srcSs = SpreadsheetApp.openByUrl(targetUrl);
      var srcSheet = srcSs.getSheetByName("シート1") || srcSs.getSheets()[0];

      // 別スプレッドシートのシートを現在のスプレッドシートに一時コピー（同じファイル内にする）
      var currentSs = SpreadsheetApp.getActiveSpreadsheet();
      console.log("データソースシートを一時コピーします...");
      var tempSheet = srcSheet.copyTo(currentSs);

      try {
        // コピー元の範囲（A1:Pの最終行まで）を取得
        var srcLastRow = tempSheet.getLastRow();
        console.log("一時コピーシートの最終行: " + srcLastRow);
        var srcRange = tempSheet.getRange("A1:P" + srcLastRow);

        // 出力先シートの9行目、1列目（A9セルの位置）を貼り付けの起点にする
        var destRange = sheet.getRange(9, 1);

        // コピー元の数式・書式・値をすべて丸ごと上書きコピー
        console.log("データをコピーします...");
        srcRange.copyTo(destRange);

        // 画面の右下に一時的に完了ポップアップを表示
        currentSs.toast(facilityName + " のデータを反映しました。", "読込完了", 3);
        console.log("処理が正常に完了しました");

      } finally {
        // 例外が発生した場合でも一時シートは確実に削除する
        if (tempSheet) {
          console.log("一時コピーシートを削除します");
          currentSs.deleteSheet(tempSheet);
        }
      }
    }
  } catch (error) {
    console.error("エラーが発生しました: " + error.toString());
    SpreadsheetApp.getUi().alert("エラーが発生しました:\n" + error.toString());
  }
}