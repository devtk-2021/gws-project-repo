function autoImportCSV() {
  // =========================================================================
  // 【設定エリア】環境に合わせて、各種IDとシート名を書き換えてください。
  // =========================================================================
  const FOLDER_INPUT_ID = '1rRufbpsviqtbo9XMyJxH6HnpsNTQlA9u'; // 01_未処理CSVフォルダ
  const FOLDER_OUTPUT_ID = '1Hu-iX-SBVqiIH8pfYz2XAKyE8b-JoFJA'; // 02_処理済みCSVフォルダ
  const FOLDER_LOG_ID = '1pj_xMTOr0gPJvOuhQfsoCXrj5MOJrvqh'; // log出力フォルダ
  const SPREADSHEET_ID = '1vR6Ri-q2RgKIpu5B2UcwovjM3Ozqo_woJg-X1jrLWC8'; // 出力先スプレッドシート
  const SHEET_NAME = 'シート1'; // 転記先のシート名
  // ========================================================

// -------------------------------------------------------------------------
  // 1. ログ管理用の準備
  // -------------------------------------------------------------------------
  // 今回の実行中に発生したログメッセージ（正常・異常ともに）を一時的に保存しておくための配列
  let logMessages = [];

  /**
   * ログを記録するための専用サブ関数
   * @param {string} message - ログに残す文章
   * @param {string} type - ログの種類（'INFO' または 'ERROR'）
   */
  function addLog(message, type = 'INFO') {
    // 現在の時刻を「yyyy/MM/dd HH:mm:ss」の形式で取得
    let timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    // ログの1行を生成（例：[2026/06/14 10:00:00] [INFO] 処理を開始します）
    let logLine = `[${timestamp}] [${type}] ${message}`;
    
    console.log(logLine);      // GASの「実行ログ」画面（黒い画面）にリアルタイム表示する
    logMessages.push(logLine); // 後でテキストファイルに保存するために、配列の末尾に追加する
  }


  // -------------------------------------------------------------------------
  // 2. システムの初期化と接続チェック（エラーハンドリング）
  // -------------------------------------------------------------------------
  addLog("CSV自動転記処理を開始します。");

  let inputFolder, outputFolder, logFolder, ss, sheet;
  
  try {
    // 各フォルダやスプレッドシートのIDをシステムに認識させる
    inputFolder  = DriveApp.getFolderById(FOLDER_INPUT_ID);
    outputFolder = DriveApp.getFolderById(FOLDER_OUTPUT_ID);
    logFolder    = DriveApp.getFolderById(FOLDER_LOG_ID); 
    ss           = SpreadsheetApp.openById(SPREADSHEET_ID);
    sheet        = ss.getSheetByName(SHEET_NAME);
    
    // 指定したシート名が間違っている等で、シートが見つからなかった場合の安全対策
    if (!sheet) {
      throw new Error(`指定されたシート名「${SHEET_NAME}」がスプレッドシート内に見つかりません。`);
    }
  } catch (initError) {
    // 設定エラーが起きた場合は、原因をログに記録して処理を強制終了する（これ以上進むとエラーになるため）
    addLog(`初期設定に失敗しました。IDやシート名を確認してください。エラー内容: ${initError.message}`, 'ERROR');
    saveLogToFile(logFolder, logMessages); // そこまでのエラーログだけをファイルに書き出して終了
    return; 
  }


  // -------------------------------------------------------------------------
  // 3. 未処理フォルダ内のCSVファイルの存在チェック
  // -------------------------------------------------------------------------
  // 「01_未処理CSV」フォルダ内にあるファイルの一覧を取得
  const files = inputFolder.getFiles();

  // ファイルが1つも無い場合は、何もせず安全に終了する
  if (!files.hasNext()) {
    addLog("処理するCSVファイルはありません。処理を終了します。");
    saveLogToFile(logFolder, logMessages); // 終了ログを保存
    return;
  }


  // -------------------------------------------------------------------------
  // 4. ファイルのループ処理（フォルダ内のファイルを1つずつ順番に処理）
  // -------------------------------------------------------------------------
  while (files.hasNext()) {
    let file = files.next();
    let fileName = file.getName();
    
    // 拡張子が「.csv」である、またはMIMEタイプがCSVであるファイルだけを処理対象にする（誤作動防止）
    if (file.getMimeType() === MimeType.CSV || fileName.endsWith('.csv')) {
      addLog(`「${fileName}」の解析を始めます。`);
      
      // --- 【ファイル単位のエラーハンドリング】 ---
      try {
        // CSVファイルの中身を「UTF-8」の文字コードでテキストとして読み込む
        // ※もし転記された文字がバグる（文字化けする）場合は、下の 'UTF-8' を 'Shift_JIS' に書き換えてください。
        let data = file.getBlob().getDataAsString('UTF-8'); 
        
        // GASの標準機能を使って、テキストを2次元配列（行と列の格子状データ）に変換
        let csvData = Utilities.parseCsv(data);
        addLog(`CSVデータの解析に成功しました（総行数: ${csvData.length}行）。転記を開始します。`);
        
        let successCount = 0; // 転記に成功した行数を数えるカウンター
        let skipCount    = 0; // スキップ（空行や対象外データ）した行数を数えるカウンター

        // --- 【行単位のループ処理】 ---
        // 「i = 1」から始めることで、CSVの1行目（ヘッダー行：取引日,施設名...）を無視し、2行目のデータから処理する
        for (let i = 1; i < csvData.length; i++) {
          let row = csvData[i]; // 現在処理している1行分のデータ
          
          // 列数が足りない、または1列目（日付など）が空っぽの「不完全な空行」があれば処理を飛ばす
          if (row.length < 8 || row[0] === "") {
            skipCount++;
            continue;
          }

          // --- 【行ごとのデータエラーハンドリング】 ---
          try {
            // CSVの各列のデータを分かりやすい名前の変数に格納（※プログラムの世界は0番目から数えます）
            let facilityName = row[2];        // 3列目：施設名
            let jobName      = row[4];        // 5列目：職種名
            let type         = row[5];        // 6列目：収支区分（売上 / 支出）
            let amount       = Number(row[7]); // 8列目：金額（計算や比較ができるよう文字列から「数値」に変換）

            // 万が一、金額の列に「未定」などの文字が混じっていて数値化できなかった場合の対策
            if (isNaN(amount)) {
              throw new Error(`金額列（${row[7]}）を数値に変換できません。データ形式を確認してください。`);
            }

            // --- 【転記先のマッピングルール設定】 ---
            let targetRange = null; // 転記先セル（例: 'B5'）を入れるための空の変数

            // 条件分岐（if / else if）：CSVのデータ内容に応じて、書き込むセルのアドレスを決定する
            if (facilityName === "施設A（さくら）" && jobName === "介護職" && type === "売上") {
              targetRange = 'B5';
            } else if (facilityName === "施設A（さくら）" && jobName === "看護職" && type === "支出") {
              targetRange = 'B6';
            } else if (facilityName === "施設A（さくら）" && jobName === "共通" && type === "支出") {
              targetRange = 'B7';
            } else if (facilityName === "施設B（ひまわり）" && jobName === "介護職" && type === "売上") {
              targetRange = 'C5';
            } else if (facilityName === "施設B（ひまわり）" && jobName === "事務職" && type === "支出") {
              targetRange = 'C6';
            }

            // 上記のルールに合致するセルが見つかった場合のみ、スプレッドシートに書き込む
            if (targetRange) {
              sheet.getRange(targetRange).setValue(amount); // 指定セルに数値を上書き（※足し算したい場合は要修正）
              successCount++;
            } else {
              // どのルールにも該当しないデータ（新しい施設や未定義の職種など）は安全のためにスキップ
              skipCount++;
            }

          } catch (rowError) {
            // 【行エラー時の対応】
            // データの不備などでエラーが出ても、システムを全停止させず、エラー行だけスキップして次の行の処理を続ける
            // 何行目（CSVの行番号＝i+1）で何のエラーが起きたかをログに明記する
            addLog(`ファイル「${fileName}」の ${i + 1}行目でエラーが発生したためスキップしました。内容: ${rowError.message}`, 'ERROR');
          }
        } // --- 【行単位のループ処理 終了】 ---
        
        addLog(`「${fileName}」のデータ処理が完了（転記成功: ${successCount}件 / スキップ・対象外: ${skipCount}件）`);

        // --- 処理済みファイルの移動処理 ---
        // 次回の実行時に同じファイルを2重に処理（書き込み）してしまうのを防ぐため、ファイルを移動させる
        outputFolder.addFile(file);   // 「02_処理済みCSV」フォルダにファイルをリンクさせる
        inputFolder.removeFile(file); // 「01_未処理CSV」フォルダからリンクを外す（これで移動が完了する）
        addLog(`ファイル「${fileName}」を【02_処理済みCSV】フォルダに移動しました。`);

      } catch (fileError) {
        // 【ファイルエラー時の対応】
        // CSVファイル自体が破損している、または文字コード起因で読み込めなかった場合、
        // このファイルの処理を中断して、ログを残し、次のCSVファイルの処理へ進む（連鎖停止の防止）
        addLog(`「${fileName}」は重大な問題があるため処理を中断し、次のファイルに移ります。内容: ${fileError.message}`, 'ERROR');
      }
    } else {
      // CSV以外のファイル（PDFや画像など）が誤って置かれていた場合は、何もせず無視する
      addLog(`ファイル「${fileName}」はCSVではないためスキップしました。`);
    }
  } // --- 【ファイルのループ処理 終了】 ---
  
  addLog("すべてのCSV自動転記処理が終了しました。");
  
  // -------------------------------------------------------------------------
  // 5. ログのファイル保存処理を実行
  // -------------------------------------------------------------------------
  // 溜まった logMessages 配列を、指定されたフォルダ内のテキストファイルに書き出す
  saveLogToFile(logFolder, logMessages);
}


/**
 * 発生したログを、1日1ファイルのテキストファイルとして先頭に追記（降順）保存するサブ関数
 * @param {Folder} logFolder - ログを保存するGoogleドライブのフォルダオブジェクト
 * @param {Array} logMessages - 記録するログ文字列が格納された配列
 */
function saveLogToFile(logFolder, logMessages) {
  // 保存するログが無い、またはログフォルダが正しく読み込めていない場合は処理をしない
  if (logMessages.length === 0 || !logFolder) return;

  // 今日の日付を「yyyyMMdd（例: 20260614）」の形式の文字列にする
  let dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  // ファイル名を決定（例: 20260614_実行ログ.txt）
  let logFileName = `${dateStr}_実行ログ.txt`;
  
  // 今回の実行で蓄積されたログの配列（時系列順 [開始 -> 処理 -> 終了]）を、改行コードで1つの長い文章に合体させる
  let newLogText = logMessages.join('\n');
  
  // ログフォルダ内に、今日の日付のログファイルが既に存在するか検索する
  let existingFiles = logFolder.getFilesByName(logFileName);
  let logFile = null;
  
  // 【ゴミ箱対策】
  // 過去に削除して「ゴミ箱」に入っている同名ファイルをシステムが誤認しないよう、
  // ゴミ箱に入っていない（有効な）ファイルだけをループで探す
  while (existingFiles.hasNext()) {
    let file = existingFiles.next();
    if (!file.isTrashed()) {
      logFile = file; // 有効なファイルが見つかったら変数にキープしてループを抜ける
      break; 
    }
  }
  
  // 【条件分岐】すでに今日のログファイルが存在する場合（2回目以降の実行）
  if (logFile) {
    let currentContent = "";
    try {
      // 既にファイルに書き込まれている「過去のログ（古いデータ）」を一度すべて読み込む
      currentContent = logFile.getBlob().getDataAsString('UTF-8');
    } catch (e) {
      console.error("既存ログの読み込みに失敗したため新規上書きします:", e);
    }
    
    // 過去のログデータが正常に読み込めた場合
    if (currentContent) {
      // 最新のログ（newLogText）を「一番上（先頭）」に配置し、
      // 間に2つの改行（\n\n）を挟んでから、古いログ（currentContent）を下に連結してファイルの中身を更新する
      // これにより、「実行ブロック単位での降順」と「実行ブロック内での時系列（昇順）」が両立する
      logFile.setContent(newLogText + '\n\n' + currentContent);
    } else {
      // ファイルが空だった場合は、今回のログだけを書き込む
      logFile.setContent(newLogText);
    }
  } 
  // 【条件分岐】まだ今日のログファイルが存在しない場合（その日の一番最初の実行）
  else {
    // ログフォルダ内に、プレーンテキスト形式（MimeType.PLAIN_TEXT）で新しくファイルを作成する
    logFolder.createFile(logFileName, newLogText, MimeType.PLAIN_TEXT);
  }
}