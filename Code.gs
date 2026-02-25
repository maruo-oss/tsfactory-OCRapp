/**
 * AI-OCR 自動発注管理システム（納品先対応版）
 */

const PROPS = PropertiesService.getScriptProperties();
const SHEET_NAME = 'OrderData';
const PROMPT_SHEET_NAME = 'プロンプト';

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('AI-OCR 発注管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function logFromClient(msg) {
  console.log(msg);
}

function runManualProcess() {
  const count = processOrders(); 
  return `処理完了: ${count}件`;
}

// --- メイン処理 ---
function processOrders() {
  console.log('[開始] 全体処理');
  let totalProcessedCount = 0;
  try {
    const rootInId = PROPS.getProperty('ROOT_IN_FOLDER_ID');
    const rootProcessedId = PROPS.getProperty('ROOT_PROCESSED_FOLDER_ID');
    const ssId = PROPS.getProperty('SPREADSHEET_ID');
    const apiKey = PROPS.getProperty('GEMINI_API_KEY');

    if (!rootInId || !rootProcessedId || !ssId || !apiKey) throw new Error('スクリプトプロパティ未設定');

    const geminiPrompt = getPromptFromSheet(ssId);
    const rootInFolder = DriveApp.getFolderById(rootInId);
    const rootProcessedFolder = DriveApp.getFolderById(rootProcessedId);
    const branchFolders = rootInFolder.getFolders();
    
    while (branchFolders.hasNext()) {
      const branchFolder = branchFolders.next(); 
      const branchName = branchFolder.getName();
      const targetBranchFolder = getOrCreateTargetFolder(rootProcessedFolder, branchName);
      const files = branchFolder.getFilesByType(MimeType.PDF);
      
      while (files.hasNext()) {
        processSingleFile(files.next(), branchName, targetBranchFolder, ssId, apiKey, geminiPrompt);
        totalProcessedCount++;
      }
    }
  } catch (e) {
    console.error(`全体エラー: ${e.message}`);
    throw e;
  }
  return totalProcessedCount;
}

// --- PDFの総ページ数を取得（Gemini Flash使用） ---
function getPageCount(file, apiKey) {
  const MODEL_NAME = 'gemini-2.5-flash';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${apiKey}`;
  const blob = file.getBlob();
  const base64Data = Utilities.base64Encode(blob.getBytes());
  const payload = {
    contents: [{ parts: [
      { text: 'このPDFの総ページ数を数字のみで回答してください。例: 3' },
      { inline_data: { mime_type: blob.getContentType(), data: base64Data } }
    ] }]
  };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200) throw new Error(json.error.message);
    const text = json.candidates[0].content.parts[0].text.trim();
    const pageCount = parseInt(text, 10);
    console.log(`[ページ数取得] ${file.getName()}: ${pageCount}ページ`);
    return isNaN(pageCount) || pageCount < 1 ? 1 : pageCount;
  } catch (e) {
    console.warn(`[ページ数取得失敗] ${file.getName()}: ${e.message} → 1ページとして処理`);
    return 1;
  }
}

// --- 個別ファイル処理（ページごとにAPI呼び出し） ---
function processSingleFile(file, branchName, targetBranchFolder, ssId, apiKey, promptText) {
  const fileName = file.getName();
  try {
    const totalPages = getPageCount(file, apiKey);
    const allRecords = [];
    let globalItemOrder = 1;

    for (let page = 1; page <= totalPages; page++) {
      console.log(`[OCR] ${fileName} - ページ ${page}/${totalPages} 処理中...`);

      // ページ指定のプロンプトを生成
      const pagePrompt = totalPages === 1
        ? promptText
        : `【重要】このPDFは${totalPages}ページありますが、${page}ページ目の内容のみを対象に読み取ってください。他のページの内容は絶対に含めないでください。\n\n${promptText}`;

      const extractedData = callGeminiApi(file, apiKey, pagePrompt);
      const items = (extractedData.items && extractedData.items.length > 0)
                    ? extractedData.items
                    : (page === 1 ? [{ product_code: "", product_name: "(明細なし)", quantity: 0, unit_price: 0 }] : []);

      const records = items.map(item => ({
        file_id: file.getId(),
        branch_name: branchName,
        file_name: fileName,
        status: '未処理',
        order_date: extractedData.order_date || '',
        order_number: extractedData.order_number || '',
        maker_name: extractedData.maker_name || '',
        shop_name: extractedData.shop_name || '',
        delivery_destination: extractedData.delivery_destination || '',
        product_code: item.product_code || '',
        product_name: item.product_name || '',
        quantity: safeParseFloat(item.quantity),
        unit_price: safeParseFloat(item.unit_price),
        processed_at: new Date(),
        item_order: globalItemOrder++,
        page_number: page  // プログラム的にページ番号を付与
      }));

      records.forEach(r => r.line_total = r.quantity * r.unit_price);
      allRecords.push(...records);
    }

    if (allRecords.length > 0) {
      saveToSpreadsheet(ssId, allRecords);
    }
    file.moveTo(targetBranchFolder);
    console.log(`[完了] ${fileName}: ${allRecords.length}件（${totalPages}ページ）`);
  } catch (e) {
    console.error(`エラー ${fileName}: ${e.message}`);
    file.setName(`【ERROR】${fileName}`);
  }
}

function safeParseFloat(val) {
  if (!val) return 0;
  const num = Number(String(val).replace(/[^0-9.-]+/g,""));
  return isNaN(num) ? 0 : num;
}

function callGeminiApi(file, apiKey, promptText) {
  const MODEL_NAME = 'gemini-2.5-pro'; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${apiKey}`;
  const blob = file.getBlob();
  const base64Data = Utilities.base64Encode(blob.getBytes());
  const payload = {
    contents: [{ parts: [{ text: promptText }, { inline_data: { mime_type: blob.getContentType(), data: base64Data } }] }],
    generationConfig: { response_mime_type: "application/json" }
  };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    if (response.getResponseCode() !== 200) throw new Error(json.error.message);
    let text = json.candidates[0].content.parts[0].text;
    text = text.replace(/^```json\s*/, "").replace(/\s*```$/, "");
    return JSON.parse(text);
  } catch (e) { throw new Error("API解析失敗: " + e.message); }
}

function getPromptFromSheet(ssId) {
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName(PROMPT_SHEET_NAME);
  return sheet ? sheet.getRange('A1').getValue() : "Default Prompt";
}

function saveToSpreadsheet(ssId, records) {
  const ss = SpreadsheetApp.openById(ssId);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ['branch_name', 'file_id', 'file_name', 'status', 'order_date',
                    'maker_name', 'shop_name', 'product_code', 'product_name',
                    'unit_price', 'quantity', 'line_total', 'processed_at',
                    'delivery_destination', 'order_number', 'comment', 'item_order', 'page_number'];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  const rows = records.map(r => [
    r.branch_name, r.file_id, r.file_name, r.status,
    r.order_date, r.maker_name, r.shop_name,
    r.product_code, r.product_name, r.unit_price, r.quantity, r.line_total, r.processed_at,
    r.delivery_destination, r.order_number || '', r.comment || '', r.item_order || 0, r.page_number || 1
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function getOrCreateTargetFolder(root, name) {
  const iter = root.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : root.createFolder(name);
}

/**
 * データ取得：画面表示用にスプレッドシートから読み込む
 */
function getDataFromSpreadsheet() {
  console.log("[SERVER] データ取得開始");
  try {
    const ssId = PROPS.getProperty('SPREADSHEET_ID');
    if (!ssId) throw new Error("スプレッドシートIDが未設定");

    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`シート「${SHEET_NAME}」が見つかりません`);

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const MAX_ROWS = 500;
    const startRow = Math.max(2, lastRow - MAX_ROWS + 1);
    const numRows = lastRow - startRow + 1;
    const data = sheet.getRange(startRow, 1, numRows, 18).getValues();
    
    const formattedData = data.map((row, index) => {
      try {
        const rowIndex = startRow + index;
        
        let dateStr = "";
        try {
          const rawDate = row[4];
          if (rawDate instanceof Date) {
            dateStr = Utilities.formatDate(rawDate, "JST", "yyyy-MM-dd");
          } else if (rawDate) {
            const d = new Date(rawDate);
            dateStr = !isNaN(d.getTime()) ? Utilities.formatDate(d, "JST", "yyyy-MM-dd") : String(rawDate);
          }
        } catch (dErr) { dateStr = "(日付エラー)"; }

        const toNum = (val) => {
          if (typeof val === 'number') return val;
          const n = Number(String(val).replace(/[^0-9.-]/g, ''));
          return isNaN(n) ? 0 : n;
        };

        return {
          uniqueKey: (row[1] || 'unknown') + '_' + rowIndex,
          branch: String(row[0] || ""),
          fileId: String(row[1] || ""),
          fileName: String(row[2] || ""),
          status: String(row[3] || ""),
          date: dateStr,
          maker: String(row[5] || ""),
          shop: String(row[6] || ""),
          dest: String(row[13] || ""),
          p_code: String(row[7] || ""),
          p_name: String(row[8] || ""),
          price: toNum(row[9]),
          qty: toNum(row[10]),
          total: toNum(row[11]),
          order_number: String(row[14] || ""),
          comment: String(row[15] || ""),
          item_order: toNum(row[16]),
          page_number: toNum(row[17]) || 1
        };
      } catch (rowErr) {
        return { uniqueKey: 'error_' + index, branch: 'データ破損', items: [] };
      }
    });

    // 完了ステータスのデータを除外
    const filteredData = formattedData.filter(item => {
      return item.status !== '完了';
    });

    return filteredData;
  } catch (e) {
    console.error(`[SERVER] 重大エラー: ${e.message}`);
    throw new Error(`サーバーエラー: ${e.message}`);
  }
}

/**
 * Excel出力：選択されたデータをテンポラリSS経由でxlsx出力
 * ★納品先を一番右に配置
 */
function exportSelectedDataToExcel(selectedKeys, autoArchive = true) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('SPREADSHEET_ID'));
  const sheet = ss.getSheetByName(SHEET_NAME);
  const allValues = sheet.getDataRange().getValues();
  const exportTargets = [];
  const exportedFileIds = new Set();

  selectedKeys.forEach(key => {
    const idx = key.lastIndexOf('_');
    if (idx === -1) return;
    const fileId = key.substring(0, idx);
    const rowNum = parseInt(key.substring(idx + 1), 10);
    const arrIdx = rowNum - 1;
    if (allValues[arrIdx] && String(allValues[arrIdx][1]) === fileId) {
      exportTargets.push(allValues[arrIdx]);
      exportedFileIds.add(fileId);
    }
  });

  if (!exportTargets.length) throw new Error("データなし");

  const tempSS = SpreadsheetApp.create("Export_" + Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmm"));
  const tempSheet = tempSS.getSheets()[0];
  
  // ヘッダー更新（発注No追加）
  const headers = ['拠点', '発注日', '発注No', 'メーカー', '店舗名', '品番', '商品名', '単価', '数量', '小計', 'ステータス', '納品先'];
  tempSheet.appendRow(headers);

  const rows = exportTargets.map(r => [
    r[0], r[4], r[14], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[3], r[13]
  ]);
  tempSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  const exportUrl = `https://docs.google.com/spreadsheets/d/${tempSS.getId()}/export?format=xlsx`;

  // 自動アーカイブ処理（データは保持）
  if (autoArchive && exportedFileIds.size > 0) {
    try {
      const fileIdsArray = Array.from(exportedFileIds);
      const archiveResult = archiveOrdersByFileIds(fileIdsArray);
      console.log(`[EXPORT+ARCHIVE] ${archiveResult.updatedCount}行を「完了」に変更`);
    } catch (e) {
      console.error(`[EXPORT] 自動アーカイブ失敗: ${e.message}`);
    }
  }

  return exportUrl;
}

/**
 * データ更新：画面からの編集内容を保存
 */
function updateOrderData(updates) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('SPREADSHEET_ID'));
  const sheet = ss.getSheetByName(SHEET_NAME);

  updates.forEach(update => {
    const idx = update.uniqueKey.lastIndexOf('_');
    if (idx === -1) return;
    const rowNum = parseInt(update.uniqueKey.substring(idx + 1), 10);

    // ファイル名更新
    if (update.fileName !== undefined) {
      sheet.getRange(rowNum, 3).setValue(update.fileName);
    }

    // ステータス・コメント更新
    if (update.status !== undefined) {
      sheet.getRange(rowNum, 4).setValue(update.status);
    }
    if (update.comment !== undefined) {
      sheet.getRange(rowNum, 16).setValue(update.comment);
    }

    // 既存フィールド更新
    sheet.getRange(rowNum, 5).setValue(update.date);
    sheet.getRange(rowNum, 15).setValue(update.orderNum);
    sheet.getRange(rowNum, 6).setValue(update.maker);
    sheet.getRange(rowNum, 7).setValue(update.shop);
    sheet.getRange(rowNum, 14).setValue(update.dest);
    sheet.getRange(rowNum, 8).setValue(update.p_code);
    sheet.getRange(rowNum, 9).setValue(update.p_name);
    sheet.getRange(rowNum, 10).setValue(update.price);
    sheet.getRange(rowNum, 11).setValue(update.qty);
    sheet.getRange(rowNum, 12).setFormula(`=J${rowNum}*K${rowNum}`);

    // page_number更新（R列 = 18列目）
    if (update.page_number !== undefined) {
      sheet.getRange(rowNum, 18).setValue(update.page_number);
    }
  });
  return "保存しました";
}

/**
 * アーカイブ機能：選択されたfileIdを「完了」ステータスに変更
 * データは物理削除せず、スプレッドシートに保持される
 */
function archiveOrdersByFileIds(fileIds) {
  console.log(`[ARCHIVE] 開始 - 対象件数: ${fileIds.length}`);

  if (!fileIds || fileIds.length === 0) {
    throw new Error("対象が指定されていません");
  }

  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) return { updatedCount: 0, message: "対象データなし" };

    // 対象行を収集してステータス更新
    let updatedCount = 0;
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();  // branch, file_id

    for (let i = 0; i < data.length; i++) {
      const fileId = String(data[i][1]);
      if (fileIds.includes(fileId)) {
        const rowNum = i + 2;
        sheet.getRange(rowNum, 4).setValue('完了');  // Column D (status)
        updatedCount++;
      }
    }

    console.log(`[ARCHIVE] 完了 - ${updatedCount}行を「完了」に変更`);
    return {
      updatedCount: updatedCount,
      fileCount: fileIds.length,
      message: `${fileIds.length}ファイル(${updatedCount}行)を完了しました`
    };

  } catch (e) {
    console.error(`[ARCHIVE] エラー: ${e.message}`);
    throw new Error(`アーカイブ処理失敗: ${e.message}`);
  }
}

/**
 * フォルダ診断用
 */
function debugFolderCheck() {
  const props = PropertiesService.getScriptProperties();
  const rootId = props.getProperty('ROOT_IN_FOLDER_ID');
  if (!rootId) return console.error("ROOT_IN_FOLDER_ID 未設定");

  try {
    const folder = DriveApp.getFolderById(rootId);
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      const sub = subFolders.next();
      console.log(`📁 ${sub.getName()}`);
      const files = sub.getFilesByType(MimeType.PDF);
      while (files.hasNext()) console.log(`  📄 ${files.next().getName()}`);
    }
  } catch (e) { console.error(e.message); }
}

/**
 * PDFの読み取り結果を詳細ログ出力（診断専用）
 * フォルダ構造確認 → PDF検出 → Gemini API呼び出し → 結果出力
 */
function debugOcrWithDetailedLog() {
  console.log('========================================');
  console.log('【診断開始】PDF読み取りテスト');
  console.log('========================================\n');

  try {
    // 1. スクリプトプロパティの確認
    console.log('--- ステップ1: スクリプトプロパティ確認 ---');
    const rootInId = PROPS.getProperty('ROOT_IN_FOLDER_ID');
    const rootProcessedId = PROPS.getProperty('ROOT_PROCESSED_FOLDER_ID');
    const ssId = PROPS.getProperty('SPREADSHEET_ID');
    const apiKey = PROPS.getProperty('GEMINI_API_KEY');

    console.log(`ROOT_IN_FOLDER_ID: ${rootInId ? '✓ 設定済み' : '✗ 未設定'}`);
    console.log(`ROOT_PROCESSED_FOLDER_ID: ${rootProcessedId ? '✓ 設定済み' : '✗ 未設定'}`);
    console.log(`SPREADSHEET_ID: ${ssId ? '✓ 設定済み' : '✗ 未設定'}`);
    console.log(`GEMINI_API_KEY: ${apiKey ? '✓ 設定済み' : '✗ 未設定'}\n`);

    if (!rootInId || !rootProcessedId || !ssId || !apiKey) {
      throw new Error('スクリプトプロパティが不足しています');
    }

    // 2. プロンプト取得
    console.log('--- ステップ2: Geminiプロンプト取得 ---');
    const geminiPrompt = getPromptFromSheet(ssId);
    console.log(`プロンプト長: ${geminiPrompt.length}文字`);
    console.log(`プロンプト内容:\n${geminiPrompt}\n`);

    // 3. フォルダ構造の確認
    console.log('--- ステップ3: フォルダ構造確認 ---');
    const rootInFolder = DriveApp.getFolderById(rootInId);
    console.log(`ルートフォルダ名: ${rootInFolder.getName()}`);

    const branchFolders = rootInFolder.getFolders();
    let totalBranchCount = 0;
    let totalPdfCount = 0;
    const pdfList = [];

    while (branchFolders.hasNext()) {
      const branchFolder = branchFolders.next();
      const branchName = branchFolder.getName();
      totalBranchCount++;

      console.log(`\n📁 拠点フォルダ: ${branchName}`);

      const files = branchFolder.getFilesByType(MimeType.PDF);
      let branchPdfCount = 0;

      while (files.hasNext()) {
        const file = files.next();
        branchPdfCount++;
        totalPdfCount++;

        console.log(`  📄 ${file.getName()} (ID: ${file.getId()})`);
        pdfList.push({ file: file, branchName: branchName });
      }

      if (branchPdfCount === 0) {
        console.log('  ⚠️ PDFファイルなし');
      }
    }

    console.log(`\n拠点フォルダ数: ${totalBranchCount}`);
    console.log(`PDF総数: ${totalPdfCount}\n`);

    if (totalPdfCount === 0) {
      console.warn('⚠️ 処理対象のPDFファイルが見つかりませんでした');
      console.warn('フォルダ構造を確認してください:');
      console.warn('ROOT_IN_FOLDER → 拠点フォルダ → PDFファイル');
      return;
    }

    // 4. 最初の1件だけOCR実行（全件だと時間がかかるため）
    console.log('========================================');
    console.log('--- ステップ4: OCR処理実行（最初の1ファイルのみ）---');
    console.log('========================================\n');

    const firstPdf = pdfList[0];
    const file = firstPdf.file;
    const branchName = firstPdf.branchName;

    console.log(`対象ファイル: ${file.getName()}`);
    console.log(`拠点名: ${branchName}`);
    console.log(`ファイルID: ${file.getId()}`);
    console.log(`ファイルサイズ: ${(file.getSize() / 1024).toFixed(2)} KB\n`);

    // 4-1. ページ数取得
    console.log('--- ステップ4-1: ページ数取得 (Gemini Flash) ---');
    const totalPages = getPageCount(file, apiKey);
    console.log(`✓ 総ページ数: ${totalPages}\n`);

    // 4-2. ページごとにOCR実行
    let totalItemCount = 0;
    const overallStartTime = new Date();

    for (let page = 1; page <= totalPages; page++) {
      console.log('========================================');
      console.log(`--- ステップ5: ページ ${page}/${totalPages} のOCR処理 ---`);
      console.log('========================================\n');

      const pagePrompt = totalPages === 1
        ? geminiPrompt
        : `【重要】このPDFは${totalPages}ページありますが、${page}ページ目の内容のみを対象に読み取ってください。他のページの内容は絶対に含めないでください。\n\n${geminiPrompt}`;

      console.log('Gemini API呼び出し中...\n');
      const startTime = new Date();
      const extractedData = callGeminiApi(file, apiKey, pagePrompt);
      const endTime = new Date();
      const elapsedTime = ((endTime - startTime) / 1000).toFixed(2);
      console.log(`✓ API応答成功（処理時間: ${elapsedTime}秒）\n`);

      console.log('【基本情報】');
      console.log(`  発注日: ${extractedData.order_date || '(なし)'}`);
      console.log(`  発注番号: ${extractedData.order_number || '(なし)'}`);
      console.log(`  メーカー名: ${extractedData.maker_name || '(なし)'}`);
      console.log(`  店舗名: ${extractedData.shop_name || '(なし)'}`);
      console.log(`  納品先: ${extractedData.delivery_destination || '(なし)'}\n`);

      console.log(`【商品明細 - ページ${page}】`);
      if (extractedData.items && extractedData.items.length > 0) {
        console.log(`  明細数: ${extractedData.items.length}件\n`);
        totalItemCount += extractedData.items.length;

        extractedData.items.forEach((item, index) => {
          console.log(`  [${index + 1}] 品番: ${item.product_code || '(なし)'}`);
          console.log(`      商品名: ${item.product_name || '(なし)'}`);
          console.log(`      数量: ${item.quantity || 0}`);
          console.log(`      単価: ${item.unit_price || 0}`);
          console.log(`      小計: ${(safeParseFloat(item.quantity) * safeParseFloat(item.unit_price)).toLocaleString()}円\n`);
        });
      } else {
        console.log('  ⚠️ 商品明細なし\n');
      }

      console.log(`--- 生JSONデータ（ページ${page}） ---`);
      console.log(JSON.stringify(extractedData, null, 2));
      console.log('\n');
    }

    const overallElapsed = ((new Date() - overallStartTime) / 1000).toFixed(2);

    // まとめ
    console.log('========================================');
    console.log('【診断完了】');
    console.log('========================================');
    console.log(`✓ フォルダ構造: OK（拠点数: ${totalBranchCount}, PDF数: ${totalPdfCount}）`);
    console.log(`✓ ページ数: ${totalPages}ページ`);
    console.log(`✓ API呼び出し: OK（総処理時間: ${overallElapsed}秒）`);
    console.log(`✓ データ抽出: OK（総明細数: ${totalItemCount}件）\n`);

    console.log('※ 実際の処理を実行するには processOrders() を実行してください');

  } catch (e) {
    console.error('\n========================================');
    console.error('【エラー発生】');
    console.error('========================================');
    console.error(`エラーメッセージ: ${e.message}`);
    console.error(`スタックトレース:\n${e.stack}`);
    throw e;
  }
}

/**
 * スプレッドシートのスキーマ診断
 * 現在の列構成とV2スキーマとの比較を表示
 */
function debugSchemaCheck() {
  console.log('========================================');
  console.log('【スキーマ診断】スプレッドシート列構成チェック');
  console.log('========================================\n');

  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      console.error(`エラー: シート「${SHEET_NAME}」が見つかりません`);
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    console.log(`シート名: ${SHEET_NAME}`);
    console.log(`データ行数: ${lastRow}行`);
    console.log(`列数: ${lastColumn}列\n`);

    // ヘッダー行を取得
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    console.log('--- 現在の列構成 ---');
    headers.forEach((header, index) => {
      const colLetter = String.fromCharCode(65 + index); // A=65
      console.log(`${colLetter}列 (${index + 1}): ${header}`);
    });

    // V2スキーマの期待値
    const expectedHeaders = [
      'branch_name',
      'file_id',
      'file_name',
      'status',
      'order_date',
      'maker_name',
      'shop_name',
      'product_code',
      'product_name',
      'unit_price',
      'quantity',
      'line_total',
      'processed_at',
      'delivery_destination',
      'order_number',
      'comment',
      'item_order',
      'page_number'
    ];

    console.log('\n--- V2スキーマとの比較 ---');
    let hasError = false;

    for (let i = 0; i < expectedHeaders.length; i++) {
      const colLetter = String.fromCharCode(65 + i);
      const expected = expectedHeaders[i];
      const actual = headers[i] || '(なし)';
      const match = expected === actual;

      if (match) {
        console.log(`✓ ${colLetter}列: ${expected}`);
      } else {
        console.error(`✗ ${colLetter}列: 期待=${expected}, 実際=${actual}`);
        hasError = true;
      }
    }

    console.log('\n========================================');
    if (hasError || lastColumn < 18) {
      console.warn('【結果】スキーマが一致しません');
      console.warn(`現在: ${lastColumn}列, 必要: 18列\n`);
      console.warn('👉 migrateToV3Schema() を実行してください');
    } else {
      console.log('【結果】✓ V2スキーマと一致しています');
    }
    console.log('========================================');

    // サンプルデータを1行表示
    if (lastRow > 1) {
      console.log('\n--- サンプルデータ（2行目） ---');
      const sampleRow = sheet.getRange(2, 1, 1, lastColumn).getValues()[0];
      sampleRow.forEach((value, index) => {
        const colLetter = String.fromCharCode(65 + index);
        const header = headers[index] || `列${index + 1}`;
        console.log(`${colLetter}列 (${header}): ${value}`);
      });
    }

  } catch (e) {
    console.error(`エラー: ${e.message}`);
    throw e;
  }
}

/**
 * 既存データをV2スキーマ（17列）にマイグレーション
 * ※初回デプロイ時に1度だけ実行
 */
function migrateToV2Schema() {
  console.log('========================================');
  console.log('【マイグレーション開始】V1 → V2スキーマ');
  console.log('========================================\n');

  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      console.log('データなし: マイグレーション不要');
      return "データなし";
    }

    const lastColumn = sheet.getLastColumn();
    console.log(`現在の列数: ${lastColumn}列`);
    console.log(`データ行数: ${lastRow - 1}行\n`);

    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    if (lastColumn >= 17) {
      console.log('既にマイグレーション済み（17列以上）');
      return "既にマイグレーション済み";
    }

    console.log('--- ステップ1: 列ヘッダー追加 ---');

    // O列（15列目）: order_number
    sheet.getRange(1, 15).setValue('order_number');
    console.log('✓ O列 (15): order_number 追加');

    // P列（16列目）: comment
    sheet.getRange(1, 16).setValue('comment');
    console.log('✓ P列 (16): comment 追加');

    // Q列（17列目）: item_order
    sheet.getRange(1, 17).setValue('item_order');
    console.log('✓ Q列 (17): item_order 追加\n');

    console.log('--- ステップ2: 既存データに初期値設定 ---');

    // 既存データに初期値を設定
    const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B列（file_id）を取得
    let currentFileId = null;
    let itemOrder = 0;
    let processedCount = 0;

    for (let i = 0; i < data.length; i++) {
      const fileId = String(data[i][0]);
      const rowNum = i + 2;

      // file_idが変わったら商品順序をリセット
      if (fileId !== currentFileId) {
        currentFileId = fileId;
        itemOrder = 1;
      } else {
        itemOrder++;
      }

      // O列: order_number（空文字）
      sheet.getRange(rowNum, 15).setValue('');
      // P列: comment（空文字）
      sheet.getRange(rowNum, 16).setValue('');
      // Q列: item_order
      sheet.getRange(rowNum, 17).setValue(itemOrder);

      processedCount++;

      if (processedCount % 50 === 0) {
        console.log(`処理中... ${processedCount}/${data.length}行`);
      }
    }

    console.log(`✓ ${processedCount}行のデータを更新\n`);

    console.log('========================================');
    console.log('【マイグレーション完了】');
    console.log('========================================');
    console.log(`処理行数: ${lastRow - 1}行`);
    console.log(`列数: 14列 → 17列\n`);
    console.log('✓ order_number, comment, item_order を追加しました');

    return `マイグレーション完了: ${lastRow - 1}行`;

  } catch (e) {
    console.error(`マイグレーションエラー: ${e.message}`);
    throw e;
  }
}

/**
 * 既存データをV3スキーマ（18列）にマイグレーション
 * page_number列を追加（デフォルト値: 1）
 * ※初回デプロイ時に1度だけ実行
 */
function migrateToV3Schema() {
  console.log('========================================');
  console.log('【マイグレーション開始】V2 → V3スキーマ (page_number追加)');
  console.log('========================================\n');

  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastColumn >= 18) {
      console.log('既にマイグレーション済み（18列以上）');
      return "既にマイグレーション済み";
    }

    if (lastRow <= 1) {
      sheet.getRange(1, 18).setValue('page_number');
      console.log('✓ R列 (18): page_number ヘッダー追加（データなし）');
      return "ヘッダーのみ追加";
    }

    // R列（18列目）: page_number ヘッダー追加
    sheet.getRange(1, 18).setValue('page_number');
    console.log('✓ R列 (18): page_number 追加');

    // 既存データに初期値 1 を一括設定
    const numRows = lastRow - 1;
    const defaultValues = Array(numRows).fill([1]);
    sheet.getRange(2, 18, numRows, 1).setValues(defaultValues);
    console.log(`✓ ${numRows}行にデフォルト値 page_number=1 を設定`);

    console.log('\n========================================');
    console.log('【マイグレーション完了】V2 → V3');
    console.log('========================================');
    console.log(`処理行数: ${numRows}行`);
    console.log(`列数: 17列 → 18列`);

    return `マイグレーション完了: ${numRows}行`;

  } catch (e) {
    console.error(`マイグレーションエラー: ${e.message}`);
    throw e;
  }
}