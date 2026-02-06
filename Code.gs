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

// --- 個別ファイル処理 ---
function processSingleFile(file, branchName, targetBranchFolder, ssId, apiKey, promptText) {
  const fileName = file.getName();
  try {
    const extractedData = callGeminiApi(file, apiKey, promptText);
    const items = (extractedData.items && extractedData.items.length > 0) 
                  ? extractedData.items 
                  : [{ product_code: "", product_name: "(明細なし)", quantity: 0, unit_price: 0 }];

    const records = items.map((item, index) => ({
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
      item_order: index + 1
    }));

    records.forEach(r => r.line_total = r.quantity * r.unit_price);
    saveToSpreadsheet(ssId, records);
    file.moveTo(targetBranchFolder);
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
                    'delivery_destination', 'order_number', 'comment', 'item_order'];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  const rows = records.map(r => [
    r.branch_name, r.file_id, r.file_name, r.status,
    r.order_date, r.maker_name, r.shop_name,
    r.product_code, r.product_name, r.unit_price, r.quantity, r.line_total, r.processed_at,
    r.delivery_destination, r.order_number || '', r.comment || '', r.item_order || 0
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
    const data = sheet.getRange(startRow, 1, numRows, 17).getValues();
    
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
          item_order: toNum(row[16])
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
 * 既存データをV2スキーマ（17列）にマイグレーション
 * ※初回デプロイ時に1度だけ実行
 */
function migrateToV2Schema() {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('SPREADSHEET_ID'));
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return "データなし";

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.length >= 17) return "既にマイグレーション済み";

  // 新カラムを追加
  sheet.getRange(1, 15).setValue('order_number');
  sheet.getRange(1, 16).setValue('comment');
  sheet.getRange(1, 17).setValue('item_order');

  // 既存データに初期値を設定
  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  let currentFileId = null;
  let itemOrder = 0;

  for (let i = 0; i < data.length; i++) {
    const fileId = String(data[i][0]);
    const rowNum = i + 2;

    if (fileId !== currentFileId) {
      currentFileId = fileId;
      itemOrder = 1;
    } else {
      itemOrder++;
    }

    sheet.getRange(rowNum, 15).setValue('');
    sheet.getRange(rowNum, 16).setValue('');
    sheet.getRange(rowNum, 17).setValue(itemOrder);
  }

  return `マイグレーション完了: ${lastRow - 1}行`;
}