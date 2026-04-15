/**
 * データ削除・パージ関連処理
 * - スプレッドシート完了レコード削除
 * - 処理済みフォルダ内ファイル削除
 * - 明細行削除
 * - トリガー設定
 */

/**
 * 明細行削除：指定されたuniqueKeyの行をスプレッドシートから物理削除
 */
function deleteOrderRows(keysToDelete) {
  if (!keysToDelete || keysToDelete.length === 0) return '削除対象なし';

  const ss = SpreadsheetApp.openById(PROPS.getProperty('SPREADSHEET_ID'));
  const sheet = ss.getSheetByName(SHEET_NAME);

  // uniqueKeyからスプレッドシートの行番号を抽出
  const rowNumbers = keysToDelete.map(key => {
    const idx = key.lastIndexOf('_');
    if (idx === -1) return -1;
    return parseInt(key.substring(idx + 1), 10);
  }).filter(n => n > 1).sort((a, b) => b - a); // 降順ソート（下から削除）

  let deletedCount = 0;
  rowNumbers.forEach(rowNum => {
    try {
      sheet.deleteRow(rowNum);
      deletedCount++;
    } catch (e) {
      console.warn(`行${rowNum}の削除失敗: ${e.message}`);
    }
  });

  console.log(`[DELETE] ${deletedCount}行を削除`);
  return `${deletedCount}行を削除しました`;
}

/**
 * 完了レコード定期削除
 * ステータスが「完了」の行をスプレッドシートから物理削除する
 * トリガーまたはGASエディタから手動実行可能
 */
function purgeCompletedOrders() {
  console.log('[PURGE] 完了レコード削除処理 開始');

  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      console.log('[PURGE] データなし - スキップ');
      return { deletedCount: 0, message: 'データなし' };
    }

    // D列（4列目）= status を取得
    const statuses = sheet.getRange(2, 4, lastRow - 1, 1).getValues();

    // 削除対象の行番号を収集（1-indexed, ヘッダー行+1から）
    const rowsToDelete = [];
    for (let i = 0; i < statuses.length; i++) {
      if (String(statuses[i][0]) === '完了') {
        rowsToDelete.push(i + 2); // スプレッドシートの行番号
      }
    }

    console.log(`[PURGE] 削除対象: ${rowsToDelete.length}件 / 全${lastRow - 1}件`);

    if (rowsToDelete.length === 0) {
      console.log('[PURGE] 完了レコードなし - スキップ');
      return { deletedCount: 0, message: '完了レコードなし' };
    }

    // 下から削除してインデックスずれを防止
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }

    console.log(`[PURGE] 完了 - ${rowsToDelete.length}行を削除`);
    return {
      deletedCount: rowsToDelete.length,
      message: `${rowsToDelete.length}件の完了レコードを削除しました`
    };

  } catch (e) {
    console.error(`[PURGE] エラー: ${e.message}`);
    throw new Error(`完了レコード削除失敗: ${e.message}`);
  }
}

/**
 * 処理済みフォルダ内のファイルを全削除
 * ROOT_PROCESSED_FOLDER 配下のサブフォルダを巡回し、中のファイルをすべてゴミ箱に移動する
 * サブフォルダ自体は残す（拠点フォルダ構成を維持）
 */
function purgeProcessedFiles() {
  console.log('[PURGE-FILES] 処理済みファイル削除処理 開始');

  try {
    const rootProcessedId = PROPS.getProperty('ROOT_PROCESSED_FOLDER_ID');
    if (!rootProcessedId) throw new Error('ROOT_PROCESSED_FOLDER_ID 未設定');

    const rootFolder = DriveApp.getFolderById(rootProcessedId);
    const subFolders = rootFolder.getFolders();
    let totalDeleted = 0;

    while (subFolders.hasNext()) {
      const subFolder = subFolders.next();
      const folderName = subFolder.getName();
      const files = subFolder.getFiles();
      let count = 0;

      while (files.hasNext()) {
        const file = files.next();
        file.setTrashed(true);
        count++;
      }

      if (count > 0) {
        console.log(`[PURGE-FILES] ${folderName}: ${count}件のファイルを削除`);
      }
      totalDeleted += count;
    }

    console.log(`[PURGE-FILES] 完了 - 合計${totalDeleted}件を削除`);
    return {
      deletedCount: totalDeleted,
      message: `処理済みフォルダから${totalDeleted}件のファイルを削除しました`
    };

  } catch (e) {
    console.error(`[PURGE-FILES] エラー: ${e.message}`);
    throw new Error(`処理済みファイル削除失敗: ${e.message}`);
  }
}

/**
 * 定期パージ：スプレッドシート完了レコード削除 + 処理済みフォルダ内ファイル削除
 * トリガーから呼び出される統合関数
 */
function purgeAll() {
  console.log('[PURGE-ALL] 定期パージ処理 開始');

  const results = {};

  // 1. スプレッドシートの完了レコードを削除
  try {
    results.orders = purgeCompletedOrders();
  } catch (e) {
    results.orders = { error: e.message };
    console.error(`[PURGE-ALL] レコード削除エラー: ${e.message}`);
  }

  // 2. 処理済みフォルダ内のファイルを削除
  try {
    results.files = purgeProcessedFiles();
  } catch (e) {
    results.files = { error: e.message };
    console.error(`[PURGE-ALL] ファイル削除エラー: ${e.message}`);
  }

  console.log(`[PURGE-ALL] 完了: ${JSON.stringify(results)}`);
  return results;
}

/**
 * 定期パージトリガーを設定
 * 毎日深夜2時に purgeAll を自動実行する
 * ※ 1回だけ実行すればトリガーが登録される
 */
function setupPurgeTrigger() {
  // 既存の同名トリガーを削除（重複防止）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'purgeAll') {
      ScriptApp.deleteTrigger(trigger);
      console.log('[TRIGGER] 既存トリガーを削除');
    }
  });

  // 毎日午前2時に実行するトリガーを作成
  ScriptApp.newTrigger('purgeAll')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  console.log('[TRIGGER] purgeAll を毎日午前2時に実行するトリガーを登録しました');
  return 'トリガー登録完了: 毎日午前2時に完了レコード削除 + 処理済みファイル削除を自動実行';
}
