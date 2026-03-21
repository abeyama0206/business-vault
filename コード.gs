/**
 * メインの処理を実行する関数
 * @version 1.2 - Time Basedトリガー対応。getUi().alert() 削除。待機時間を3秒に変更。
 */
function executeReviewMaterialCreation() {
  try {
    console.log("処理を開始します。");

    // --- 1. 日付とシートの準備 ---
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const tomorrow = new Date();
    tomorrow.setDate(today.getDate() + 1);
    tomorrow.setHours(0, 0, 0, 0);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('振り返り資料作成');
    if (!sheet) {
      throw new Error('「振り返り資料作成」シートが見つかりませんでした。');
    }
    console.log("シートの取得に成功しました。");

    // --- 2. 今年度の4月1日を計算 ---
    let fiscalYearStart;
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth();

    if (currentMonth < 3) { // 1月, 2月, 3月
      fiscalYearStart = new Date(currentYear - 1, 3, 1); // 前年の4月1日
    } else { // 4月以降
      fiscalYearStart = new Date(currentYear, 3, 1); // 今年の4月1日
    }
    console.log(`年度開始日を ${Utilities.formatDate(fiscalYearStart, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd')} に設定しました。`);

    // --- 3. シートから明日付のorgIdを収集 ---
    console.log("シートから明日付のorgIdの収集を開始します。");
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const targetOrgIds = [];

    for (let i = 1; i < values.length; i++) { // 1行目はヘッダーと想定
      const row = values[i];
      const orgId = row[0];
      const dateValue = row[2];

      if (orgId && dateValue instanceof Date) {
        const targetDate = new Date(dateValue);
        targetDate.setHours(0, 0, 0, 0);

        if (targetDate.getTime() === tomorrow.getTime()) {
          targetOrgIds.push(orgId);
        }
      }
    }
    console.log(`${targetOrgIds.length}件の対象データが見つかりました。`);

    // --- 4. orgIdごとにライブラリ関数を繰り返し実行 ---
    if (targetOrgIds.length > 0) {
      const startDateString = Utilities.formatDate(fiscalYearStart, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      const endDateString = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      let successCount = 0;
      let errorCount = 0;
      let errorDetails = [];

      for (const orgId of targetOrgIds) {
        try {
          console.log(`【開始】orgId: ${orgId}`);

          // ライブラリとして追加した関数を1件ずつ実行
          CsReviewMaterialAutomation.createCsReviewMaterialSuccessTeam(orgId, startDateString, endDateString);

          console.log(`【成功】orgId: ${orgId}`);
          successCount++;

        } catch (e) {
          // 個別のエラーを捕捉
          console.error(`【エラー】orgId: ${orgId} でエラーが発生しました。メッセージ: ${e.message}`);
          errorCount++;
          errorDetails.push(`- orgId ${orgId}: ${e.message}`);
        }

        // ★★★ 処理の間に3秒の待機時間を設ける ★★★
        Utilities.sleep(3000); // 3秒待機
      }

      // --- 5. 最終結果をまとめてログに出力 ---
      let summaryMessage = `全 ${targetOrgIds.length} 件の処理が完了しました。\n\n成功: ${successCount}件\n失敗: ${errorCount}件`;
      if (errorCount > 0) {
        summaryMessage += `\n\n--- エラー詳細 ---\n${errorDetails.join('\n')}`;
      }
      
      console.log("すべての繰り返し処理が完了しました。\n" + summaryMessage);

    } else {
      console.log('明日付の対象データはありませんでした。処理を終了します。');
    }

  } catch (e) {
    const errorMessage = `処理の途中で致命的なエラーが発生しました。\n\n【エラーメッセージ】\n${e.message}\n\nスタックトレース:\n${e.stack}`;
    
    console.error(errorMessage);
  }
}

/**
 * スプレッドシートに実行メニューを追加する
 * (onOpenはTime Basedトリガーでは実行されませんが、手動実行用に残します)
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カスタムメニュー')
    .addItem('レビュー資料作成実行', 'executeReviewMaterialCreation')
    .addToUi();
}