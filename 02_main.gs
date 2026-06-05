function processAggregationCore(targetDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('【1/3】データソースから各月の実績を抽出しています...', '処理開始', 10);
  
  try {
    // 1. 動的抽出（当月・前月・昨年の3期間分をドライブ等から全て抽出）
    const allData = extractAllData(targetDate);
    
    // 2. バックシート（データベース）へ当月分を保存
    ss.toast('【2/3】データベースに当月の履歴を保存中...', '進行状況', 5);
    saveToBackSheet(ss, allData.current, allData.monthStrs.current);
    
    // 3. ベースシートへの書き込み
    ss.toast('【3/3】ベースシートへ書き込んでいます...', '進行状況', 5);
    writeToBaseSheet(ss, allData);
    
    ss.toast('すべての集計と書き込みが完了しました！', '完了', 5);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー', '処理中にエラーが発生しました。\n' + e.stack, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}