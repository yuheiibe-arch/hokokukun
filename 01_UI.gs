function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🌟 月次報告システム')
    .addItem('⚡ 先月分をワンクリック集計', 'runLastMonthQuick')
    .addSeparator()
    .addItem('📅 期間を指定して集計 (詳細モード)', 'showAggregationDialog')
    .addSeparator()
    .addItem('📝 月次報告を作成 (AIレビュー付)', 'showReportDialog') // ★新規追加
    .addToUi();
}

// ==========================================
// ★ 月次報告用のUIを立ち上げる関数（新規追加）
// ==========================================
function showReportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ReportUI')
      .setTitle('月次報告・AIレビュージェネレーター')
      .setWidth(850)    // ★横幅を広げる
      .setHeight(900);  // ★縦幅を広げる（これ以上大きくすると見切れるPCがあります）
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function runLastMonthQuick() {
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const end = new Date(today.getFullYear(), today.getMonth(), 0);
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('確認', 
    `【先月分】の集計を開始します。\n期間: ${Utilities.formatDate(start, "GMT+9", "yyyy/MM/dd")} 〜 ${Utilities.formatDate(end, "GMT+9", "yyyy/MM/dd")}\n\nよろしいですか？`, 
    ui.ButtonSet.YES_NO);
    
  if (response === ui.Button.YES) {
    processAggregationCore(start); // 月次処理なので開始日を基準とする
  }
}

function processAggregationFromDialog(startStr, endStr) {
  const start = new Date(startStr);
  start.setHours(0,0,0,0);
  processAggregationCore(start);
}

// ==========================================
// ★ 月次報告用のUIを立ち上げる関数（新規追加）
// ==========================================
function showReportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ReportUI')
      .setTitle('月次報告・AIレビュージェネレーター')
      .setWidth(750)
      .setHeight(1000); // ★修正: 見切れないように縦幅を広げました
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}