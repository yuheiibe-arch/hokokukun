// ==========================================
// ▼既存の processAggregationCore と差し替えてください
// ==========================================
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
    
    // ★★★ 新規追加：「年度平均時給」シートへ当月分を自動追記 ★★★
    updateAnnualSheetSingleMonth(ss, targetDate, allData.current);
    
    ss.toast('すべての集計と書き込みが完了しました！', '完了', 5);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー', '処理中にエラーが発生しました。\n' + e.stack, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==========================================
// ▼上記の下に、そのまま新規追加してください
// ==========================================
function updateAnnualSheetSingleMonth(ss, targetDate, monthData) {
  const targetSheet = ss.getSheetByName('年度平均時給');
  if (!targetSheet) return;

  let lastCol = targetSheet.getLastColumn();
  if (lastCol < 2) return;
  const topHeaders = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let colIdx = -1;
  const tYear = targetDate.getFullYear();
  const tMonth = targetDate.getMonth(); // 0始まり

  for (let c = 1; c < lastCol; c++) {
    const d = topHeaders[c];
    if (d instanceof Date) {
      if (d.getFullYear() === tYear && d.getMonth() === tMonth) {
        colIdx = c + 1; // getRange用の1始まりインデックス
        break;
      }
    }
  }

  // ★自己拡張：該当月の列がなければ右端に自動で列を追加する
  if (colIdx === -1) {
    targetSheet.insertColumnAfter(lastCol);
    lastCol++;
    colIdx = lastCol;
    targetSheet.getRange(1, colIdx).setValue(targetDate).setNumberFormat('yyyy/MM');
    // 左隣から背景色や罫線の書式のみをコピー
    targetSheet.getRange(1, colIdx - 1, 46, 1).copyTo(targetSheet.getRange(1, colIdx, 46, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }

  const baseRowMap = { '平均時給': 0, '稼働人員': 12, '対象拠点数': 24, '医師勤務時間': 36 };
  const areaOffsetMap = { '関東': 0, '関西': 1, '関東第一': 2, '関東第二': 3, '埼玉': 4, '神奈川': 5, '千葉': 6, '大阪': 7, '茨城': 8, 'グループ全体': 9 };

  const outputValues = targetSheet.getRange(2, colIdx, 46, 1).getValues();
  const outputColors = targetSheet.getRange(2, colIdx, 46, 1).getBackgrounds();

  Object.keys(areaOffsetMap).forEach(area => {
    const d = monthData[area];
    if (!d) return;

    const rWage = baseRowMap['平均時給'] + areaOffsetMap[area];
    const rUU   = baseRowMap['稼働人員'] + areaOffsetMap[area];
    const rBase = baseRowMap['対象拠点数'] + areaOffsetMap[area];
    const rTime = baseRowMap['医師勤務時間'] + areaOffsetMap[area];

    // 拠点稼働があれば白背景で書き込み
    if (d.baseCount.total > 0 || d.uu.total > 0) {
      outputValues[rWage][0] = d.wage.areaAvg > 0 ? d.wage.areaAvg : '';
      outputColors[rWage][0] = '#ffffff';

      outputValues[rUU][0] = d.uu.total > 0 ? d.uu.total : '';
      outputColors[rUU][0] = '#ffffff';

      outputValues[rBase][0] = d.baseCount.total > 0 ? d.baseCount.total : '';
      outputColors[rBase][0] = '#ffffff';

      outputValues[rTime][0] = d.ratioTime.total > 0 ? (d.ratioTime.total / 60).toFixed(1) : '';
      outputColors[rTime][0] = '#ffffff';
    } else {
      // グレーアウト処理
      const gray = '#f0f0f0';
      outputValues[rWage][0] = ''; outputColors[rWage][0] = gray;
      outputValues[rUU][0] = '';   outputColors[rUU][0] = gray;
      outputValues[rBase][0] = ''; outputColors[rBase][0] = gray;
      outputValues[rTime][0] = ''; outputColors[rTime][0] = gray;
    }
  });

  targetSheet.getRange(2, colIdx, 46, 1).setValues(outputValues).setBackgrounds(outputColors);
}