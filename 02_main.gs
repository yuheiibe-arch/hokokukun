// ==========================================
// ▼ メイン処理（既存の processAggregationCore と差し替えてください）
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
    
    // ★★★ 既存機能：「年度平均時給」シートへ当月分を自動追記 ★★★
    updateAnnualSheetSingleMonth(ss, targetDate, allData.current);

    // ★★★ 新規追加：「平均時給推移グラフ」シートへ直接書き込み ★★★
    updateGraphSheetVertical(ss, targetDate, allData.current);
    
    ss.toast('すべての集計と書き込みが完了しました！', '完了', 5);
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー', '処理中にエラーが発生しました。\n' + e.stack, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==========================================
// ★ 修正強化版：年度平均時給シートへの書き込み（日付型/文字列型の両方に対応）
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

  // 横方向（1行目の列ヘッダー）のスキャン (文字列とDate型の両方に対応)
  for (let c = 1; c < lastCol; c++) {
    const d = topHeaders[c];
    let dObj = null;

    if (d instanceof Date) {
      dObj = d;
    } else if (d && (typeof d === 'string' || d instanceof String)) {
      const cleanStr = String(d).trim().replace(/-/g, '/');
      const parsed = new Date(cleanStr.includes('/') ? cleanStr + '/01' : cleanStr);
      if (!isNaN(parsed.getTime())) dObj = parsed;
    }

    if (dObj && dObj.getFullYear() === tYear && dObj.getMonth() === tMonth) {
      colIdx = c + 1; // getRange用の1始まりインデックス
      break;
    }
  }

  const lastRow = targetSheet.getLastRow();

  // ★自己拡張：該当月の列がなければ右端に自動で列を追加する
  if (colIdx === -1) {
    // 完全に空っぽではない実際の最終列を再計算して、その右に追加
    let realLastCol = 1;
    for (let c = topHeaders.length - 1; c >= 0; c--) {
      if (topHeaders[c] !== "" && topHeaders[c] !== null && topHeaders[c] !== undefined) {
        realLastCol = c + 1;
        break;
      }
    }
    targetSheet.insertColumnAfter(realLastCol);
    colIdx = realLastCol + 1;
    targetSheet.getRange(1, colIdx).setValue(targetDate).setNumberFormat('yyyy/MM');
    
    // 左隣から背景色や罫線の書式のみをコピー（固定行数ではなく動的に最終行まで）
    if (lastRow > 1) {
      targetSheet.getRange(1, colIdx - 1, lastRow, 1).copyTo(targetSheet.getRange(1, colIdx, lastRow, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
  }

  if (lastRow < 2) return;

  // A列のみを取得してパース
  const aValues = targetSheet.getRange(1, 1, lastRow, 1).getDisplayValues();
  const metrics = ['平均時給', '稼働人員', '対象拠点数', '医師勤務時間'];
  // ★ ここを「東京第一」「東京第二」に修正
  const targetAreas = ['関東', '関西', '東京第一', '東京第二', '埼玉', '神奈川', '千葉', '大阪', '茨城', 'グループ全体'];

  const rowMap = {};
  let currentMetric = '';

  // A列を上からスキャンして、「指標_エリア」に対応する正確な行番号をマッピング
  for (let i = 0; i < aValues.length; i++) {
    const cellText = String(aValues[i][0]).replace(/\s+/g, '');
    if (!cellText) continue;

    // A列のセルに指標名が含まれていたら、現在のブロックを切り替え
    const foundMetric = metrics.find(m => cellText.includes(m));
    if (foundMetric) {
      currentMetric = foundMetric;
      continue;
    }

    // ブロックが特定されている状態で、エリア名が完全一致したら行番号を記録
    if (currentMetric) {
      const foundArea = targetAreas.find(a => cellText === a);
      if (foundArea) {
        rowMap[`${currentMetric}_${foundArea}`] = i + 1;
      }
    }
  }

  // 特定した行に対してデータを安全に書き込み
  targetAreas.forEach(area => {
    const d = monthData[area];
    if (!d) return;

    // 拠点稼働があるかチェック
    const isWorking = (d.baseCount.total > 0 || d.uu.total > 0);

    metrics.forEach(metric => {
      const targetRowNum = rowMap[`${metric}_${area}`];
      
      if (targetRowNum) {
        const range = targetSheet.getRange(targetRowNum, colIdx);
        if (isWorking) {
          range.setBackground('#ffffff');
          let val = '';
          if (metric === '平均時給' && d.wage.areaAvg > 0) val = d.wage.areaAvg;
          if (metric === '稼働人員' && d.uu.total > 0) val = d.uu.total;
          if (metric === '対象拠点数' && d.baseCount.total > 0) val = d.baseCount.total;
          if (metric === '医師勤務時間' && d.ratioTime.total > 0) val = (d.ratioTime.total / 60).toFixed(1);
          range.setValue(val);
        } else {
          // 稼働がない場合はグレーアウト処理
          range.setBackground('#f0f0f0');
          range.setValue('');
        }
      }
    });
  });
}

// ==========================================
// ★ 【新規追加】平均時給推移グラフシート（縦方向）への直接書き込み処理
// ==========================================
function updateGraphSheetVertical(ss, targetDate, monthData) {
  const sheet = ss.getSheetByName('平均時給推移グラフ');
  if (!sheet) {
    console.error('「平均時給推移グラフ」シートが見つかりません。');
    return;
  }

  // 探したい年月（例: "2026/06"）
  const targetMonthStr = Utilities.formatDate(targetDate, "GMT+9", "yyyy/MM");
  const lastRow = sheet.getLastRow();
  
  // 画像から推測し、A列とB列の値を一気に取得して日付を探す
  const dateValues = sheet.getRange(1, 1, lastRow, 2).getDisplayValues(); 
  let targetRow = -1;
  let dateCol = 2; // 日付がある列（デフォルトはB列と想定）

  for (let i = 0; i < dateValues.length; i++) {
    if (String(dateValues[i][0]).includes(targetMonthStr)) { targetRow = i + 1; dateCol = 1; break; }
    if (String(dateValues[i][1]).includes(targetMonthStr)) { targetRow = i + 1; dateCol = 2; break; }
  }

  if (targetRow === -1) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`「平均時給推移グラフ」に ${targetMonthStr} の行がありません。手動で日付を追加してください。`, '警告');
    return;
  }

  // ★ ここを「東京第一」「東京第二」に修正（画像の列順に合わせる）
  const areaCols = {
    '関東': dateCol + 1,
    '関西': dateCol + 2,
    '東京第一': dateCol + 3,
    '東京第二': dateCol + 4,
    '埼玉': dateCol + 5,
    '神奈川': dateCol + 6,
    '千葉': dateCol + 7,
    '大阪': dateCol + 8,
    '茨城': dateCol + 9,
    'グループ全体': dateCol + 10
  };

  // 各エリアの平均時給を該当セルに書き込む
  Object.keys(areaCols).forEach(area => {
    if (monthData[area] && monthData[area].wage && monthData[area].wage.areaAvg > 0) {
      sheet.getRange(targetRow, areaCols[area]).setValue(monthData[area].wage.areaAvg);
    }
  });
}