function restoreDashboardFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('📊 年度別トレンドボード');
  const archiveSheet = ss.getSheetByName('月次アーカイブ');
  const debugArchive = ss.getSheetByName('🛠️デバッグ用アーカイブ');
  
  if (!dbSheet) {
    SpreadsheetApp.getUi().alert('ダッシュボードが見つかりません。');
    return;
  }

  // ==========================================
  // 1. デバッグ用データを本番アーカイブに移行（数字を復活させるため）
  // ==========================================
  if (debugArchive && archiveSheet) {
    const dData = debugArchive.getDataRange().getValues();
    if (dData.length > 1) {
      const dataToCopy = dData.slice(1); // 見出しを除く
      // 重複を防ぐため一旦クリアして書き直す
      if (archiveSheet.getLastRow() > 1) {
        archiveSheet.getRange(2, 1, archiveSheet.getLastRow() - 1, archiveSheet.getLastColumn()).clearContent();
      }
      archiveSheet.getRange(2, 1, dataToCopy.length, dataToCopy[0].length).setValues(dataToCopy);
      console.log('✅ 4月・5月の計算データを本番アーカイブに移行しました。');
    }
  }

  // ==========================================
  // 2. 指標名(C列)を読み取って、正しい数式を埋め込む
  // ==========================================
  const cVals = dbSheet.getRange(1, 3, dbSheet.getLastRow(), 1).getValues();
  
  // [識別キーワード, 参照列, 集計タイプ]
  const metrics = [
    ['来院数実績', 'T', 'SUM'],
    ['売上実績', 'U', 'SUM'],
    ['エリア平均時給', 'D', 'AVG'],
    ['依頼手当総額', 'E', 'SUM'],
    ['稼働人員（全医師）', 'J', 'AVG'],
    ['稼働人員（常勤除く）', '', 'SUM_NON_REG'],
    ['所定休出医師数', 'O', 'AVG'],
    ['常勤比率', 'K', 'RATIO'],
    ['定期非常勤比率', 'L', 'RATIO'],
    ['直応募(スポット)比率', 'M', 'RATIO'],
    ['紹介会社比率', 'N', 'RATIO'],
    ['２診時間', 'R', 'SUM'],
    ['不在時間', 'P', 'SUM'],
    ['残業時間', 'Q', 'SUM'],
    ['対象拠点数', 'S', 'AVG'],
    ['新規採用（直接', 'F', 'SUM'],
    ['新規採用（紹介', 'G', 'SUM'],
    ['在籍医師数（常勤', 'H', 'AVG'],
    ['在籍医師数（定期', 'I', 'AVG']
  ];

  for (let r = 0; r < cVals.length; r++) {
    const metricName = String(cVals[r][0]).replace(/\s/g, ''); // 空白を詰めて判定
    if (!metricName) continue;
    
    const matched = metrics.find(m => metricName.includes(m[0].replace(/\s/g, '')));
    if (matched) {
      const colL = matched[1];
      const aggType = matched[2];
      
      // D$4 を基準にセットすれば、E列以降は E$4, F$4... とスプレッドシートが勝手に調整してくれます
      const cellMonth = 'D$4'; 
      const baseCond = `'月次アーカイブ'!$B:$B, ${cellMonth}, '月次アーカイブ'!$C:$C, $C$3`;
      const checkExists = `COUNTIFS('月次アーカイブ'!$B:$B, ${cellMonth})>0`;

      let formula = '';
      if (aggType === 'SUM') {
        formula = `=IF(${checkExists}, SUMIFS('月次アーカイブ'!$${colL}:$${colL}, ${baseCond}), "")`;
      } else if (aggType === 'AVG') {
        formula = `=IF(${checkExists}, IFERROR(AVERAGEIFS('月次アーカイブ'!$${colL}:$${colL}, ${baseCond}), 0), "")`;
      } else if (aggType === 'SUM_NON_REG') {
        formula = `=IF(${checkExists}, SUMIFS('月次アーカイブ'!$L:$L, ${baseCond}) + SUMIFS('月次アーカイブ'!$M:$M, ${baseCond}) + SUMIFS('月次アーカイブ'!$N:$N, ${baseCond}) + SUMIFS('月次アーカイブ'!$O:$O, ${baseCond}), "")`;
      } else if (aggType === 'RATIO') {
        formula = `=IF(${checkExists}, IFERROR(AVERAGEIFS('月次アーカイブ'!$${colL}:$${colL}, ${baseCond}) / AVERAGEIFS('月次アーカイブ'!$J:$J, ${baseCond}), 0), "")`;
      }

      // D列〜O列 (12ヶ月分) に数式を一括セット
      if (formula) {
        dbSheet.getRange(r + 1, 4, 1, 12).setFormula(formula);
      }
    }
  }
  
  // 用済みのデバッグシートを非表示にする
  if (debugArchive) debugArchive.hideSheet();
  if (archiveSheet) archiveSheet.hideSheet();

  console.log('✨ カスタムされたレイアウトを保持したまま、数式を復旧しました！');
}