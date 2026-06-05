function writeToBaseSheet(ss, allData) {
  const sheet = ss.getSheetByName('ベースシート');
  
  // ラベルの完全書き換え
  const metricStartRows = [2, 5, 8, 12, 15, 18, 21, 24, 28, 32, 35, 39, 42, 46, 49, 53, 56];
  metricStartRows.forEach(r => {
    sheet.getRange(r, 3).setValue('対象月');
    sheet.getRange(r + 1, 3).setValue('対象前月');
    sheet.getRange(r + 2, 3).setValue('昨年月');
  });

  const lastCol = Math.max(sheet.getLastColumn(), 13);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const areaCols = {}; 
  for (let c = 3; c < headers.length; c++) {
    const areaName = String(headers[c]).trim();
    if (areaName && areaName !== '(空白)') areaCols[areaName] = c + 1;
  }
  
  // 背景色・文字色リセット
  sheet.getRange(2, 4, 58, lastCol - 3).clearContent().setBackground('#ffffff').setFontColor('#000000').setHorizontalAlignment("left");

  // 実数フォーマット関数
  const formatRaw = (val, type) => {
    if (val === null || val === undefined || isNaN(val) || val === 0) return '-';
    if (type === 'yen') return `¥${Number(val).toLocaleString()}`;
    if (type === 'people') return `${Number(val).toLocaleString()}人`;
    if (type === 'mins') {
       const m = Math.round(Number(val));
       return `${Math.floor(m / 60)}時間${m % 60}分`;
    }
    return Number(val).toLocaleString();
  };

  // 【限定適用】対象月との差分を赤字・青字で出力する関数
  const setDiffValue = (range, curr, past, type) => {
    const c = Number(curr);
    const p = Number(past);
    
    if (isNaN(p) || past === null || past === undefined || p === 0) {
      range.setValue('-').setFontColor('#000000');
      return;
    }

    const pastStr = formatRaw(p, type);

    if (isNaN(c) || curr === null || curr === undefined || c === 0) {
      range.setValue(pastStr).setFontColor('#000000');
      return;
    }

    const diff = c - p;
    let diffStr = '';
    let color = '#000000';

    if (diff > 0) {
      diffStr = ` (+${formatRaw(diff, type)})`;
      color = '#ff0000'; // プラスなら赤字
    } else if (diff < 0) {
      diffStr = ` (-${formatRaw(Math.abs(diff), type)})`;
      color = '#0000ff'; // マイナスなら青字
    } else {
      diffStr = ` (±0)`;
    }

    const fullStr = `${pastStr}${diffStr}`;
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(fullStr)
      .setTextStyle(0, pastStr.length, SpreadsheetApp.newTextStyle().setForegroundColor('#000000').build())
      .setTextStyle(pastStr.length, fullStr.length, SpreadsheetApp.newTextStyle().setForegroundColor(color).build())
      .build();
      
    range.setRichTextValue(richText);
  };

  const calcRate = (act, tgt) => (!tgt || tgt === 0) ? '-' : ((act / tgt) * 100).toFixed(1) + '%';

  const dataC = allData.current;
  const dataP = allData.prev;
  const dataL = allData.last;

  Object.keys(dataC).forEach(area => {
    const col = areaCols[area];
    if (!col) return;
    
    const c = dataC[area];
    const p = dataP[area];
    const l = dataL[area];
    
    // ===============================================
    // ★ 増減（赤青）表示を「する」項目
    // ===============================================
    
    // --- 時給 ---
    sheet.getRange(2, col).setValue(formatRaw(c.wage.areaAvg, 'yen'));
    setDiffValue(sheet.getRange(3, col), c.wage.areaAvg, p.wage.areaAvg, 'yen');
    setDiffValue(sheet.getRange(4, col), c.wage.areaAvg, l.wage.areaAvg, 'yen');
    
    sheet.getRange(5, col).setValue(formatRaw(c.wage.areaAvg, 'yen')); 
    setDiffValue(sheet.getRange(6, col), c.wage.areaAvg, p.wage.areaAvg, 'yen');
    setDiffValue(sheet.getRange(7, col), c.wage.areaAvg, l.wage.areaAvg, 'yen');
    
    // --- 依頼手当総額 ---
    sheet.getRange(8, col).setValue(formatRaw(c.wage.requestAllowance, 'yen'));
    setDiffValue(sheet.getRange(9, col), c.wage.requestAllowance, p.wage.requestAllowance, 'yen');
    setDiffValue(sheet.getRange(10, col), c.wage.requestAllowance, l.wage.requestAllowance, 'yen');

    // --- 売上・来院数 ---
    sheet.getRange(46, col).setValue(c.sales.visitAct > 0 ? `${Number(c.sales.visitAct).toLocaleString()} (${calcRate(c.sales.visitAct, c.sales.visitTgt)})` : '-');
    setDiffValue(sheet.getRange(47, col), c.sales.visitAct, p.sales.visitAct, 'number');
    setDiffValue(sheet.getRange(48, col), c.sales.visitAct, l.sales.visitAct, 'number');
    
    sheet.getRange(49, col).setValue(c.sales.salesAct > 0 ? `${Number(c.sales.salesAct).toLocaleString()} (${calcRate(c.sales.salesAct, c.sales.salesTgt)})` : '-');
    setDiffValue(sheet.getRange(50, col), c.sales.salesAct, p.sales.salesAct, 'yen');
    setDiffValue(sheet.getRange(51, col), c.sales.salesAct, l.sales.salesAct, 'yen');

    // ===============================================
    // ★ 増減表示を「しない」（実数のみ）の項目
    // ===============================================

    // --- 稼働人数（UU頭数） ---
    sheet.getRange(12, col).setValue(formatRaw(c.uu.total, 'people'));
    sheet.getRange(13, col).setValue(formatRaw(p.uu.total, 'people'));
    sheet.getRange(14, col).setValue(formatRaw(l.uu.total, 'people'));

    const nonRegC = c.uu.part + c.uu.dir + c.uu.agc + c.uu.hol;
    const nonRegP = p.uu.part + p.uu.dir + p.uu.agc + p.uu.hol;
    const nonRegL = l.uu.part + l.uu.dir + l.uu.agc + l.uu.hol;
    sheet.getRange(15, col).setValue(formatRaw(nonRegC, 'people'));
    sheet.getRange(16, col).setValue(formatRaw(nonRegP, 'people'));
    sheet.getRange(17, col).setValue(formatRaw(nonRegL, 'people'));

    sheet.getRange(18, col).setValue(formatRaw(c.uu.hol, 'people'));
    sheet.getRange(19, col).setValue(formatRaw(p.uu.hol, 'people'));
    sheet.getRange(20, col).setValue(formatRaw(l.uu.hol, 'people'));

    // --- 構成比（★「改行（\n）」で縦に並べる） ---
    const makeRatioStr = (rt) => {
      if (!rt || rt.total === 0) return '-';
      return `常: ${(rt.reg/rt.total*100).toFixed(1)}%\n定: ${(rt.part/rt.total*100).toFixed(1)}%\n直(ス): ${(rt.dir/rt.total*100).toFixed(1)}%\n紹: ${(rt.agc/rt.total*100).toFixed(1)}%\n休: ${(rt.hol/rt.total*100).toFixed(1)}%`;
    };
    sheet.getRange(21, col).setValue(makeRatioStr(c.ratioTime));
    sheet.getRange(22, col).setValue(makeRatioStr(p.ratioTime));
    sheet.getRange(23, col).setValue(makeRatioStr(l.ratioTime));

    // --- 2診・拠点数 ---
    const twoDocMinsC = Math.round(c.twoDoc.hours * 60);
    const twoDocMinsP = Math.round(p.twoDoc.hours * 60);
    const twoDocMinsL = Math.round(l.twoDoc.hours * 60);
    sheet.getRange(24, col).setValue(formatRaw(twoDocMinsC, 'mins'));
    sheet.getRange(25, col).setValue(formatRaw(twoDocMinsP, 'mins'));
    sheet.getRange(26, col).setValue(formatRaw(twoDocMinsL, 'mins'));

    sheet.getRange(28, col).setValue(formatRaw(c.baseCount.total, 'number'));
    sheet.getRange(29, col).setValue(formatRaw(p.baseCount.total, 'number'));
    sheet.getRange(30, col).setValue(formatRaw(l.baseCount.total, 'number'));

    // --- 採用人数 ---
    sheet.getRange(32, col).setValue(formatRaw(c.hires.direct, 'people'));
    sheet.getRange(33, col).setValue(formatRaw(p.hires.direct, 'people'));
    sheet.getRange(34, col).setValue(formatRaw(l.hires.direct, 'people'));
    
    sheet.getRange(35, col).setValue(formatRaw(c.hires.agency, 'people'));
    sheet.getRange(36, col).setValue(formatRaw(p.hires.agency, 'people'));
    sheet.getRange(37, col).setValue(formatRaw(l.hires.agency, 'people'));

    // --- シフト調整関係 ---
    sheet.getRange(39, col).setValue(formatRaw(c.shiftDiff.absenceMins, 'mins'));
    sheet.getRange(40, col).setValue(formatRaw(p.shiftDiff.absenceMins, 'mins'));
    sheet.getRange(41, col).setValue(formatRaw(l.shiftDiff.absenceMins, 'mins'));

    sheet.getRange(42, col).setValue(formatRaw(c.shiftDiff.overtimeMins, 'mins'));
    sheet.getRange(43, col).setValue(formatRaw(p.shiftDiff.overtimeMins, 'mins'));
    sheet.getRange(44, col).setValue(formatRaw(l.shiftDiff.overtimeMins, 'mins'));

    // --- 在籍数 ---
    sheet.getRange(53, col).setValue(formatRaw(c.enrolled.regular, 'people'));
    sheet.getRange(54, col).setValue(formatRaw(p.enrolled.regular, 'people'));
    sheet.getRange(55, col).setValue(formatRaw(l.enrolled.regular, 'people'));
    
    sheet.getRange(56, col).setValue(formatRaw(c.enrolled.partTime, 'people'));
    sheet.getRange(57, col).setValue(formatRaw(p.enrolled.partTime, 'people'));
    sheet.getRange(58, col).setValue(formatRaw(l.enrolled.partTime, 'people'));
  });
}

// ===============================================
// ★ 裏側（バックアップ用）への保存機能（復活）
// ===============================================
function saveToBackSheet(ss, allData) {
  // 裏側シートの取得（存在しなければ「裏側データ」という名前で自動作成）
  let backSheet = ss.getSheetByName('裏側データ');
  
  if (!backSheet) {
    backSheet = ss.insertSheet('裏側データ');
    backSheet.appendRow(['実行日時', '対象月', '保存データ(JSON)']);
    backSheet.getRange(1, 1, 1, 3).setBackground('#f3f3f3').setFontWeight('bold');
    backSheet.setColumnWidth(3, 800); 
  }

  const timestamp = Utilities.formatDate(new Date(), "GMT+9", "yyyy/MM/dd HH:mm:ss");
  // allData から対象月文字列を取得（なければ "-"）
  const targetMonthStr = (allData && allData.monthStrs) ? allData.monthStrs.current : "-";

  // 最新の集計結果を文字列（JSON）として丸ごとバックアップ保存
  const jsonData = JSON.stringify(allData);

  // 最終行に追記
  backSheet.appendRow([timestamp, targetMonthStr, jsonData]);

  console.log(`✅ 裏側シートに ${targetMonthStr} のデータをバックアップ保存しました。`);
}