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
// ★ バックシート（月次アーカイブ）への保存機能
// ===============================================
function saveToBackSheet(ss, allDataOrCurrent, monthStrOptional) {
  // 呼び出し元の引数の違いを吸収する処理
  let currentData, targetMonthStr;
  if (monthStrOptional) {
    currentData = allDataOrCurrent;
    targetMonthStr = monthStrOptional;
  } else {
    currentData = allDataOrCurrent.current;
    targetMonthStr = allDataOrCurrent.monthStrs.current;
  }

  if (!currentData || !targetMonthStr || targetMonthStr === "-") return;

  // シートの取得（存在しなければ「月次アーカイブ」という名前で自動作成）
  let backSheet = ss.getSheetByName('月次アーカイブ');
  const headers = [
    '実行日時', '対象月', 'エリア', 
    '平均時給(円)', '依頼手当(円)', 
    '新規採用_直接(人)', '新規採用_紹介(人)', 
    '在籍数_常勤(人)', '在籍数_定期(人)',
    '稼働UU_総数(人)', '稼働UU_常勤(人)', '稼働UU_定期(人)', '稼働UU_直接(人)', '稼働UU_紹介(人)', '稼働UU_休出(人)', 
    '不在時間(分)', '残業時間(分)', '2診時間(h)', 
    '稼働拠点数', '来院数実績', '売上実績'
  ];

  if (!backSheet) {
    backSheet = ss.insertSheet('月次アーカイブ');
    backSheet.appendRow(headers);
    backSheet.getRange(1, 1, 1, headers.length).setBackground('#e3f2fd').setFontWeight('bold');
    backSheet.setFrozenRows(1);
  }

  // ★重複防止：すでに同じ対象月のデータがあれば行を削除（最新データで上書きするため）
  const data = backSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === targetMonthStr) {
      backSheet.deleteRow(i + 1);
    }
  }

  // 書き込むデータをエリアごとに配列として作成
  const timestamp = Utilities.formatDate(new Date(), "GMT+9", "yyyy/MM/dd HH:mm:ss");
  const targetAreas = ['関東', '関西', '関東第一', '関東第二', '埼玉', '神奈川', '千葉', '茨城', '大阪', 'グループ全体'];
  const rowsToAppend = [];

  targetAreas.forEach(area => {
    const d = currentData[area];
    if (!d) return;

    const row = [
      timestamp,                    // 実行日時
      targetMonthStr,               // 対象月
      area,                         // エリア
      d.wage.areaAvg || 0,          // 平均時給
      d.wage.requestAllowance || 0, // 依頼手当
      d.hires.direct || 0,          // 新規採用(直接)
      d.hires.agency || 0,          // 新規採用(紹介)
      d.enrolled.regular || 0,      // 在籍数(常勤)
      d.enrolled.partTime || 0,     // 在籍数(定期)
      d.uu.total || 0,              // 稼働UU(総数)
      d.uu.reg || 0,                // 稼働UU(常勤)
      d.uu.part || 0,               // 稼働UU(定期)
      d.uu.dir || 0,                // 稼働UU(直接)
      d.uu.agc || 0,                // 稼働UU(紹介)
      d.uu.hol || 0,                // 稼働UU(休出)
      d.shiftDiff.absenceMins || 0, // 不在時間
      d.shiftDiff.overtimeMins || 0,// 残業時間
      d.twoDoc.hours || 0,          // 2診時間
      d.baseCount.total || 0,       // 稼働拠点数
      d.sales.visitAct || 0,        // 来院数実績
      d.sales.salesAct || 0         // 売上実績
    ];
    rowsToAppend.push(row);
  });

  // 一括でシートへ追記
  if (rowsToAppend.length > 0) {
    backSheet.getRange(backSheet.getLastRow() + 1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
  }

  console.log(`✅ 月次アーカイブシートに ${targetMonthStr} のデータを表形式で保存しました。`);
}