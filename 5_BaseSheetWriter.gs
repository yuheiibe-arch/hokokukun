// =========================================================
// 7. ベースシートへの書き込み処理（誤爆防止・実数表示の完全版）
// =========================================================
function writeToBaseSheet(ss, allData) {
  const sheet = ss.getSheetByName('ベースシート');
  
  const abColValues = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  const searchTexts = abColValues.map((row, index) => ({
    rowNum: index + 1,
    text: (String(row[0]) + String(row[1])).replace(/\s+/g, '')
  }));

  const searchTargets = [
    { key: 'エリア平均時給', id: 'wageAvg', exclude: [] },
    { key: 'エリア別詳細時給', id: 'wageDetail', exclude: [] },
    { key: '依頼手当', id: 'allowance', exclude: [] },
    { key: '稼働人員（全医師', id: 'uuTotal', exclude: [] },
    { key: '稼働人員(全医師', id: 'uuTotal', exclude: [] },
    { key: '稼働人員（常勤除く', id: 'uuNonReg', exclude: [] },
    { key: '稼働人員(常勤除く', id: 'uuNonReg', exclude: [] },
    { key: '所定休出医師数', id: 'uuHol', exclude: [] },
    { key: '構成比', id: 'ratio', exclude: [] },
    { key: '医師勤務時間', id: 'workTimeTotal', exclude: ['固定', '募集'] }, 
    { key: '固定シフト', id: 'workTimeFixed', exclude: [] }, 
    { key: '募集シフト', id: 'workTimeRecruit', exclude: [] }, 
    { key: '２診時間', id: 'twoDoc', exclude: [] },
    { key: '2診拠点数', id: 'twoDocBase', exclude: [] },
    { key: '２診拠点数', id: 'twoDocBase', exclude: [] },
    // ★修正：2診拠点数の行に「対象拠点数」の文字が残っていても絶対に誤爆しないようバリアを追加
    { key: '対象拠点数', id: 'baseCount', exclude: ['2診', '２診'] }, 
    { key: '新規採用', id: 'hireDirect', exclude: ['紹介', 'エージェント'] },
    { key: '新規採用', id: 'hireAgency', exclude: ['直接'] },
    { key: '不在時間', id: 'absence', exclude: [] },
    { key: '残業時間', id: 'overtime', exclude: [] },
    { key: '来院数', id: 'visit', exclude: [] },
    { key: '売上（目標達成率）', id: 'sales', exclude: [] }, 
    { key: '在籍医師数（常勤', id: 'enrollReg', exclude: [] },
    { key: '在籍医師数（定期', id: 'enrollPart', exclude: [] }
  ];

  const rowMap = {};
  searchTargets.forEach(target => {
    if (rowMap[target.id]) return; 
    const found = searchTexts.find(item => {
      const hasKey = item.text.includes(target.key);
      const hasNoExclude = target.exclude.every(exKw => !item.text.includes(exKw));
      return hasKey && hasNoExclude;
    });
    if (found) rowMap[target.id] = found.rowNum;
  });

  Object.values(rowMap).forEach(r => {
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
  
  sheet.getRange(2, 4, sheet.getLastRow() - 1, lastCol - 3).clearContent().setBackground('#ffffff').setFontColor('#000000').setHorizontalAlignment("left");

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
      color = '#ff0000';
    } else if (diff < 0) {
      diffStr = ` (-${formatRaw(Math.abs(diff), type)})`;
      color = '#0000ff';
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
    
    if (rowMap.wageAvg) {
      sheet.getRange(rowMap.wageAvg, col).setValue(formatRaw(c.wage.areaAvg, 'yen'));
      setDiffValue(sheet.getRange(rowMap.wageAvg + 1, col), c.wage.areaAvg, p.wage.areaAvg, 'yen');
      setDiffValue(sheet.getRange(rowMap.wageAvg + 2, col), c.wage.areaAvg, l.wage.areaAvg, 'yen');
    }
    if (rowMap.wageDetail) { 
      sheet.getRange(rowMap.wageDetail, col).setValue(formatRaw(c.wage.areaAvg, 'yen')); 
      setDiffValue(sheet.getRange(rowMap.wageDetail + 1, col), c.wage.areaAvg, p.wage.areaAvg, 'yen');
      setDiffValue(sheet.getRange(rowMap.wageDetail + 2, col), c.wage.areaAvg, l.wage.areaAvg, 'yen');
    }
    if (rowMap.allowance) {
      sheet.getRange(rowMap.allowance, col).setValue(formatRaw(c.wage.requestAllowance, 'yen'));
      setDiffValue(sheet.getRange(rowMap.allowance + 1, col), c.wage.requestAllowance, p.wage.requestAllowance, 'yen');
      setDiffValue(sheet.getRange(rowMap.allowance + 2, col), c.wage.requestAllowance, l.wage.requestAllowance, 'yen');
    }
    if (rowMap.visit) {
      sheet.getRange(rowMap.visit, col).setValue(c.sales.visitAct > 0 ? `${Number(c.sales.visitAct).toLocaleString()} (${calcRate(c.sales.visitAct, c.sales.visitTgt)})` : '-');
      setDiffValue(sheet.getRange(rowMap.visit + 1, col), c.sales.visitAct, p.sales.visitAct, 'number');
      setDiffValue(sheet.getRange(rowMap.visit + 2, col), c.sales.visitAct, l.sales.visitAct, 'number');
    }
    if (rowMap.sales) {
      sheet.getRange(rowMap.sales, col).setValue(c.sales.salesAct > 0 ? `${Number(c.sales.salesAct).toLocaleString()} (${calcRate(c.sales.salesAct, c.sales.salesTgt)})` : '-');
      setDiffValue(sheet.getRange(rowMap.sales + 1, col), c.sales.salesAct, p.sales.salesAct, 'yen');
      setDiffValue(sheet.getRange(rowMap.sales + 2, col), c.sales.salesAct, l.sales.salesAct, 'yen');
    }

    if (rowMap.uuTotal) {
      sheet.getRange(rowMap.uuTotal, col).setValue(formatRaw(c.uu.total, 'people'));
      sheet.getRange(rowMap.uuTotal + 1, col).setValue(formatRaw(p.uu.total, 'people'));
      sheet.getRange(rowMap.uuTotal + 2, col).setValue(formatRaw(l.uu.total, 'people'));
    }
    if (rowMap.uuNonReg) {
      const getNonReg = d => d.uu.part + d.uu.dir + d.uu.agc + d.uu.hol;
      sheet.getRange(rowMap.uuNonReg, col).setValue(formatRaw(getNonReg(c), 'people'));
      sheet.getRange(rowMap.uuNonReg + 1, col).setValue(formatRaw(getNonReg(p), 'people'));
      sheet.getRange(rowMap.uuNonReg + 2, col).setValue(formatRaw(getNonReg(l), 'people'));
    }
    if (rowMap.uuHol) {
      sheet.getRange(rowMap.uuHol, col).setValue(formatRaw(c.uu.hol, 'people'));
      sheet.getRange(rowMap.uuHol + 1, col).setValue(formatRaw(p.uu.hol, 'people'));
      sheet.getRange(rowMap.uuHol + 2, col).setValue(formatRaw(l.uu.hol, 'people'));
    }
    if (rowMap.ratio) {
      const makeRatioStr = (rt) => {
        if (!rt || rt.total === 0) return '-';
        return `常: ${(rt.reg/rt.total*100).toFixed(1)}%\n定: ${(rt.part/rt.total*100).toFixed(1)}%\n直(ス): ${(rt.dir/rt.total*100).toFixed(1)}%\n紹: ${(rt.agc/rt.total*100).toFixed(1)}%\n休: ${(rt.hol/rt.total*100).toFixed(1)}%`;
      };
      sheet.getRange(rowMap.ratio, col).setValue(makeRatioStr(c.ratioTime));
      sheet.getRange(rowMap.ratio + 1, col).setValue(makeRatioStr(p.ratioTime));
      sheet.getRange(rowMap.ratio + 2, col).setValue(makeRatioStr(l.ratioTime));
    }

    if (rowMap.workTimeTotal) {
      sheet.getRange(rowMap.workTimeTotal, col).setValue(formatRaw(c.ratioTime.total, 'mins'));
      sheet.getRange(rowMap.workTimeTotal + 1, col).setValue(formatRaw(p.ratioTime.total, 'mins'));
      sheet.getRange(rowMap.workTimeTotal + 2, col).setValue(formatRaw(l.ratioTime.total, 'mins'));
    }
    if (rowMap.workTimeFixed) {
      const getFixed = d => d.ratioTime.reg + d.ratioTime.part;
      sheet.getRange(rowMap.workTimeFixed, col).setValue(formatRaw(getFixed(c), 'mins'));
      sheet.getRange(rowMap.workTimeFixed + 1, col).setValue(formatRaw(getFixed(p), 'mins'));
      sheet.getRange(rowMap.workTimeFixed + 2, col).setValue(formatRaw(getFixed(l), 'mins'));
    }
    if (rowMap.workTimeRecruit) {
      const getRecruit = d => d.ratioTime.dir + d.ratioTime.agc + d.ratioTime.hol;
      sheet.getRange(rowMap.workTimeRecruit, col).setValue(formatRaw(getRecruit(c), 'mins'));
      sheet.getRange(rowMap.workTimeRecruit + 1, col).setValue(formatRaw(getRecruit(p), 'mins'));
      sheet.getRange(rowMap.workTimeRecruit + 2, col).setValue(formatRaw(getRecruit(l), 'mins'));
    }

    if (rowMap.twoDoc) {
      sheet.getRange(rowMap.twoDoc, col).setValue(formatRaw(Math.round(c.twoDoc.hours * 60), 'mins'));
      sheet.getRange(rowMap.twoDoc + 1, col).setValue(formatRaw(Math.round(p.twoDoc.hours * 60), 'mins'));
      sheet.getRange(rowMap.twoDoc + 2, col).setValue(formatRaw(Math.round(l.twoDoc.hours * 60), 'mins'));
    }
    
    // ★ 修正：2診拠点数はプラスマイナス（増減）を表示せず、setValueで実数のみを出力する
    if (rowMap.twoDocBase) {
      sheet.getRange(rowMap.twoDocBase, col).setValue(formatRaw(c.twoDoc.baseCount, 'number'));
      sheet.getRange(rowMap.twoDocBase + 1, col).setValue(formatRaw(p.twoDoc.baseCount, 'number'));
      sheet.getRange(rowMap.twoDocBase + 2, col).setValue(formatRaw(l.twoDoc.baseCount, 'number'));
    }
    
    if (rowMap.baseCount) {
      sheet.getRange(rowMap.baseCount, col).setValue(formatRaw(c.baseCount.total, 'number'));
      sheet.getRange(rowMap.baseCount + 1, col).setValue(formatRaw(p.baseCount.total, 'number'));
      sheet.getRange(rowMap.baseCount + 2, col).setValue(formatRaw(l.baseCount.total, 'number'));
    }
    if (rowMap.hireDirect) {
      sheet.getRange(rowMap.hireDirect, col).setValue(formatRaw(c.hires.direct, 'people'));
      sheet.getRange(rowMap.hireDirect + 1, col).setValue(formatRaw(p.hires.direct, 'people'));
      sheet.getRange(rowMap.hireDirect + 2, col).setValue(formatRaw(l.hires.direct, 'people'));
    }
    if (rowMap.hireAgency) {
      sheet.getRange(rowMap.hireAgency, col).setValue(formatRaw(c.hires.agency, 'people'));
      sheet.getRange(rowMap.hireAgency + 1, col).setValue(formatRaw(p.hires.agency, 'people'));
      sheet.getRange(rowMap.hireAgency + 2, col).setValue(formatRaw(l.hires.agency, 'people'));
    }
    if (rowMap.absence) {
      sheet.getRange(rowMap.absence, col).setValue(formatRaw(c.shiftDiff.absenceMins, 'mins'));
      sheet.getRange(rowMap.absence + 1, col).setValue(formatRaw(p.shiftDiff.absenceMins, 'mins'));
      sheet.getRange(rowMap.absence + 2, col).setValue(formatRaw(l.shiftDiff.absenceMins, 'mins'));
    }
    if (rowMap.overtime) {
      sheet.getRange(rowMap.overtime, col).setValue(formatRaw(c.shiftDiff.overtimeMins, 'mins'));
      sheet.getRange(rowMap.overtime + 1, col).setValue(formatRaw(p.shiftDiff.overtimeMins, 'mins'));
      sheet.getRange(rowMap.overtime + 2, col).setValue(formatRaw(l.shiftDiff.overtimeMins, 'mins'));
    }
    if (rowMap.enrollReg) {
      sheet.getRange(rowMap.enrollReg, col).setValue(formatRaw(c.enrolled.regular, 'people'));
      sheet.getRange(rowMap.enrollReg + 1, col).setValue(formatRaw(p.enrolled.regular, 'people'));
      sheet.getRange(rowMap.enrollReg + 2, col).setValue(formatRaw(l.enrolled.regular, 'people'));
    }
    if (rowMap.enrollPart) {
      sheet.getRange(rowMap.enrollPart, col).setValue(formatRaw(c.enrolled.partTime, 'people'));
      sheet.getRange(rowMap.enrollPart + 1, col).setValue(formatRaw(p.enrolled.partTime, 'people'));
      sheet.getRange(rowMap.enrollPart + 2, col).setValue(formatRaw(l.enrolled.partTime, 'people'));
    }
  });
}

// =========================================================
// 6. 月次アーカイブへの保存（2診拠点数対応版）
// =========================================================
function saveToBackSheet(ss, monthData, monthStr) {
  if (!monthData || !monthStr || monthStr === "-") return;
  const sheet = ss.getSheetByName('月次アーカイブ');
  if (!sheet) return;

  const headers = sheet.getDataRange().getValues()[0].map(String).map(h => h.replace(/\s+/g, ''));
  const findCol = (kw) => headers.findIndex(h => h.includes(kw));

  const cMonth = findCol('対象月');
  const cArea = findCol('エリア');
  if (cMonth === -1 || cArea === -1) return; 

  const areaKeys = Object.keys(monthData);
  const newRows = [];
  const timestamp = Utilities.formatDate(new Date(), "GMT+9", "yyyy/MM/dd HH:mm:ss");

  areaKeys.forEach(area => {
    const d = monthData[area];
    const rt = d.ratioTime;
    const row = new Array(headers.length).fill('');

    row[findCol('実行日時')] = timestamp;
    row[cMonth] = new Date(monthStr + '/01');
    row[cArea] = area;

    const setVal = (kw, val) => {
      const idx = findCol(kw);
      if (idx !== -1) row[idx] = val;
    };

    setVal('来院数実績', d.sales.visitAct);
    setVal('売上実績', d.sales.salesAct);
    setVal('平均時給', d.wage.areaAvg);
    setVal('依頼手当', d.wage.requestAllowance);
    setVal('稼働UU_総数', d.uu.total);
    setVal('稼働UU_常勤', d.uu.reg);
    setVal('稼働UU_定期', d.uu.part);
    setVal('稼働UU_直接', d.uu.dir);
    setVal('稼働UU_紹介', d.uu.agc);
    setVal('稼働UU_休出', d.uu.hol);
    setVal('在籍数_常勤', d.enrolled.regular);
    setVal('在籍数_定期', d.enrolled.partTime);
    setVal('稼働拠点数', d.baseCount.total);

    setVal('稼働時間_常勤', rt.reg / 60);
    setVal('稼働時間_定期', rt.part / 60);
    setVal('稼働時間_直接', rt.dir / 60);
    setVal('稼働時間_紹介', rt.agc / 60);

    setVal('常勤比率', rt.total > 0 ? rt.reg / rt.total : 0);
    setVal('定期非常勤比率', rt.total > 0 ? rt.part / rt.total : 0);
    setVal('直応募比率', rt.total > 0 ? rt.dir / rt.total : 0);
    setVal('紹介会社比率', rt.total > 0 ? rt.agc / rt.total : 0);

    setVal('2診時間', d.twoDoc.hours);
    setVal('2診拠点数', d.twoDoc.baseCount);
    setVal('不在時間', d.shiftDiff.absenceMins / 60);
    setVal('残業時間', d.shiftDiff.overtimeMins / 60);

    setVal('医師勤務時間', rt.total / 60);
    setVal('固定シフト時間', (rt.reg + rt.part) / 60);
    setVal('募集シフト時間', (rt.dir + rt.agc + rt.hol) / 60);

    setVal('新規採用_直接', d.hires.direct);
    setVal('新規採用_紹介', d.hires.agency);

    newRows.push(row);
  });

  const allData = sheet.getDataRange().getValues();
  for (let r = allData.length - 1; r > 0; r--) {
    const rowMonth = allData[r][cMonth];
    if (!rowMonth) continue;
    
    let rowMonthStr = "";
    if (rowMonth instanceof Date) {
      rowMonthStr = Utilities.formatDate(rowMonth, "GMT+9", "yyyy/MM");
    } else {
      rowMonthStr = String(rowMonth).replace('実績', '').trim();
    }

    if (rowMonthStr === monthStr) {
      sheet.deleteRow(r + 1);
    }
  }

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
  }
}