// =========================================================
// 4. 月別のデータ抽出（オーケストレーター）
// =========================================================
function extractForMonth(ctx, period) {
  const res = {};
  ctx.targetAreas.forEach(a => {
    res[a] = {
      hires: { direct: 0, agency: 0 },
      sales: { visitAct: 0, visitTgt: 0, salesAct: 0, salesTgt: 0 },
      uuDict: {}, uu: { total: 0, reg: 0, hol: 0, part: 0, agc: 0, dir: 0 }, 
      ratioTime: { total: 0, reg: 0, hol: 0, part: 0, agc: 0, dir: 0 }, 
      wage: { areaAvg: 0, sum: 0, workMins: 0, requestAllowance: 0 },
      shiftDiff: { absenceMins: 0, overtimeMins: 0 },
      twoDoc: { hours: 0, baseCount: 0 },
      _twoDocDaysMap: {}, // ★2診発生日数のカウント用辞書に変更
      baseCount: { total: 0 },
      enrolled: { regular: 0, partTime: 0 }, enrolledRegSet: new Set(), enrolledPartSet: new Set(),
      dailyShifts: {}
    };
  });

  processBaseCounts(ctx, period, res);
  processHiringData(ctx, period, res);
  processSalesData(ctx, period, res);
  processShiftData(ctx, period, res);
  finalizePeriodData(res, ctx.targetAreas);

  return res;
}

// =========================================================
// 5. 個別集計ロジック
// =========================================================
function processBaseCounts(ctx, period, res) {
  const endOfMonth = new Date(period.dateObj.getFullYear(), period.dateObj.getMonth() + 1, 0);
  Object.keys(ctx.clinicAttrs).forEach(cName => {
    const oDate = ctx.clinicAttrs[cName].openDate;
    if (oDate && oDate <= endOfMonth) {
      const areasToPush = ctx.getTargetAreasForClinic(cName);
      areasToPush.forEach(a => { if (res[a]) res[a].baseCount.total++; });
    }
  });
}

function processHiringData(ctx, period, res) {
  if (ctx.chkData.length <= 2) return;
  
  const cleanHead = ctx.chkHead.map(h => h.replace(/\s/g, ''));
  const colPostMonth = cleanHead.findIndex(h => h === '投稿月' || h.includes('投稿月'));
  const colStatus = cleanHead.findIndex(h => h.includes('採用可否'));
  const colSpec = cleanHead.findIndex(h => h.includes('診療科') || h.includes('常勤先での'));
  const colChannel = cleanHead.findIndex(h => h.includes('経緯'));
  const colNameChk = cleanHead.findIndex(h => h.includes('氏名') || h.includes('名前'));
  const colIdChk = cleanHead.findIndex(h => h.includes('医籍番号'));
  const colClinic = cleanHead.findIndex(h => h === '拠点' || (h.includes('拠点') && !h.includes('初回')));
  const colFirstDate = cleanHead.findIndex(h => h.includes('初回勤務予定日'));
  const colFirstClinic = cleanHead.findIndex(h => h.includes('初回勤務拠点'));

  if (colPostMonth === -1) return;

  for (let r = 2; r < ctx.chkData.length; r++) {
    const rawPostMonth = ctx.chkData[r][colPostMonth];
    let dStr = '';
    if (rawPostMonth instanceof Date) {
      dStr = Utilities.formatDate(rawPostMonth, "GMT+9", "yyyy/MM");
    } else if (String(rawPostMonth).trim() !== '') {
      let pmStr = String(rawPostMonth).trim().replace(/\./g, '/');
      const match = pmStr.match(/^(\d{4})\/(\d{1,2})/);
      if (match) dStr = `${match[1]}/${match[2].padStart(2, '0')}`;
      else dStr = pmStr.substring(0, 7);
    }
    if (dStr !== period.str) continue;

    const status = colStatus !== -1 ? String(ctx.chkData[r][colStatus]) : '';
    const spec = colSpec !== -1 ? String(ctx.chkData[r][colSpec]) : '';
    const rawFirstDate = colFirstDate !== -1 ? ctx.chkData[r][colFirstDate] : '';

    const hasFirstDate = (rawFirstDate instanceof Date) || (String(rawFirstDate).trim() !== '');
    const isHired = status.includes('採用') || hasFirstDate;
    const isPed = spec.includes('小児科');

    if (isHired && isPed) {
      const name = colNameChk !== -1 ? ctx.cleanStr(ctx.chkData[r][colNameChk]) : '';
      const id = colIdChk !== -1 ? ctx.cleanId(ctx.chkData[r][colIdChk]) : '';
      const channel = colChannel !== -1 ? String(ctx.chkData[r][colChannel]).trim() : '';
      
      let rawClinic = colClinic !== -1 ? String(ctx.chkData[r][colClinic]).trim() : '';
      const firstClinicVal = colFirstClinic !== -1 ? String(ctx.chkData[r][colFirstClinic]).trim() : '';

      if (!rawClinic && firstClinicVal) rawClinic = firstClinicVal;
      if (!rawClinic) {
        const shiftInfo = ctx.firstShiftMap[id] || ctx.firstShiftMap[name];
        if (shiftInfo) rawClinic = shiftInfo.clinic;
      }

      const areasToPush = ctx.getTargetAreasForClinic(rawClinic);
      areasToPush.forEach(a => {
        if (res[a]) {
          if (channel.includes('直接')) res[a].hires.direct++;
          else if (channel.match(/エムスリー|民間医局|エムステージ|Mステージ|マイナビ|MRT/)) res[a].hires.agency++;
        }
      });
    }
  }
}

function processSalesData(ctx, period, res) {
  try {
    const q = `title contains '業績管理シート' and title contains '全拠点' and title contains '${period.dot}' and mimeType = '${MimeType.GOOGLE_SHEETS}'`;
    const files = DriveApp.searchFiles(q);
    if (!files.hasNext()) return;
    
    const sData = ctx.getSsByIdCached(files.next().getId()).getSheetByName('来院数').getDataRange().getValues();
    let totalColIndex = -1;
    for (let c = 4; c < sData[2].length; c++) { if (String(sData[2][c]).includes('合計') || String(sData[3][c]).includes('合計')) { totalColIndex = c; break; } }
    
    if (totalColIndex !== -1) {
      const processed = { visitAct: new Set(), visitTgt: new Set(), salesAct: new Set(), salesTgt: new Set() };
      let currentMode = null;
      
      for (let r = 0; r < sData.length; r++) {
        const combinedVal = ctx.cleanStr(sData[r][0]) + ctx.cleanStr(sData[r][1]);
        if (combinedVal.includes('来院数：実績') || combinedVal.includes('来院数:実績')) { currentMode = 'visitAct'; continue; }
        if (combinedVal.includes('来院数目標') || combinedVal.includes('来院数：目標') || combinedVal.includes('来院数：予算')) { currentMode = 'visitTgt'; continue; }
        if (combinedVal.includes('売上：実績') || combinedVal.includes('売上:実績')) { currentMode = 'salesAct'; continue; }
        if (combinedVal.includes('売上：目標') || combinedVal.includes('売上：予算')) { currentMode = 'salesTgt'; continue; }
        if (combinedVal.includes('目標') || combinedVal.includes('実績')) { currentMode = null; continue; }

        if (currentMode) {
          const rawClinic = String(sData[r][1]).trim();
          if (!rawClinic || rawClinic === 'OK' || rawClinic.includes('計')) continue;
          const offClinic = ctx.clinicDict[rawClinic] || rawClinic;

          if (!processed[currentMode].has(offClinic)) {
             const amt = Number(String(sData[r][totalColIndex]).replace(/,/g, ''));
             if (!isNaN(amt)) {
               const areasToPush = ctx.getTargetAreasForClinic(offClinic);
               areasToPush.forEach(a => { if (res[a]) res[a].sales[currentMode] += amt; });
               processed[currentMode].add(offClinic);
             }
          }
        }
      }
    }
  } catch(e) {}
}

function processShiftData(ctx, period, res) {
  let shiftSheet = null;
  const arcInfo = ctx.getDsInfo(`${period.year}確定シフト`);
  if (arcInfo && arcInfo.url) {
    try {
      for (const s of ctx.getSheetsCached(arcInfo.url)) { 
        if (s.getName().match(new RegExp(`^0?${period.monthNum}月?`)) || s.getName().includes(period.str)) { shiftSheet = s; break; } 
      }
    } catch(e) {}
  }

  if (!shiftSheet) {
    const reqInfo = ctx.getDsInfo('請求書くん');
    if (reqInfo && reqInfo.url) {
      try {
        for (const s of ctx.getSheetsCached(reqInfo.url)) { 
          if (s.getName().match(new RegExp(`^0?${period.monthNum}月?`)) || s.getName().includes(period.str)) { shiftSheet = s; break; } 
        }
        if (!shiftSheet) shiftSheet = ctx.getSsByUrl(reqInfo.url).ss.getSheetByName('確定シフト') || ctx.getSheetsCached(reqInfo.url)[0];
      } catch(e) {}
    }
  }

  if (!shiftSheet) return;

  const sData = shiftSheet.getDataRange().getValues();
  const headers = sData[0].map(String);
  const colClinic = headers.findIndex(h => h.includes('クリニック名'));
  const colDepart = headers.findIndex(h => h.includes('診療科'));
  const colDate = headers.findIndex(h => h.includes('勤務日'));
  const colStart = headers.findIndex(h => h.includes('勤務開始時間') && !h.includes('元のシフト'));
  const colEnd = headers.findIndex(h => h.includes('勤務終了時間') && !h.includes('元のシフト'));
  const colId = headers.findIndex(h => h.includes('医籍番号'));
  const colName = headers.findIndex(h => h.includes('氏名') || h.includes('名前'));
  const colWageTotal = headers.findIndex(h => h.includes('時給合計'));
  const colComments = [];
  headers.forEach((h, idx) => { if (h.includes('スタッフコメント')) colComments.push(idx); });

  const allowanceCols = [];
  for (let num = 1; num <= 5; num++) {
    const itemCol = headers.findIndex(h => h === `追加支給項目${num}`);
    const amountCol = headers.findIndex(h => h === `追加支給額${num}`);
    if (itemCol !== -1 && amountCol !== -1) allowanceCols.push({ item: itemCol, amount: amountCol });
  }

  for (let i = 1; i < sData.length; i++) {
    const rawClinic = colClinic !== -1 ? String(sData[i][colClinic]).trim() : '';
    const officialClinic = ctx.clinicDict[rawClinic];
    if (!officialClinic) continue; 
    
    // 亀有・北葛西は小児科限定
    if (colDepart !== -1) {
      const depart = String(sData[i][colDepart]).trim();
      if (officialClinic.includes('亀有') || officialClinic.includes('北葛西')) {
        if (depart !== '小児科') continue;
      }
    }

    const dateObj = ctx.parseDateSafe(sData[i][colDate]);
    if (!dateObj || Utilities.formatDate(dateObj, "GMT+9", "yyyy/MM") !== period.str) continue;
    
    const startMin = ctx.timeToMins(sData[i][colStart]);
    const endMin = ctx.timeToMins(sData[i][colEnd]);
    if (startMin === null || endMin === null) continue;
    
    const shiftDurMins = endMin - startMin;
    if (shiftDurMins < 30) continue; 
    
    const id = colId !== -1 ? ctx.cleanId(sData[i][colId]) : '';
    const name = colName !== -1 ? ctx.cleanStr(sData[i][colName]) : '';
    const totalWage = colWageTotal !== -1 ? Number(String(sData[i][colWageTotal]).replace(/,/g, '')) : 0;
    
    let requestAllowance = 0;
    allowanceCols.forEach(cols => {
      const itemVal = String(sData[i][cols.item] || '');
      if (itemVal.includes('依頼手当')) {
        const amountVal = Number(String(sData[i][cols.amount]).replace(/[^\d.-]/g, ''));
        if (!isNaN(amountVal) && amountVal > 0) requestAllowance += amountVal;
      }
    });
    
    let empType = 'スポット';
    let route = '';
    let isHol = false;
    colComments.forEach(cIdx => {
      const comment = String(sData[i][cIdx] || '');
      if (comment.match(/所定休出|所定外勤務|所定外休出/)) isHol = true;
      if (comment.match(/紹介|エムスリー|民間医局|エムステージ|Mステージ|マイナビ|MRT/)) route = '紹介会社';
    });

    const docInfos = [...(ctx.docInfoDict[id] || []), ...(ctx.docInfoDict[name] || [])];
    docInfos.forEach(info => {
      if (info.route) route = info.route; 
      if (info.type) {
        if (info.masterYear && info.masterYear !== period.year) return;
        let isWithinPeriod = true;
        if (info.joinDate && dateObj < info.joinDate) isWithinPeriod = false;
        if (info.leaveDate && dateObj > info.leaveDate) isWithinPeriod = false;
        if (isWithinPeriod) {
          empType = info.type;
          if (!info.hasContract) isHol = true; 
        }
      }
    });

    let finalAttr = 'dir';
    if (empType === '常勤') finalAttr = isHol ? 'hol' : 'reg';
    else if (empType === '定期非常勤') finalAttr = isHol ? 'dir' : 'part'; 
    else finalAttr = route === '紹介会社' ? 'agc' : 'dir';

    const dateStr = Utilities.formatDate(dateObj, "GMT+9", "yyyy/MM/dd");
    const key = `${dateStr}_${officialClinic}`;
    const areasToPush = ctx.getTargetAreasForClinic(officialClinic);

    areasToPush.forEach(a => {
      const d = res[a];
      if (!d) return;

      if (!isNaN(totalWage) && totalWage > 0) { d.wage.sum += totalWage; d.wage.workMins += shiftDurMins; }
      if (requestAllowance > 0) d.wage.requestAllowance += requestAllowance;
      
      d.ratioTime.total += shiftDurMins;
      d.ratioTime[finalAttr] += shiftDurMins;

      const uuKey = id || name || `row_${i}`;
      if (!d.uuDict[uuKey]) {
        d.uuDict[uuKey] = { type: finalAttr };
      } else {
        const priority = { 'hol': 5, 'reg': 4, 'part': 3, 'agc': 2, 'dir': 1 };
        if (priority[finalAttr] > priority[d.uuDict[uuKey].type]) d.uuDict[uuKey].type = finalAttr;
      }

      if (empType === '常勤' && !isHol) d.enrolledRegSet.add(uuKey);
      if (empType === '定期非常勤' && !isHol) d.enrolledPartSet.add(uuKey);

      if (!d.dailyShifts[key]) d.dailyShifts[key] = { clinic: officialClinic, dateStr: dateStr, dateObj: dateObj, shifts: [] };
      d.dailyShifts[key].shifts.push({ start: startMin, end: endMin });
    });
  }

  ctx.targetAreas.forEach(a => {
    const d = res[a];
    for (const key in d.dailyShifts) {
      const info = d.dailyShifts[key];
      const attr = ctx.clinicAttrs[info.clinic];
      if (attr && attr.openDate && info.dateObj < attr.openDate) continue; 
      
      let eveningClose = (info.clinic.includes('北葛西')) ? 20 * 60 : 21 * 60;
      let expectedPeriods = [{ start: 9 * 60, end: 13 * 60 }, { start: 15 * 60, end: eveningClose }];
      
      const dayClosures = ctx.closures[info.dateStr] || [];
      const applicableClosures = dayClosures.filter(c => c.clinic === '全拠点' || c.clinic === info.clinic);
      
      applicableClosures.forEach(c => {
        let cStart, cEnd;
        if (c.type.includes('全日')) { cStart = 9*60; cEnd = 21*60; }
        else if (c.type.includes('午後夜間')) { cStart = 15*60; cEnd = 21*60; }
        else if (c.type.includes('午前')) { cStart = 9*60; cEnd = 13*60; }
        else if (c.type.includes('夜間')) { cStart = 18*60; cEnd = 21*60; }
        else if (c.type.includes('午後')) { cStart = 15*60; cEnd = 18*60; }
        else return;

        let newExpected = [];
        for (let ep of expectedPeriods) {
          if (cEnd <= ep.start || cStart >= ep.end) newExpected.push(ep);
          else {
            if (cStart > ep.start) newExpected.push({ start: ep.start, end: cStart });
            if (cEnd < ep.end) newExpected.push({ start: cEnd, end: ep.end });
          }
        }
        expectedPeriods = newExpected;
      });
      
      const shifts = info.shifts;
      let totalDoctorMins = 0; 
      let dailyOvertimeMins = 0;
      const intervals = [];

      for (const s of shifts) {
        totalDoctorMins += (s.end - s.start);
        intervals.push([s.start, s.end]);
        let ot = 0;
        if (s.end > 13 * 60 && s.end < 15 * 60) ot = s.end - (13 * 60);
        else if (s.end > eveningClose) ot = s.end - eveningClose; 
        dailyOvertimeMins += ot;
      }
      
      intervals.sort((a, b) => a[0] - b[0]);
      const merged = [intervals[0]];
      for (let i = 1; i < intervals.length; i++) {
        const last = merged[merged.length - 1];
        const curr = intervals[i];
        if (curr[0] <= last[1]) last[1] = Math.max(last[1], curr[1]);
        else merged.push(curr);
      }
      
      let uniqueCoveredMins = 0;
      for (const m of merged) uniqueCoveredMins += (m[1] - m[0]);
      const twoDoctorMins = totalDoctorMins - uniqueCoveredMins;
      
      let absenceMins = 0;
      for (const ep of expectedPeriods) {
         let coveredInThisPeriod = 0;
         for (const m of merged) {
            const overlapStart = Math.max(ep.start, m[0]);
            const overlapEnd = Math.min(ep.end, m[1]);
            if (overlapStart < overlapEnd) coveredInThisPeriod += (overlapEnd - overlapStart);
         }
         absenceMins += ((ep.end - ep.start) - coveredInThisPeriod);
      }
      
      d.shiftDiff.overtimeMins += dailyOvertimeMins;
      d.twoDoc.hours += (twoDoctorMins / 60); 
      d.shiftDiff.absenceMins += absenceMins;

      // ★ 修正：1日の重複時間（2診時間）が180分以上の場合、その拠点の発生日数を+1する
      if (twoDoctorMins >= 180) {
        d._twoDocDaysMap[info.clinic] = (d._twoDocDaysMap[info.clinic] || 0) + 1;
      }
    }
  });
}

function finalizePeriodData(res, targetAreas) {
  targetAreas.forEach(a => {
    const d = res[a];
    Object.values(d.uuDict).forEach(u => {
      d.uu.total++;
      d.uu[u.type]++;
    });
    d.wage.areaAvg = (d.wage.sum > 0 && d.wage.workMins > 0) ? Math.round(d.wage.sum / (d.wage.workMins / 60)) : 0;
    
    d.enrolled.regular = d.enrolledRegSet.size;
    d.enrolled.partTime = d.enrolledPartSet.size;
    
    // ★ 修正：月に4日以上（週1回ペース）発生している拠点のみをカウント
    d.twoDoc.baseCount = Object.values(d._twoDocDaysMap).filter(days => days >= 4).length;
    
    delete d.enrolledRegSet; 
    delete d.enrolledPartSet;
    delete d._twoDocDaysMap; // ★使い終わった辞書を削除
    delete d.uuDict; 
    delete d.dailyShifts; 
    delete d.wage.sum; 
    delete d.wage.workMins;
  });
}