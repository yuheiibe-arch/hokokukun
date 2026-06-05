function extractAllData(targetDate) {
  const ssId = '1zZaoAZ1q76hq_Dcpg0XGTzZFHQXrFU9E1GqYHeE7F2s';
  const ss = SpreadsheetApp.openById(ssId);
  
  // ==========================================
  // 🚀 高速化エンジン：スプレッドシートの記憶（キャッシュ）
  // ==========================================
  const ssCache = {};
  const getSsByUrl = (url) => {
    if (!ssCache[url]) ssCache[url] = { ss: SpreadsheetApp.openByUrl(url), sheets: null };
    return ssCache[url];
  };
  const getSheetsCached = (url) => {
    const cacheObj = getSsByUrl(url);
    if (!cacheObj.sheets) cacheObj.sheets = cacheObj.ss.getSheets();
    return cacheObj.sheets;
  };
  const getSsByIdCached = (id) => {
    if (!ssCache[id]) ssCache[id] = SpreadsheetApp.openById(id);
    return ssCache[id];
  };
  // ==========================================

  const dCurr = new Date(targetDate.getFullYear(), targetDate.getMonth(), 1);
  const dPrev = new Date(targetDate.getFullYear(), targetDate.getMonth() - 1, 1);
  const dLast = new Date(targetDate.getFullYear() - 1, targetDate.getMonth(), 1);
  
  const periods = {
    current: { str: Utilities.formatDate(dCurr, "GMT+9", "yyyy/MM"), dot: `${dCurr.getFullYear()}.${dCurr.getMonth() + 1}`, year: dCurr.getFullYear(), monthNum: dCurr.getMonth() + 1, dateObj: dCurr, isRecent: true },
    prev:    { str: Utilities.formatDate(dPrev, "GMT+9", "yyyy/MM"), dot: `${dPrev.getFullYear()}.${dPrev.getMonth() + 1}`, year: dPrev.getFullYear(), monthNum: dPrev.getMonth() + 1, dateObj: dPrev, isRecent: true },
    last:    { str: Utilities.formatDate(dLast, "GMT+9", "yyyy/MM"), dot: `${dLast.getFullYear()}.${dLast.getMonth() + 1}`, year: dLast.getFullYear(), monthNum: dLast.getMonth() + 1, dateObj: dLast, isRecent: false }
  };

  // データソースのO(1)検索（高速化）
  const dsMap = {};
  const dsData = ss.getSheetByName('データソース').getDataRange().getValues();
  for (let i = 1; i < dsData.length; i++) {
    const name = String(dsData[i][0]).trim();
    if (name) dsMap[name] = { url: String(dsData[i][1]).trim(), s1: String(dsData[i][2]).trim(), s2: String(dsData[i][3]).trim() };
  }
  const getDsInfo = (name) => dsMap[name] || null;

  const parseDateSafe = (val) => {
    if (!val) return null;
    if (val instanceof Date) return val;
    const d = new Date(val);
    return !isNaN(d.getTime()) ? d : null;
  };
  const timeToMins = (dateObj) => {
    const d = parseDateSafe(dateObj);
    return d ? d.getHours() * 60 + d.getMinutes() : null;
  };
  const cleanId = (rawId) => String(rawId).replace(/\.0+$/, '').trim();
  const cleanStr = (str) => String(str).replace(/\s+/g, '');

  const clinicDict = {}; 
  const clinicAttrs = {}; 
  
  const regInfo = getDsInfo('正規表現');
  if (regInfo && regInfo.url) {
    try {
      const regData = getSheetsCached(regInfo.url)[0].getDataRange().getValues();
      const rHead = regData[0].map(String);
      const cOfficial = rHead.findIndex(h => h.includes('正規記載'));
      const cGroup = rHead.findIndex(h => h.includes('グループ')); 
      const cArea = rHead.findIndex(h => h === 'エリア' || h === '詳細エリア'); 
      const cOpen = rHead.findIndex(h => h.includes('開院日'));
      
      if (cOfficial !== -1) {
        for (let r = 1; r < regData.length; r++) {
          const offName = String(regData[r][cOfficial]).trim();
          if (!offName) continue;
          clinicAttrs[offName] = {
            group: cGroup !== -1 ? String(regData[r][cGroup]).trim() : '',
            area: cArea !== -1 ? String(regData[r][cArea]).trim() : '',
            openDate: parseDateSafe(regData[r][cOpen])
          };
          clinicDict[offName] = offName;
          rHead.forEach((h, i) => { if (h.includes('表記揺れ')) { const v = String(regData[r][i]).trim(); if (v) clinicDict[v] = offName; } });
        }
      }
    } catch(e) {}
  }

  const docInfoDict = {}; 
  ['紹介会社応募表', '特別対応医師'].forEach(dsName => {
    const info = getDsInfo(dsName);
    if (info && info.url) {
      try {
        const sData = getSheetsCached(info.url)[0].getDataRange().getValues();
        let hIdx = 0;
        for (let i = 0; i < 5; i++) { if (sData[i] && sData[i].join('').includes('医籍番号')) { hIdx = i; break; } }
        const cId = sData[hIdx].map(String).findIndex(h => h.includes('医籍番号'));
        const cName = sData[hIdx].map(String).findIndex(h => h.includes('氏名') || h.includes('名前'));
        for (let r = hIdx + 1; r < sData.length; r++) {
          const id = cId !== -1 ? cleanId(sData[r][cId]) : '';
          const name = cName !== -1 ? cleanStr(sData[r][cName]) : '';
          if (id) { if (!docInfoDict[id]) docInfoDict[id] = []; docInfoDict[id].push({ route: '紹介会社' }); }
          if (name) { if (!docInfoDict[name]) docInfoDict[name] = []; docInfoDict[name].push({ route: '紹介会社' }); }
        }
      } catch(e) {}
    }
  });

  [dCurr.getFullYear(), dLast.getFullYear()].forEach(year => {
    const mKeys = year === 2026 ? ['医師情報'] : ['2025定期非常勤', '2025常勤', '2025医師情報'];
    mKeys.forEach(mKey => {
      const mInfo = getDsInfo(mKey);
      if (!mInfo || !mInfo.url) return;
      try {
        const mSs = getSsByUrl(mInfo.url).ss;
        const sheetsToRead = year === 2026 ? [mInfo.s1, mInfo.s2].filter(Boolean) : [getSheetsCached(mInfo.url)[0].getName()];
        
        sheetsToRead.forEach(sName => {
          const mSheet = mSs.getSheetByName(sName);
          if (!mSheet) return;
          const mData = mSheet.getDataRange().getValues();
          let hIdx = -1;
          for (let i = 0; i < 10; i++) { if (mData[i].join('').includes('医籍番号')) { hIdx = i; break; } }
          if (hIdx === -1) return;
          
          const mHead = mData[hIdx].map(String);
          const cId = mHead.findIndex(h => h.includes('医籍番号'));
          const cType = mHead.findIndex(h => h.includes('医師区分') || h.includes('雇用区分')); 
          const cJoin = mHead.findIndex(h => h.includes('入職日'));
          const cLeave = mHead.findIndex(h => h.includes('退職日'));
          const cContract = mHead.findIndex(h => h.includes('契約内容') || h.includes('備考'));
          
          for (let r = hIdx + 1; r < mData.length; r++) {
            const id = cleanId(mData[r][cId]);
            if (!id) continue;
            
            let eType = '';
            if (cType !== -1) {
              const rawType = String(mData[r][cType]);
              if (rawType.match(/定期|非常勤/)) eType = '定期非常勤';
              else if (rawType.includes('常勤')) eType = '常勤';
            }
            if (!eType) eType = (sName.includes('定期') || sName.includes('非常勤') || mKey.includes('定期')) ? '定期非常勤' : '常勤';

            if (!docInfoDict[id]) docInfoDict[id] = [];
            docInfoDict[id].push({
              type: eType,
              joinDate: parseDateSafe(mData[r][cJoin]),
              leaveDate: parseDateSafe(mData[r][cLeave]),
              hasContract: (cContract !== -1 && String(mData[r][cContract]).trim().length > 5),
              masterYear: year
            });
          }
        });
      } catch(e) {}
    });
  });

  const closures = {}; 
  const holInfo = getDsInfo('未充足報告');
  if (holInfo && holInfo.url) {
    try {
      const holSheet = getSsByUrl(holInfo.url).ss.getSheetByName(holInfo.s1 || '休館日');
      if (holSheet) {
        const hData = holSheet.getDataRange().getValues();
        const hHead = hData[0].map(String);
        const colDate = hHead.findIndex(h => h.includes('日'));
        const colClinic = hHead.findIndex(h => h.includes('拠点') || h.includes('クリニック'));
        const colTime = hHead.findIndex(h => h.includes('時間'));
        for (let r = 1; r < hData.length; r++) {
          const dObj = parseDateSafe(hData[r][colDate]);
          const rClinic = colClinic !== -1 ? String(hData[r][colClinic]).trim() : '';
          const tPeriod = colTime !== -1 ? String(hData[r][colTime]).trim() : '全日';
          if (dObj && rClinic) {
            const dStr = Utilities.formatDate(dObj, "GMT+9", "yyyy/MM/dd");
            const oClinic = rClinic === '全拠点' ? '全拠点' : (clinicDict[rClinic] || rClinic);
            if (!closures[dStr]) closures[dStr] = [];
            closures[dStr].push({ clinic: oClinic, type: tPeriod });
          }
        }
      }
    } catch(e) {}
  }

  let cData = [], cHead = [];
  const chkInfo = getDsInfo('チェックリスト');
  if (chkInfo && chkInfo.url) {
    try {
      cData = getSheetsCached(chkInfo.url)[0].getDataRange().getValues();
      if(cData.length > 1) cHead = cData[1].map(String);
    } catch(e) {}
  }

  const targetAreas = ['関東', '関西', '関東第一', '関東第二', '埼玉', '神奈川', '千葉', '茨城', '大阪', 'グループ全体'];

  function extractForMonth(period) {
    const res = {};
    targetAreas.forEach(a => {
      res[a] = {
        hires: { direct: 0, agency: 0 },
        sales: { visitAct: 0, visitTgt: 0, salesAct: 0, salesTgt: 0 },
        uuDict: {}, 
        uu: { total: 0, reg: 0, hol: 0, part: 0, agc: 0, dir: 0 }, 
        ratioTime: { total: 0, reg: 0, hol: 0, part: 0, agc: 0, dir: 0 }, 
        wage: { areaAvg: 0, sum: 0, workMins: 0, requestAllowance: 0 },
        shiftDiff: { absenceMins: 0, overtimeMins: 0 },
        twoDoc: { hours: 0 },
        baseCount: { total: 0 },
        enrolled: { regular: 0, partTime: 0 },
        dailyShifts: {}
      };
    });

    const endOfMonth = new Date(period.dateObj.getFullYear(), period.dateObj.getMonth() + 1, 0);
    Object.keys(clinicAttrs).forEach(cName => {
      const oDate = clinicAttrs[cName].openDate;
      if (oDate && oDate <= endOfMonth) {
        ['グループ全体', clinicAttrs[cName].group, clinicAttrs[cName].area].filter(Boolean).forEach(a => { if (res[a]) res[a].baseCount.total++; });
      }
    });

    Object.keys(docInfoDict).forEach(id => {
      const infos = docInfoDict[id];
      let validInfo = null;
      for (const info of infos) {
        if (info.masterYear && info.masterYear !== period.year) continue; 
        if (!info.type || info.type === 'スポット') continue;
        if (info.joinDate && info.joinDate > endOfMonth) continue;
        if (info.leaveDate && info.leaveDate < period.dateObj) continue;
        validInfo = info; break;
      }
      if (validInfo) {
        res['グループ全体'].enrolled.regular += (validInfo.type === '常勤' ? 1 : 0);
        res['グループ全体'].enrolled.partTime += (validInfo.type === '定期非常勤' ? 1 : 0);
      }
    });

    if (cData.length > 2) {
      const colId = cHead.findIndex(h => h.includes('医籍番号'));
      const colDate = cHead.findIndex(h => h.includes('DS申請日'));
      const colStatus = cHead.findIndex(h => h.includes('採用可否'));
      const colSpec = cHead.findIndex(h => h.includes('診療科'));
      const colChannel = cHead.findIndex(h => h.includes('経緯'));
      const colClinic = cHead.findIndex(h => h.includes('拠点') && !h.includes('日'));

      for (let r = 2; r < cData.length; r++) {
        if (colDate === -1) continue;
        const rawDate = cData[r][colDate];
        let dStr = '';
        if (rawDate instanceof Date) dStr = Utilities.formatDate(rawDate, "GMT+9", "yyyy/MM");
        else if (rawDate) dStr = String(rawDate).trim().replace(/\./g, '/').substring(0, 7);
        
        if (dStr !== period.str) continue;
        
        const status = colStatus !== -1 ? String(cData[r][colStatus]) : '';
        const spec = colSpec !== -1 ? String(cData[r][colSpec]) : '';
        if (status.includes('採用') && spec.includes('小児科')) {
          const rawClinic = colClinic !== -1 ? String(cData[r][colClinic]).trim() : '';
          const officialClinic = clinicDict[rawClinic] || rawClinic;
          const channel = colChannel !== -1 ? String(cData[r][colChannel]).trim() : '';
          
          const areasToPush = ['グループ全体'];
          if (clinicAttrs[officialClinic]) {
            if (clinicAttrs[officialClinic].group) areasToPush.push(clinicAttrs[officialClinic].group);
            if (clinicAttrs[officialClinic].area) areasToPush.push(clinicAttrs[officialClinic].area);
          }
          areasToPush.forEach(a => {
            if (res[a]) {
              if (channel.includes('直接')) res[a].hires.direct++;
              else if (channel.match(/エムスリー|民間医局|エムステージ|Mステージ|マイナビ|MRT/)) res[a].hires.agency++;
            }
          });
        }
      }
    }

    try {
      const q = `title contains '業績管理シート' and title contains '全拠点' and title contains '${period.dot}' and mimeType = '${MimeType.GOOGLE_SHEETS}'`;
      const files = DriveApp.searchFiles(q);
      if (files.hasNext()) {
        const sData = getSsByIdCached(files.next().getId()).getSheetByName('来院数').getDataRange().getValues();
        let totalColIndex = -1;
        for (let c = 4; c < sData[2].length; c++) { if (String(sData[2][c]).includes('合計') || String(sData[3][c]).includes('合計')) { totalColIndex = c; break; } }
        
        if (totalColIndex !== -1) {
          const processed = { visitAct: new Set(), visitTgt: new Set(), salesAct: new Set(), salesTgt: new Set() };
          let currentMode = null;
          
          for (let r = 0; r < sData.length; r++) {
            const combinedVal = cleanStr(sData[r][0]) + cleanStr(sData[r][1]);
            
            if (combinedVal.includes('来院数：実績') || combinedVal.includes('来院数:実績')) { currentMode = 'visitAct'; continue; }
            if (combinedVal.includes('来院数目標') || combinedVal.includes('来院数：目標') || combinedVal.includes('来院数：予算')) { currentMode = 'visitTgt'; continue; }
            if (combinedVal.includes('売上：実績') || combinedVal.includes('売上:実績')) { currentMode = 'salesAct'; continue; }
            if (combinedVal.includes('売上：目標') || combinedVal.includes('売上：予算')) { currentMode = 'salesTgt'; continue; }
            if (combinedVal.includes('目標') || combinedVal.includes('実績')) { currentMode = null; continue; }

            if (currentMode) {
              const rawClinic = String(sData[r][1]).trim();
              if (!rawClinic || rawClinic === 'OK' || rawClinic.includes('計')) continue;
              const offClinic = clinicDict[rawClinic] || rawClinic;
              
              if (!processed[currentMode].has(offClinic)) {
                 const amt = Number(String(sData[r][totalColIndex]).replace(/,/g, ''));
                 if (!isNaN(amt)) {
                   const areasToPush = ['グループ全体'];
                   if (clinicAttrs[offClinic]) {
                     if (clinicAttrs[offClinic].group) areasToPush.push(clinicAttrs[offClinic].group);
                     if (clinicAttrs[offClinic].area) areasToPush.push(clinicAttrs[offClinic].area);
                   }
                   areasToPush.forEach(a => { if (res[a]) res[a].sales[currentMode] += amt; });
                   processed[currentMode].add(offClinic);
                 }
              }
            }
          }
        }
      }
    } catch(e) {}

    try {
      let shiftSheet = null;
      const arcInfo = getDsInfo(`${period.year}確定シフト`);
      if (arcInfo && arcInfo.url) {
        try {
          for (const s of getSheetsCached(arcInfo.url)) { 
            if (s.getName().match(new RegExp(`^0?${period.monthNum}月?`)) || s.getName().includes(period.str)) { shiftSheet = s; break; } 
          }
        } catch(e) {}
      }

      if (!shiftSheet) {
        const reqInfo = getDsInfo('請求書くん');
        if (reqInfo && reqInfo.url) {
          try {
            for (const s of getSheetsCached(reqInfo.url)) { 
              if (s.getName().match(new RegExp(`^0?${period.monthNum}月?`)) || s.getName().includes(period.str)) { shiftSheet = s; break; } 
            }
            if (!shiftSheet) shiftSheet = getSsByUrl(reqInfo.url).ss.getSheetByName('確定シフト') || getSheetsCached(reqInfo.url)[0];
          } catch(e) {}
        }
      }

      if (shiftSheet) {
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
          if (itemCol !== -1 && amountCol !== -1) {
            allowanceCols.push({ item: itemCol, amount: amountCol });
          }
        }

        for (let i = 1; i < sData.length; i++) {
          const rawClinic = colClinic !== -1 ? String(sData[i][colClinic]).trim() : '';
          const officialClinic = clinicDict[rawClinic];
          if (!officialClinic) continue; 
          
          if (colDepart !== -1 && String(sData[i][colDepart]).trim() !== '小児科') continue;
          
          const dateObj = parseDateSafe(sData[i][colDate]);
          if (!dateObj || Utilities.formatDate(dateObj, "GMT+9", "yyyy/MM") !== period.str) continue;
          
          const startMin = timeToMins(sData[i][colStart]);
          const endMin = timeToMins(sData[i][colEnd]);
          if (startMin === null || endMin === null) continue;
          
          const shiftDurMins = endMin - startMin;
          if (shiftDurMins < 30) continue; 
          
          const id = colId !== -1 ? cleanId(sData[i][colId]) : '';
          const name = colName !== -1 ? cleanStr(sData[i][colName]) : '';
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

          const docInfos = [...(docInfoDict[id] || []), ...(docInfoDict[name] || [])];
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

          const areasToPush = ['グループ全体'];
          if (clinicAttrs[officialClinic]) {
            if (clinicAttrs[officialClinic].group) areasToPush.push(clinicAttrs[officialClinic].group);
            if (clinicAttrs[officialClinic].area) areasToPush.push(clinicAttrs[officialClinic].area);
          }

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

            if (!d.dailyShifts[key]) d.dailyShifts[key] = { clinic: officialClinic, dateStr: dateStr, dateObj: dateObj, shifts: [] };
            d.dailyShifts[key].shifts.push({ start: startMin, end: endMin });
          });
        }

        targetAreas.forEach(a => {
          const d = res[a];
          for (const key in d.dailyShifts) {
            const info = d.dailyShifts[key];
            const attr = clinicAttrs[info.clinic];
            
            if (attr && attr.openDate && info.dateObj < attr.openDate) continue; 
            
            let eveningClose = (info.clinic.includes('北葛西')) ? 20 * 60 : 21 * 60;
            let expectedPeriods = [{ start: 9 * 60, end: 13 * 60 }, { start: 15 * 60, end: eveningClose }];
            
            const dayClosures = closures[info.dateStr] || [];
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
          }
        });
      }
    } catch (e) {}

    targetAreas.forEach(a => {
      const d = res[a];
      Object.values(d.uuDict).forEach(u => {
        d.uu.total++;
        d.uu[u.type]++;
      });
      d.wage.areaAvg = (d.wage.sum > 0 && d.wage.workMins > 0) ? Math.round(d.wage.sum / (d.wage.workMins / 60)) : 0;
      delete d.uuDict; delete d.dailyShifts; delete d.wage.sum; delete d.wage.workMins;
    });

    return res;
  }

  return {
    monthStrs: { current: periods.current.str, prev: periods.prev.str, last: periods.last.str },
    current: extractForMonth(periods.current),
    prev: extractForMonth(periods.prev),
    last: extractForMonth(periods.last)
  };
}