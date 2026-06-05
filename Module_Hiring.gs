function debugStep20_TrueAgencyLogic() {
  const ssId = '1zZaoAZ1q76hq_Dcpg0XGTzZFHQXrFU9E1GqYHeE7F2s';
  const ss = SpreadsheetApp.openById(ssId);
  const targetDate = new Date(2026, 4, 1); 

  const dCurr = new Date(targetDate.getFullYear(), targetDate.getMonth(), 1);
  const periodStr = Utilities.formatDate(dCurr, "GMT+9", "yyyy/MM");
  const periodYear = dCurr.getFullYear();

  const dsData = ss.getSheetByName('データソース').getDataRange().getValues();
  const getDsInfo = (name) => {
    const row = dsData.find(r => String(r[0]).trim() === name);
    return row ? { url: String(row[1]).trim(), s1: String(row[2]).trim(), s2: String(row[3]).trim() } : null;
  };

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
  const regInfo = getDsInfo('正規表現');
  if (regInfo && regInfo.url) {
    try {
      const regData = SpreadsheetApp.openByUrl(regInfo.url).getSheets()[0].getDataRange().getValues();
      const rHead = regData[0].map(String);
      const cOfficial = rHead.findIndex(h => h.includes('正規記載'));
      if (cOfficial !== -1) {
        for (let r = 1; r < regData.length; r++) {
          const offName = String(regData[r][cOfficial]).trim();
          if (offName) {
            clinicDict[offName] = offName;
            rHead.forEach((h, i) => { if (h.includes('表記揺れ')) { const v = String(regData[r][i]).trim(); if (v) clinicDict[v] = offName; } });
          }
        }
      }
    } catch(e) {}
  }

  // ==========================================
  // 【定義2 & 3】紹介会社応募表 ＆ 特別対応医師の取得
  // ==========================================
  const docInfoDict = {}; 
  ['紹介会社応募表', '特別対応医師'].forEach(dsName => {
    const info = getDsInfo(dsName);
    if (info && info.url) {
      try {
        const sData = SpreadsheetApp.openByUrl(info.url).getSheets()[0].getDataRange().getValues();
        let hIdx = 0;
        for (let i = 0; i < 5; i++) { if (sData[i] && sData[i].join('').includes('医籍番号')) { hIdx = i; break; } }
        const cId = sData[hIdx].map(String).findIndex(h => h.includes('医籍番号'));
        const cName = sData[hIdx].map(String).findIndex(h => h.includes('氏名') || h.includes('名前'));
        for (let r = hIdx + 1; r < sData.length; r++) {
          const id = cId !== -1 ? cleanId(sData[r][cId]) : '';
          const name = cName !== -1 ? cleanStr(sData[r][cName]) : '';
          if (id) { if (!docInfoDict[id]) docInfoDict[id] = []; docInfoDict[id].push({ route: '紹介会社', source: dsName }); }
          if (name) { if (!docInfoDict[name]) docInfoDict[name] = []; docInfoDict[name].push({ route: '紹介会社', source: dsName }); }
        }
      } catch(e) {}
    }
  });

  const mKeys = ['医師情報'];
  mKeys.forEach(mKey => {
    const mInfo = getDsInfo(mKey);
    if (!mInfo || !mInfo.url) return;
    try {
      const mSs = SpreadsheetApp.openByUrl(mInfo.url);
      [mInfo.s1, mInfo.s2].filter(Boolean).forEach(sName => {
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
          if (!eType) eType = (sName.includes('定期') || sName.includes('非常勤')) ? '定期非常勤' : '常勤';

          if (!docInfoDict[id]) docInfoDict[id] = [];
          docInfoDict[id].push({
            type: eType,
            joinDate: parseDateSafe(mData[r][cJoin]),
            leaveDate: parseDateSafe(mData[r][cLeave]),
            hasContract: (cContract !== -1 && String(mData[r][cContract]).trim().length > 5),
            masterYear: periodYear
          });
        }
      });
    } catch(e) {}
  });

  console.log(`\n=============================================`);
  console.log(`🔄 【対象月: 2026/05】 のシフト解析を開始...`);

  let shiftSheet = null;
  const arcInfo = getDsInfo(`2026確定シフト`);
  if (arcInfo && arcInfo.url) {
    try {
      const arcSs = SpreadsheetApp.openByUrl(arcInfo.url);
      for (const s of arcSs.getSheets()) { if (s.getName().match(/5月/) || s.getName().includes(periodStr)) { shiftSheet = s; break; } }
    } catch(e) {}
  }
  if (!shiftSheet) {
    const reqInfo = getDsInfo('請求書くん');
    if (reqInfo && reqInfo.url) {
      try {
        const reqSs = SpreadsheetApp.openByUrl(reqInfo.url);
        for (const s of reqSs.getSheets()) { if (s.getName().match(/5月/) || s.getName().includes(periodStr)) { shiftSheet = s; break; } }
        if (!shiftSheet) shiftSheet = reqSs.getSheetByName('確定シフト') || reqSs.getSheets()[0];
      } catch(e) {}
    }
  }

  if (!shiftSheet) { console.error("⚠️ 確定シフトが見つかりません。"); return; }

  const sData = shiftSheet.getDataRange().getValues();
  const headers = sData[0].map(String);
  const colClinic = headers.findIndex(h => h.includes('クリニック名'));
  const colDepart = headers.findIndex(h => h.includes('診療科'));
  const colDate = headers.findIndex(h => h.includes('勤務日'));
  const colStart = headers.findIndex(h => h.includes('勤務開始時間') && !h.includes('元のシフト'));
  const colEnd = headers.findIndex(h => h.includes('勤務終了時間') && !h.includes('元のシフト'));
  const colId = headers.findIndex(h => h.includes('医籍番号'));
  const colName = headers.findIndex(h => h.includes('氏名') || h.includes('名前'));
  
  const colComments = [];
  headers.forEach((h, idx) => { if (h.includes('スタッフコメント')) colComments.push(idx); });

  const uuDict = {};
  const ratioTime = { total: 0, reg: 0, hol: 0, part: 0, agc: 0, dir: 0 };
  let logCount = 0;

  for (let i = 1; i < sData.length; i++) {
    const rawClinic = colClinic !== -1 ? String(sData[i][colClinic]).trim() : '';
    const offClinic = clinicDict[rawClinic];
    if (!offClinic) continue; 
    if (colDepart !== -1 && String(sData[i][colDepart]).trim() !== '小児科') continue;
    
    const dateObj = parseDateSafe(sData[i][colDate]);
    if (!dateObj || Utilities.formatDate(dateObj, "GMT+9", "yyyy/MM") !== periodStr) continue;
    
    const startMin = timeToMins(sData[i][colStart]);
    const endMin = timeToMins(sData[i][colEnd]);
    if (startMin === null || endMin === null) continue;
    
    const shiftDurMins = endMin - startMin;
    if (shiftDurMins < 30) continue; 
    
    const id = colId !== -1 ? cleanId(sData[i][colId]) : '';
    const name = colName !== -1 ? cleanStr(sData[i][colName]) : '';
    
    let empType = 'スポット';
    let route = '';
    let isHol = false;

    // ==========================================
    // ★【定義1】 スタッフコメント欄による判定
    // ==========================================
    colComments.forEach(cIdx => {
      const comment = String(sData[i][cIdx] || '');
      // 休出判定
      if (comment.match(/所定休出|所定外勤務|所定外休出/)) isHol = true;
      // 紹介会社判定（スタッフコメント内）
      if (comment.match(/紹介|エムスリー|民間医局|エムステージ|Mステージ|マイナビ|MRT/)) {
         route = '紹介会社';
         if (logCount < 5) {
           console.log(`🔍 [紹介会社検知(コメント)] 氏名: ${name} | コメント: [${comment}]`);
           logCount++;
         }
      }
    });

    // ==========================================
    // ★【定義2 & 3】 マスタ登録状況による判定
    // ==========================================
    const docInfos = docInfoDict[id] || docInfoDict[name] || [];
    if (docInfos.length > 0) {
      for (const info of docInfos) {
        if (info.route) route = info.route; // 紹介会社・特別対応に存在すれば上書き
        if (info.masterYear && info.masterYear !== periodYear) continue; 
        
        let isWithinPeriod = true;
        if (info.joinDate && dateObj < info.joinDate) isWithinPeriod = false;
        if (info.leaveDate && dateObj > info.leaveDate) isWithinPeriod = false;
        
        if (isWithinPeriod && info.type) {
          empType = info.type;
          if (!info.hasContract) isHol = true; 
          break;
        }
      }
    }

    // 最終属性決定
    let finalAttr = 'dir';
    if (empType === '常勤') finalAttr = isHol ? 'hol' : 'reg';
    else if (empType === '定期非常勤') finalAttr = isHol ? 'dir' : 'part'; 
    else finalAttr = route === '紹介会社' ? 'agc' : 'dir';

    // 時間ベースの合算
    ratioTime.total += shiftDurMins;
    ratioTime[finalAttr] += shiftDurMins;

    // 人数ベース（重複排除用）
    const uuKey = id || name || `row_${i}`;
    if (!uuDict[uuKey]) {
      uuDict[uuKey] = finalAttr;
    } else {
      const priority = { 'hol': 5, 'reg': 4, 'part': 3, 'agc': 2, 'dir': 1 };
      if (priority[finalAttr] > priority[uuDict[uuKey]]) uuDict[uuKey] = finalAttr;
    }
  }

  console.log(`\n📊 【修正後の構成比 (グループ全体 / 時間ベース)】`);
  if (ratioTime.total > 0) {
    const rReg = (ratioTime.reg / ratioTime.total * 100).toFixed(1);
    const rPart = (ratioTime.part / ratioTime.total * 100).toFixed(1);
    const rDir = (ratioTime.dir / ratioTime.total * 100).toFixed(1);
    const rAgc = (ratioTime.agc / ratioTime.total * 100).toFixed(1);
    const rHol = (ratioTime.hol / ratioTime.total * 100).toFixed(1);
    console.log(`   常: ${rReg}% \n   定: ${rPart}% \n   直(ス): ${rDir}% \n   紹: ${rAgc}% \n   休: ${rHol}%`);
    console.log(`\n   (稼働時間: 総${Math.round(ratioTime.total/60)}h | 常:${Math.round(ratioTime.reg/60)}h, 定:${Math.round(ratioTime.part/60)}h, 直:${Math.round(ratioTime.dir/60)}h, 紹:${Math.round(ratioTime.agc/60)}h, 休:${Math.round(ratioTime.hol/60)}h)`);
  }
  console.log(`=============================================`);
}