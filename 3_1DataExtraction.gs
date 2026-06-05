// =========================================================
// 1. メイン抽出関数（各月ループの呼び出し元）
// =========================================================
function extractAllData(targetDate) {
  const ctx = buildContext(targetDate);
  loadMasterData(ctx);

  return {
    monthStrs: { current: ctx.periods.current.str, prev: ctx.periods.prev.str, last: ctx.periods.last.str },
    current: extractForMonth(ctx, ctx.periods.current),
    prev: extractForMonth(ctx, ctx.periods.prev),
    last: extractForMonth(ctx, ctx.periods.last)
  };
}

// =========================================================
// 2. コンテキスト（共通設定・ヘルパー）の初期化
// =========================================================
function buildContext(targetDate) {
  const ssId = '1zZaoAZ1q76hq_Dcpg0XGTzZFHQXrFU9E1GqYHeE7F2s'; // 本番用ID
  const ss = SpreadsheetApp.openById(ssId);
  
  const dCurr = new Date(targetDate.getFullYear(), targetDate.getMonth(), 1);
  const dPrev = new Date(targetDate.getFullYear(), targetDate.getMonth() - 1, 1);
  const dLast = new Date(targetDate.getFullYear() - 1, targetDate.getMonth(), 1);

  const ctx = {
    ss: ss,
    targetAreas: ['関東', '関西', '関東第一', '関東第二', '埼玉', '神奈川', '千葉', '茨城', '大阪', 'グループ全体'],
    periods: {
      current: { str: Utilities.formatDate(dCurr, "GMT+9", "yyyy/MM"), dot: `${dCurr.getFullYear()}.${dCurr.getMonth() + 1}`, year: dCurr.getFullYear(), monthNum: dCurr.getMonth() + 1, dateObj: dCurr },
      prev:    { str: Utilities.formatDate(dPrev, "GMT+9", "yyyy/MM"), dot: `${dPrev.getFullYear()}.${dPrev.getMonth() + 1}`, year: dPrev.getFullYear(), monthNum: dPrev.getMonth() + 1, dateObj: dPrev },
      last:    { str: Utilities.formatDate(dLast, "GMT+9", "yyyy/MM"), dot: `${dLast.getFullYear()}.${dLast.getMonth() + 1}`, year: dLast.getFullYear(), monthNum: dLast.getMonth() + 1, dateObj: dLast }
    },
    ssCache: {},
    getSsByUrl: function(url) { if (!this.ssCache[url]) this.ssCache[url] = { ss: SpreadsheetApp.openByUrl(url), sheets: null }; return this.ssCache[url]; },
    getSheetsCached: function(url) { const cacheObj = this.getSsByUrl(url); if (!cacheObj.sheets) cacheObj.sheets = cacheObj.ss.getSheets(); return cacheObj.sheets; },
    getSsByIdCached: function(id) { if (!this.ssCache[id]) this.ssCache[id] = SpreadsheetApp.openById(id); return this.ssCache[id]; },
    
    dsMap: {},
    getDsInfo: function(name) { return this.dsMap[name] || null; },
    parseDateSafe: (val) => { if (!val) return null; if (val instanceof Date) return val; const d = new Date(val); return !isNaN(d.getTime()) ? d : null; },
    timeToMins: function(dateObj) { const d = this.parseDateSafe(dateObj); return d ? d.getHours() * 60 + d.getMinutes() : null; },
    cleanId: (rawId) => String(rawId).replace(/\.0+$/, '').trim(),
    cleanStr: (str) => String(str).replace(/\s+/g, ''),

    clinicDict: {}, clinicAttrs: {}, docInfoDict: {}, closures: {}, chkData: [], chkHead: [], firstShiftMap: {},

    getTargetAreasForClinic: function(rawClinic) {
      const areas = new Set(['グループ全体']);
      const cStr = String(rawClinic || '').trim();
      if (!cStr) return Array.from(areas).filter(a => this.targetAreas.includes(a));

      const offClinic = this.clinicDict[cStr] || cStr;
      const attrs = this.clinicAttrs[offClinic];
      if (attrs) {
        if (attrs.group) areas.add(attrs.group);
        if (attrs.area) areas.add(attrs.area);
      }
      const currentAreas = Array.from(areas).join(',');
      if (currentAreas.match(/東京|埼玉|神奈川|千葉|茨城|関東第一|関東第二/) || cStr.match(/東京|埼玉|神奈川|千葉|茨城/)) areas.add('関東');
      if (currentAreas.match(/大阪|兵庫|関西/) || cStr.match(/大阪|兵庫/)) areas.add('関西');
      return Array.from(areas).filter(a => this.targetAreas.includes(a));
    }
  };

  const dsData = ss.getSheetByName('データソース').getDataRange().getValues();
  for (let i = 1; i < dsData.length; i++) {
    const name = String(dsData[i][0]).trim();
    if (name) ctx.dsMap[name] = { url: String(dsData[i][1]).trim(), s1: String(dsData[i][2]).trim(), s2: String(dsData[i][3]).trim() };
  }
  return ctx;
}

// =========================================================
// 3. マスタ・各種辞書の構築処理
// =========================================================
function loadMasterData(ctx) {
  // 正規表現シートの読み込み
  const regInfo = ctx.getDsInfo('正規表現');
  if (regInfo && regInfo.url) {
    try {
      const regData = ctx.getSheetsCached(regInfo.url)[0].getDataRange().getValues();
      const rHead = regData[0].map(String);
      const cOfficial = rHead.findIndex(h => h.includes('正規記載'));
      const cGroup = rHead.findIndex(h => h.includes('グループ'));
      const cArea = rHead.findIndex(h => h === 'エリア' || h === '詳細エリア'); 
      const cOpen = rHead.findIndex(h => h.includes('開院日'));

      if (cOfficial !== -1) {
        for (let r = 1; r < regData.length; r++) {
          const offName = String(regData[r][cOfficial]).trim();
          if (!offName) continue;
          ctx.clinicAttrs[offName] = { group: cGroup !== -1 ? String(regData[r][cGroup]).trim() : '', area: cArea !== -1 ? String(regData[r][cArea]).trim() : '', openDate: ctx.parseDateSafe(regData[r][cOpen]) };
          ctx.clinicDict[offName] = offName;
          rHead.forEach((h, i) => { if (h.includes('表記揺れ')) { const v = String(regData[r][i]).trim(); if (v) ctx.clinicDict[v] = offName; } });
        }
      }
    } catch(e) {}
  }

  // 紹介会社・特別対応の読み込み
  ['紹介会社応募表', '特別対応医師'].forEach(dsName => {
    const info = ctx.getDsInfo(dsName);
    if (info && info.url) {
      try {
        const sData = ctx.getSheetsCached(info.url)[0].getDataRange().getValues();
        let hIdx = 0;
        for (let i = 0; i < 5; i++) { if (sData[i] && sData[i].join('').includes('医籍番号')) { hIdx = i; break; } }
        const cId = sData[hIdx].map(String).findIndex(h => h.includes('医籍番号'));
        const cName = sData[hIdx].map(String).findIndex(h => h.includes('氏名') || h.includes('名前'));
        for (let r = hIdx + 1; r < sData.length; r++) {
          const id = cId !== -1 ? ctx.cleanId(sData[r][cId]) : '';
          const name = cName !== -1 ? ctx.cleanStr(sData[r][cName]) : '';
          if (id) { if (!ctx.docInfoDict[id]) ctx.docInfoDict[id] = []; ctx.docInfoDict[id].push({ route: '紹介会社' }); }
          if (name) { if (!ctx.docInfoDict[name]) ctx.docInfoDict[name] = []; ctx.docInfoDict[name].push({ route: '紹介会社' }); }
        }
      } catch(e) {}
    }
  });

  // 医師情報（常勤・定期）マスタの読み込み
  [ctx.periods.current.year, ctx.periods.last.year].forEach(year => {
    const mKeys = year === 2026 ? ['医師情報'] : ['2025定期非常勤', '2025常勤', '2025医師情報'];
    mKeys.forEach(mKey => {
      const mInfo = ctx.getDsInfo(mKey);
      if (!mInfo || !mInfo.url) return;
      try {
        const mSs = ctx.getSsByUrl(mInfo.url).ss;
        const sheetsToRead = year === 2026 ? [mInfo.s1, mInfo.s2].filter(Boolean) : [ctx.getSheetsCached(mInfo.url)[0].getName()];
        
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
            const id = ctx.cleanId(mData[r][cId]);
            if (!id) continue;
            let eType = '';
            if (cType !== -1) {
              const rawType = String(mData[r][cType]);
              if (rawType.match(/定期|非常勤/)) eType = '定期非常勤';
              else if (rawType.includes('常勤')) eType = '常勤';
            }
            if (!eType) eType = (sName.includes('定期') || sName.includes('非常勤') || mKey.includes('定期')) ? '定期非常勤' : '常勤';

            if (!ctx.docInfoDict[id]) ctx.docInfoDict[id] = [];
            ctx.docInfoDict[id].push({ type: eType, joinDate: ctx.parseDateSafe(mData[r][cJoin]), leaveDate: ctx.parseDateSafe(mData[r][cLeave]), hasContract: (cContract !== -1 && String(mData[r][cContract]).trim().length > 5), masterYear: year });
          }
        });
      } catch(e) {}
    });
  });

  // 採用拠点補完用の「初回勤務マップ」作成
  const shiftUrls = [
    'https://docs.google.com/spreadsheets/d/1JhLJuxkp6T5hEkfultUE57pmRXBI_lkN6ROp2-CuhCc/edit', // 2025確定シフト
    'https://docs.google.com/spreadsheets/d/10Z8jg4o7Ri9Ggf7u_jauyDDZa6bVnc-iuuuqmkNDafw/edit'  // 2026確定シフト
  ];
  shiftUrls.forEach(url => {
    try {
      const shiftSs = SpreadsheetApp.openByUrl(url);
      for (const s of shiftSs.getSheets()) {
        const sData = s.getDataRange().getValues();
        if (sData.length < 2) continue;
        const headers = sData[0].map(String);
        const colClinic = headers.findIndex(h => h.includes('クリニック名'));
        const colDate = headers.findIndex(h => h.includes('勤務日'));
        const colId = headers.findIndex(h => h.includes('医籍番号'));
        const colName = headers.findIndex(h => h.includes('氏名') || h.includes('名前'));

        if (colClinic === -1 || colDate === -1) continue;

        for (let i = 1; i < sData.length; i++) {
          const rawClinic = String(sData[i][colClinic]).trim();
          const dateObj = ctx.parseDateSafe(sData[i][colDate]);
          if (!rawClinic || !dateObj) continue;

          const id = colId !== -1 ? ctx.cleanId(sData[i][colId]) : '';
          const name = colName !== -1 ? ctx.cleanStr(sData[i][colName]) : '';
          const time = dateObj.getTime();

          if (id) { if (!ctx.firstShiftMap[id] || time < ctx.firstShiftMap[id].time) ctx.firstShiftMap[id] = { clinic: rawClinic, time: time }; }
          if (name) { if (!ctx.firstShiftMap[name] || time < ctx.firstShiftMap[name].time) ctx.firstShiftMap[name] = { clinic: rawClinic, time: time }; }
        }
      }
    } catch(e) {}
  });

  // 休館日・未充足の読み込み
  const holInfo = ctx.getDsInfo('未充足報告');
  if (holInfo && holInfo.url) {
    try {
      const holSheet = ctx.getSsByUrl(holInfo.url).ss.getSheetByName(holInfo.s1 || '休館日');
      if (holSheet) {
        const hData = holSheet.getDataRange().getValues();
        const hHead = hData[0].map(String);
        const colDate = hHead.findIndex(h => h.includes('日'));
        const colClinic = hHead.findIndex(h => h.includes('拠点') || h.includes('クリニック'));
        const colTime = hHead.findIndex(h => h.includes('時間'));

        for (let r = 1; r < hData.length; r++) {
          const dObj = ctx.parseDateSafe(hData[r][colDate]);
          const rClinic = colClinic !== -1 ? String(hData[r][colClinic]).trim() : '';
          const tPeriod = colTime !== -1 ? String(hData[r][colTime]).trim() : '全日';
          if (dObj && rClinic) {
            const dStr = Utilities.formatDate(dObj, "GMT+9", "yyyy/MM/dd");
            const oClinic = rClinic === '全拠点' ? '全拠点' : (ctx.clinicDict[rClinic] || rClinic);
            if (!ctx.closures[dStr]) ctx.closures[dStr] = [];
            ctx.closures[dStr].push({ clinic: oClinic, type: tPeriod });
          }
        }
      }
    } catch(e) {}
  }

  // チェックリスト読み込み
  const chkInfo = ctx.getDsInfo('チェックリスト');
  if (chkInfo && chkInfo.url) {
    try {
      ctx.chkData = ctx.getSheetsCached(chkInfo.url)[0].getDataRange().getValues();
      if(ctx.chkData.length > 1) ctx.chkHead = ctx.chkData[1].map(String);
    } catch(e) {}
  }
}