/**
 * 確定シフト・マスタ等を読み込み、時給・UU・シフト調整・在籍数をエリア別に計算する
 * @param {Date} targetDate - 集計対象月の1日の日付オブジェクト
 * @param {Object} urls - データソースURLの辞書
 * @returns {Object} エリアごとの計算結果
 */
function calcShiftData(targetDate, urls) {
  const targetMonthStr = Utilities.formatDate(targetDate, "GMT+9", "yyyy/MM");
  const targetYear = targetDate.getFullYear();
  const targetMonthNum = targetDate.getMonth() + 1; // 1〜12
  
  const targetAreas = ['関東', '関西', '関東第一', '関東第二', '埼玉', '神奈川', '千葉', '茨城', '大阪', 'グループ全体'];
  const res = {};
  
  // 1. エリアごとのデータ構造と、UU計算用のSet（重複排除用）を初期化
  targetAreas.forEach(area => {
    res[area] = {
      wage: { sumArea: 0, countArea: 0, sumDetail: 0, countDetail: 0, areaAvg: 0, detailAvg: 0, requestAllowance: 0 },
      shiftDiff: { absenceMins: 0, overtimeMins: 0 },
      twoDoc: { hours: 0 },
      baseCount: { total: 0 },
      enrolled: { regular: 0, partTime: 0 },
      // 最終結果用
      uu: { total: 0, spot: 0, holiday: 0, agency: 0, regular: 0, partTime: 0, direct: 0 },
      // 計算中の一時保存用Set（IDまたは名前を放り込んで重複を弾く）
      _uuSets: { total: new Set(), spot: new Set(), holiday: new Set(), agency: new Set(), regular: new Set(), partTime: new Set(), direct: new Set() },
      _baseSet: new Set() // 対象拠点数カウント用
    };
  });

  // エリア判定ヘルパー
  function getAreas(clinicName) {
    const areas = ['グループ全体'];
    let clinicStr = String(clinicName || '').trim();
    if (clinicName instanceof Date || clinicStr.match(/GMT/)) clinicStr = '';

    if (clinicStr) {
      if (clinicStr.match(/大阪|阿波座|関西|豊中|吹田|高槻|茨木|堺|枚方/)) { areas.push('関西', '大阪'); }
      else if (clinicStr.match(/埼玉|草加|大宮|川口|浦和|川越|越谷/)) { areas.push('関東', '埼玉'); }
      else if (clinicStr.match(/神奈川|茅ヶ崎|天王町|横浜|川崎|武蔵小杉|藤沢|戸塚/)) { areas.push('関東', '神奈川'); }
      else if (clinicStr.match(/千葉|村上|船橋|柏|松戸|市川/)) { areas.push('関東', '千葉'); }
      else if (clinicStr.match(/茨城|水戸|つくば|守谷/)) { areas.push('関東', '茨城'); }
      else if (clinicStr.match(/第一/)) { areas.push('関東', '関東第一'); }
      else if (clinicStr.match(/第二/)) { areas.push('関東', '関東第二'); }
      else { areas.push('関東'); } 
    } else {
      areas.push('関東');
    }
    return areas;
  }

  // ==========================================
  // ① 確定シフトの検索と集計（UU・時給・残業・不在・2診）
  // ==========================================
  try {
    // 例: "2026年度" "確定シフト" でドライブを検索
    const searchQuery = `title contains '確定シフト' and title contains '${targetYear}'`;
    const files = DriveApp.searchFiles(searchQuery);
    
    if (files.hasNext()) {
      const shiftSs = SpreadsheetApp.openById(files.next().getId());
      
      // 対象月のシートを探す（例: "5月" または "05月" または "2026/05"）
      let shiftSheet = null;
      const sheets = shiftSs.getSheets();
      for (const s of sheets) {
        const sName = s.getName();
        if (sName === `${targetMonthNum}月` || sName === `0${targetMonthNum}月` || sName.includes(targetMonthStr)) {
          shiftSheet = s; break;
        }
      }
      
      if (!shiftSheet) shiftSheet = sheets[0]; // 見つからなければ暫定で1枚目

      const sData = shiftSheet.getDataRange().getValues();
      const headers = sData[0].map(String);
      
      const findCol = (kw) => headers.findIndex(h => h.includes(kw));
      const colId = findCol('医籍番号');
      const colName = findCol('医師名') !== -1 ? findCol('医師名') : findCol('氏名');
      const colClinic = findCol('拠点') !== -1 ? findCol('拠点') : findCol('クリニック');
      const colType = findCol('勤務形態'); // 常勤、定期非常勤、スポット等
      const colRoute = findCol('経緯'); // 直接、紹介会社、所定休出等
      const colWageArea = findCol('エリア時給'); // エリア平均用
      const colWageDetail = findCol('時給'); // 詳細時給用
      const colAllow = findCol('手当');
      const colOvertime = findCol('残業');
      const colAbsence = findCol('不在');
      const colTwoDoc = findCol('２診');

      for (let r = 1; r < sData.length; r++) {
        const id = colId !== -1 ? String(sData[r][colId]).trim() : '';
        const name = colName !== -1 ? String(sData[r][colName]).replace(/\s+/g, '') : '';
        const docKey = id || name || `row_${r}`; // 医師を特定するキー
        if (!name && !id) continue; // 空行スキップ
        
        const clinic = colClinic !== -1 ? String(sData[r][colClinic]).trim() : '';
        const tAreas = getAreas(clinic);
        
        const type = colType !== -1 ? String(sData[r][colType]).trim() : '';
        const route = colRoute !== -1 ? String(sData[r][colRoute]).trim() : '';
        
        const wArea = colWageArea !== -1 ? Number(sData[r][colWageArea]) : 0;
        const wDet = colWageDetail !== -1 ? Number(sData[r][colWageDetail]) : 0;
        const allow = colAllow !== -1 ? Number(sData[r][colAllow]) : 0;
        const over = colOvertime !== -1 ? Number(sData[r][colOvertime]) : 0;
        const abs = colAbsence !== -1 ? Number(sData[r][colAbsence]) : 0;
        const twoDoc = colTwoDoc !== -1 ? Number(sData[r][colTwoDoc]) : 0;

        tAreas.forEach(a => {
          const d = res[a];
          
          // 拠点数のカウント用
          if (clinic) d._baseSet.add(clinic);

          // 数値の加算
          if (wArea > 0) { d.wage.sumArea += wArea; d.wage.countArea++; }
          if (wDet > 0) { d.wage.sumDetail += wDet; d.wage.countDetail++; }
          if (allow > 0) { d.wage.requestAllowance += allow; }
          if (over > 0) { d.shiftDiff.overtimeMins += over; }
          if (abs > 0) { d.shiftDiff.absenceMins += abs; }
          if (twoDoc > 0) { d.twoDoc.hours += twoDoc; }
          
          // UU（稼働人員）の振り分け（Setに追加することで自動的に重複排除）
          d._uuSets.total.add(docKey);
          
          if (type.includes('常勤')) d._uuSets.regular.add(docKey);
          else if (type.includes('定期') || type.includes('非常勤')) d._uuSets.partTime.add(docKey);
          else d._uuSets.spot.add(docKey); // 常勤・定期以外はスポット扱い
          
          if (route.includes('休出')) d._uuSets.holiday.add(docKey);
          if (route.includes('直接')) d._uuSets.direct.add(docKey);
          if (['エムスリー', '民間医局', 'エムステージ', 'Mステージ', 'マイナビ', 'MRT', '紹介'].some(kw => route.includes(kw))) {
            d._uuSets.agency.add(docKey);
          }
        });
      }
    }
  } catch (e) { console.log(`[警告] 確定シフトの集計エラー: ${e.message}`); }

  // ==========================================
  // ② 在籍医師数の集計（マスタ または 在籍くん）
  // ==========================================
  if (urls['マスタ']) {
    try {
      const mSs = SpreadsheetApp.openByUrl(urls['マスタ']);
      const mSheet = mSs.getSheetByName('マスタ') || mSs.getSheetByName('社員マスタ変更リスト') || mSs.getSheets()[0];
      const mData = mSheet.getDataRange().getValues();
      const mHeaders = mData[5].map(String); // 6行目がヘッダー前提
      
      const mColName = mHeaders.findIndex(h => h.includes('氏'));
      const mColClinic = mHeaders.findIndex(h => h.includes('拠点'));
      const mColType = mHeaders.findIndex(h => h.includes('雇用区分') || h.includes('職位'));
      
      for (let r = 6; r < mData.length; r++) {
        if (!mData[r][mColName]) continue;
        const clinic = mColClinic !== -1 ? String(mData[r][mColClinic]).trim() : '';
        const type = mColType !== -1 ? String(mData[r][mColType]).trim() : '';
        
        const tAreas = getAreas(clinic);
        tAreas.forEach(a => {
          if (type.includes('常勤')) res[a].enrolled.regular++;
          else if (type.includes('非常勤')) res[a].enrolled.partTime++;
        });
      }
    } catch(e) { console.log(`[警告] 在籍数の集計エラー: ${e.message}`); }
  }

  // ==========================================
  // ③ 最終計算（Setのカウント化・平均時給の算出）
  // ==========================================
  targetAreas.forEach(a => {
    const d = res[a];
    
    // Setの要素数をカウントに変換
    d.uu.total = d._uuSets.total.size;
    d.uu.spot = d._uuSets.spot.size;
    d.uu.holiday = d._uuSets.holiday.size;
    d.uu.agency = d._uuSets.agency.size;
    d.uu.regular = d._uuSets.regular.size;
    d.uu.partTime = d._uuSets.partTime.size;
    d.uu.direct = d._uuSets.direct.size;
    
    d.baseCount.total = d._baseSet.size;
    
    // 時給の平均化
    d.wage.areaAvg = d.wage.countArea > 0 ? Math.round(d.wage.sumArea / d.wage.countArea) : 0;
    d.wage.detailAvg = d.wage.countDetail > 0 ? Math.round(d.wage.sumDetail / d.wage.countDetail) : 0;
    
    // ゴミデータのお掃除
    delete d._uuSets;
    delete d._baseSet;
  });

  return res;
}