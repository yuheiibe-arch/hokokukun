// ==========================================
// 1. レポートテキストの自動生成（タグ構造最適化・関東関西対応・末尾修正版）
// ==========================================
function generateReportText(targetArea = 'グループ全体') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName('ベースシート');
  if (!baseSheet) throw new Error('ベースシートが見つかりません。');

  const bHeaders = baseSheet.getRange(1, 1, 1, baseSheet.getLastColumn()).getValues()[0];
  let colIdx = -1;
  let kantoColIdx = -1;
  let kansaiColIdx = -1;

  for (let c = 3; c < bHeaders.length; c++) {
    const hText = String(bHeaders[c]).trim();
    if (hText === targetArea) colIdx = c + 1;
    if (hText === '関東') kantoColIdx = c + 1;
    if (hText === '関西') kansaiColIdx = c + 1;
  }
  if (colIdx === -1) throw new Error(`${targetArea}の列が見つかりません。`);

  const today = new Date();
  const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const monthDisplay = (lastMonth.getMonth() + 1) + '月';

  const abColValues = baseSheet.getRange(1, 1, baseSheet.getLastRow(), 2).getValues();
  const searchTexts = abColValues.map((row, index) => ({
    rowNum: index + 1,
    text: (String(row[0]) + String(row[1])).replace(/\s+/g, '') 
  }));

  const metricsDef = [
    { key: 'エリア平均時給', displayName: 'エリア平均時給', unit: '¥', type: 'money', cat: '💰 時給・コスト指標', exclude: [] },
    
    { key: '稼働人員（全医師', displayName: '稼働人員（全医師）', unit: '人', type: 'people', cat: '👥 稼働・人員状況', exclude: [] },
    { key: '稼働人員（常勤除く', displayName: '稼働人員（常勤除く）', unit: '人', type: 'people', cat: '', exclude: [] },
    { key: '所定休出医師数', displayName: '所定休出医師数', unit: '人', type: 'people', cat: '', exclude: [] },
    { key: '構成比', displayName: '構成比 (常勤/定非/直応募/紹介/休出)', unit: '', type: 'text', cat: '', exclude: [] },
    
    { key: '２診時間', displayName: '２診時間数', unit: 'h', type: 'time', cat: '🏥 シフト・拠点状況', exclude: [] },
    { key: '2診拠点数', displayName: '２診拠点数', unit: '拠点', type: 'num', cat: '', exclude: [] },
    { key: '対象拠点数', displayName: '営業拠点数', unit: '拠点', type: 'num', cat: '', exclude: ['2診', '２診'] }, 
    
    { key: '新規採用', displayName: '新規採用（直接）', unit: '人', type: 'people', cat: '✨ 採用・在籍状況', exclude: ['紹介', 'エージェント'] },
    { key: '新規採用', displayName: '新規採用（紹介）', unit: '人', type: 'people', cat: '', exclude: ['直接'] },
    
    { key: '医師勤務時間', displayName: '医師勤務時間', unit: 'h', type: 'time', cat: '⚙️ シフト調整関係', exclude: ['固定', '募集'] },
    { key: '固定シフト', displayName: '医師勤務時間（固定）', unit: 'h', type: 'time', cat: '', exclude: [] },
    { key: '募集シフト', displayName: '医師勤務時間（募集）', unit: 'h', type: 'time', cat: '', exclude: [] },
    { key: '不在時間', displayName: '医師不在時間', unit: 'h', type: 'time', cat: '', exclude: [] },
    { key: '残業時間', displayName: '医師残業時間', unit: 'h', type: 'time', cat: '', exclude: [] },
    
    { key: '来院数', displayName: '来院数 (目標達成率)', unit: '人', type: 'mix-people', cat: '📈 業績実績', exclude: [] }, 
    { key: '売上（目標達成率', displayName: '売上 (目標達成率)', unit: '¥', type: 'mix-money', cat: '', exclude: [] },
   
    { key: '在籍医師数（常勤', displayName: '在籍医師数（小児科・常勤）', unit: '人', type: 'people', cat: '🏢 在籍医師数', exclude: [] },
    { key: '在籍医師数（定期', displayName: '在籍医師数（小児科・定期）', unit: '人', type: 'people', cat: '', exclude: [] }
  ];

  const metrics = [];
  metricsDef.forEach(def => {
    const found = searchTexts.find(item => {
      const hasKey = item.text.includes(def.key);
      const hasNoExclude = def.exclude.every(exKw => !item.text.includes(exKw));
      return hasKey && hasNoExclude;
    });
    
    if (found && !metrics.some(m => m[0] === def.displayName)) {
      metrics.push([def.displayName, found.rowNum, def.unit, def.type, def.cat]);
    }
  });

  const parseVal = (rawStr, type) => {
    if (!rawStr || rawStr === '-') return { val: 0, str: '-' };
    const str = String(rawStr).split(' ')[0]; 
    
    if (type === 'time') {
      let h = 0, m = 0;
      const hm = String(rawStr).match(/(\d+)時間/); 
      const mm = String(rawStr).match(/(\d+)分/);
      if (hm) h = parseInt(hm[1]); 
      if (mm) m = parseInt(mm[1]);
      return { val: h + (m / 60), str: (h + (m / 60)).toFixed(1) };
    }
    const n = parseFloat(str.replace(/[^\d.-]/g, ''));
    return { val: isNaN(n) ? 0 : n, str: str };
  };

  const formatNum = (num, type, unit) => {
    if (type === 'money' || type === 'mix-money') return `¥${Math.round(num).toLocaleString()}`;
    if (type === 'time') return `${num.toFixed(1)}${unit}`;
    if (type === 'people' || type === 'mix-people' || type === 'num') return `${Math.round(num).toLocaleString()}${unit}`;
    return num;
  };

  const formatDiff = (curr, past, type, unit) => {
    const diff = curr - past;
    if (diff === 0) return '±0';
    const sign = diff > 0 ? '+' : '-';
    return `${sign}${formatNum(Math.abs(diff), type, unit)}`;
  };

  let reportText = `[info]\n[title]月次実績報告（${targetArea}）[/title]\nお疲れ様です。小児科 ${monthDisplay}の月次実績をご報告いたします。\n\n`;
  let currentCategory = '';

  metrics.forEach(m => {
    const [name, bRow, unit, type, catName] = m;
    
    if (catName && catName !== currentCategory) {
      reportText += `[hr]\n■ ${catName}\n`;
      currentCategory = catName;
    }

    const valCurrRaw = baseSheet.getRange(bRow, colIdx).getDisplayValue();
    const valPrevRaw = baseSheet.getRange(bRow + 1, colIdx).getDisplayValue();
    const valLastRaw = baseSheet.getRange(bRow + 2, colIdx).getDisplayValue();

    if (type === 'text') {
      let textVal = valCurrRaw.replace(/\n/g, ' / ');
      if (name.includes('構成比')) {
        textVal = textVal.replace(/常:/g, '常勤:')
                         .replace(/定:/g, '定非:')
                         .replace(/直\(ス\):/g, '直応募（ス）:')
                         .replace(/紹:/g, '紹介:')
                         .replace(/休:/g, '所定休出:');
      }
      reportText += `・${name}\n  └ 当月実績: ${textVal}\n`;
    } else {
      const c = parseVal(valCurrRaw, type);
      const p = parseVal(valPrevRaw, type);
      const l = parseVal(valLastRaw, type);

      const diffP = formatDiff(c.val, p.val, type, unit);
      const diffL = formatDiff(c.val, l.val, type, unit);
      
      let displayCurr = valCurrRaw;
      
      if (type.startsWith('mix')) {
        const parts = String(valCurrRaw).split(' ');
        if (parts.length > 0 && parts[0] !== '-') {
          let numRaw = parseFloat(parts[0].replace(/[^\d.-]/g, ''));
          if (!isNaN(numRaw)) {
            let formattedNum = formatNum(numRaw, type, unit);
            displayCurr = parts.length > 1 ? `${formattedNum} ${parts[1]}` : formattedNum;
          }
        }
      } else {
        displayCurr = formatNum(c.val, type, unit);
      }

      reportText += `・${name}: ${displayCurr} (前月: ${diffP} / 昨年: ${diffL})\n`;

      if (name === 'エリア平均時給') {
        if (kantoColIdx !== -1) {
          const kCurr = baseSheet.getRange(bRow, kantoColIdx).getDisplayValue();
          const kPrev = baseSheet.getRange(bRow + 1, kantoColIdx).getDisplayValue();
          const kLast = baseSheet.getRange(bRow + 2, kantoColIdx).getDisplayValue();
          const ck = parseVal(kCurr, type);
          const pk = parseVal(kPrev, type);
          const lk = parseVal(kLast, type);
          const kDiffP = formatDiff(ck.val, pk.val, type, unit);
          const kDiffL = formatDiff(ck.val, lk.val, type, unit);
          const kDisp = formatNum(ck.val, type, unit);
          reportText += `  ・関東\n    └ ${kDisp} (前月: ${kDiffP} / 昨年: ${kDiffL})\n`;
        }
        if (kansaiColIdx !== -1) {
          const sCurr = baseSheet.getRange(bRow, kansaiColIdx).getDisplayValue();
          const sPrev = baseSheet.getRange(bRow + 1, kansaiColIdx).getDisplayValue();
          const sLast = baseSheet.getRange(bRow + 2, kansaiColIdx).getDisplayValue();
          const cs = parseVal(sCurr, type);
          const ps = parseVal(sPrev, type);
          const ls = parseVal(sLast, type);
          const sDiffP = formatDiff(cs.val, ps.val, type, unit);
          const sDiffL = formatDiff(cs.val, ls.val, type, unit);
          const sDisp = formatNum(cs.val, type, unit);
          reportText += `  ・関西\n    └ ${sDisp} (前月: ${sDiffP} / 昨年: ${sDiffL})\n`;
        }
      }
    }
  });

  // ★修正：実績データの[info]を閉じ、レビュー用の[info]を作成。確実に末尾を[/info]で終わらせる。
  reportText += `\n[/info]\n\n[info]\n[title]🤖 レビュー[/title]\n(ここにAIのレビューが挿入されます)\n[/info]`;
  return reportText;
}
// ==========================================
// 2. Gemini API 呼び出し処理（構造化・トレンド分析・ドライなプロ目線版）
// ==========================================
function fetchAiReview(reportData) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Gemini APIキーが設定されていません。');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=${apiKey}`;
  
  const prompt = `あなたは優秀な医療法人のデータアナリストです。以下の小児科の月次実績データを分析し、現場や経営陣に提出するAIレビューを生成してください。

【小児科の季節性・前提知識（分析の前提として使用すること）】
・当法人は毎年拠点が拡大しているため、「稼働人員（医師数）」の総数が増加するのは当然の前提である。そのため、単なる人数の増加を成果としてもてはやすのではなく、構成比（直応募や紹介の割合）や効率性などの中身の変化を評価すること。
・当法人の繁忙期は「10月～1月」である。
・大型連休（GW、お盆、年末年始など）がある月は、常勤医師の稼働比率が低下し、スポット医師（募集枠）が増加する傾向がある。
・祝日がない（または少ない）月は、常勤医師の稼働比率が上がり、募集枠が減少する傾向がある。

【出力条件・トーン＆マナー】
・「依頼手当総額」に関する言及は絶対に行わないこと。
・「市場競争力を維持しつつ」「適正なコスト管理が〜」「最適化が功を奏し〜」「非常に安定した運営が〜」といった、定型的で無駄な装飾言葉（企業構文・過剰な称賛）は一切使用しないこと。
・感情や美辞麗句を排除し、事実に即してドライかつ鋭く分析し、データから読み取れる具体的なインサイトのみを端的に記述すること。
・箇条書きの「・」や、[title]などの装飾タグは自分で出力しないこと。
・以下の4つの見出し（【】を使用）を必ずそのまま用いて、各項目を2〜3文程度で記述すること。

【平均時給推移】
「エリア平均時給」の変化（前月比・昨年比）を事実ベースで記載し、その数値の変動要因（常勤・スポットの比率変化など）を端的に分析すること。無駄な称賛は不要。

【稼働人員】
拠点拡大に伴う単純な人数増ではなく、「構成比（直応募や紹介）」の推移にフォーカスし、採用チャネルの依存度や人員構成の変化を事実ベースで記述すること。

【２診時間】
「２診時間数」「医師不在時間」「医師残業時間」に着目し、シフトの充足状況や稼働効率の変化を客観的に指摘すること。

【トレンド】
データ内に記載されている「報告対象月」と、上記の「季節性・前提知識」を掛け合わせること。対象月の実績に対する要因分析（連休の影響など）と、次月に向けた予測（祝日の有無による常勤比率・募集枠の増減、繁忙期への備えなど）を具体的に記述すること。

【月次データ】
${reportData}
`;

  const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
  const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result.candidates && result.candidates.length > 0) {
      return result.candidates[0].content.parts[0].text;
    } else {
      throw new Error(response.getContentText());
    }
  } catch (e) {
    return `AIレビューの生成に失敗しました: ${e.message}`;
  }
}