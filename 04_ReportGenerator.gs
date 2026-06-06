// ==========================================
// 1. レポートテキストの自動生成
// ==========================================
function generateReportText(targetArea = 'グループ全体') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = ss.getSheetByName('ベースシート');
  if (!baseSheet) throw new Error('ベースシートが見つかりません。');

  const bHeaders = baseSheet.getRange(1, 1, 1, baseSheet.getLastColumn()).getValues()[0];
  let colIdx = -1;
  for (let c = 3; c < bHeaders.length; c++) {
    if (String(bHeaders[c]).trim() === targetArea) { colIdx = c + 1; break; }
  }
  if (colIdx === -1) throw new Error(`${targetArea}の列が見つかりません。`);

  // ★修正：システム日付から「先月」を自動で計算する
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
    { key: 'エリア別詳細時給', displayName: 'エリア別詳細時給', unit: '¥', type: 'money', cat: '', exclude: [] },
    { key: '依頼手当', displayName: '依頼手当総額', unit: '¥', type: 'money', cat: '', exclude: [] },
    
    { key: '稼働人員（全医師', displayName: '稼働人員（全医師）', unit: '人', type: 'people', cat: '👥 稼働・人員状況', exclude: [] },
    { key: '稼働人員（常勤除く', displayName: '稼働人員（常勤除く）', unit: '人', type: 'people', cat: '', exclude: [] },
    { key: '所定休出医師数', displayName: '所定休出医師数', unit: '人', type: 'people', cat: '', exclude: [] },
    { key: '構成比', displayName: '構成比 (常勤/定非/直応募/紹介/休出)', unit: '', type: 'text', cat: '', exclude: [] },
    
    { key: '２診時間', displayName: '２診時間数', unit: 'h', type: 'time', cat: '🏥 シフト・拠点状況', exclude: [] },
    { key: '2診拠点数', displayName: '２診拠点数', unit: '拠点', type: 'num', cat: '', exclude: [] },
    { key: '２診拠点数', displayName: '２診拠点数', unit: '拠点', type: 'num', cat: '', exclude: [] }, 
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
    { key: '売上(目標達成率', displayName: '売上 (目標達成率)', unit: '¥', type: 'mix-money', cat: '', exclude: [] }, 
    
    // ★修正：「小児科」を追記
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

  // ★修正：小児科と月の間にスペースを追加
  let reportText = `[info][title]月次実績報告（${targetArea}）[/title]\nお疲れ様です。小児科 ${monthDisplay}の月次実績をご報告いたします。\n\n`;
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
    }
  });

  reportText += `\n[hr]\n[title]🤖 AI分析レビュー[/title]\n(ここにAIのレビューが挿入されます)[/info]`;
  return reportText;
}
// ==========================================
// 2. Gemini API 呼び出し処理（秀逸なアナリスト版に改修）
// ==========================================
function fetchAiReview(reportData) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Gemini APIキーが設定されていません。');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=${apiKey}`;
  
  const prompt = `あなたは優秀な医療法人のデータアナリストです。以下の月次実績データを分析し、現場のモチベーションを高めつつ、経営陣にも刺さる秀逸なAIレビューを生成してください。

【出力条件】
・見出しは使用せず、箇条書き（・）で3点にまとめること。
・「依頼手当総額」に関する言及は絶対に行わないこと。
・単なる数字の羅列を避け、その数値が意味する「背景」や「トレンドの変化」を言語化して、読みやすく洗練された文章にすること。
・以下の構成で必ず3点を記述すること：
 1点目：必ず「エリア平均時給」の変化にフォーカスした実績の振り返り。
 2点目：「稼働人員（全医師）」「構成比（特に紹介会社や直応募）」の実績にフォーカスした振り返り。
 3点目：「２診時間数」「医師不在時間」「医師残業時間」などのシフト調整・拠点状況に関する指標に着目し、急激な増加や特筆すべきトレンド（波及効果など）があれば必ず指摘すること。また、それを踏まえた「次月の注目項目」を述べること（※提言や解決策は不要）。
・数値を引用する際は、「前月比」と「昨年比」を比較し、変化幅（インパクト）がより大きい方を積極的に採用して言及すること。
・ネガティブな表現（例：「圧迫している」「悪化した」等）は避け、事実を客観的かつ前向きなトーンで記述すること。
・全体の文字数は300〜400文字程度で、インサイト（洞察）のある内容にすること。
・[title]などの装飾タグは自分で絶対に出力しないこと。

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