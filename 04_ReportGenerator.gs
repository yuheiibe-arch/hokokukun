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

  const targetMonthStr = baseSheet.getRange(1, 4).getDisplayValue() || '最新月';

  // [表示名, 行番号, 単位, 数値タイプ, カテゴリ記号]
  const metrics = [
    ['エリア平均時給', 2, '¥', 'money', '💰 時給・コスト指標'],
    ['エリア別詳細時給', 5, '¥', 'money', ''],
    ['依頼手当総額', 8, '¥', 'money', ''],
    ['稼働人員（全医師）', 12, '人', 'people', '👥 稼働・人員状況'],
    ['稼働人員（常勤除く）', 15, '人', 'people', ''],
    ['所定休出医師数', 18, '人', 'people', ''],
    ['構成比 (常/定/直/紹/休)', 21, '', 'text', ''], 
    ['２診時間数', 24, 'h', 'time', '🏥 シフト・拠点状況'],
    ['対象拠点数', 28, '拠点', 'num', ''],
    ['医師不在時間', 39, 'h', 'time', ''],
    ['医師残業時間', 42, 'h', 'time', ''],
    ['新規採用（直接）', 32, '人', 'people', '✨ 採用・在籍状況'],
    ['新規採用（紹介）', 35, '人', 'people', ''],
    ['在籍医師数（常勤）', 53, '人', 'people', ''],
    ['在籍医師数（定期）', 56, '人', 'people', ''],
    ['来院数 (目標達成率)', 46, '人', 'mix', '📈 業績実績'],
    ['売上 (目標達成率)', 49, '¥', 'mix', '']
  ];

  // ベースシートの表記から数値を正確に抽出するヘルパー
  const parseVal = (rawStr, type) => {
    if (!rawStr || rawStr === '-') return { val: 0, str: '-' };
    const str = String(rawStr).split(' ')[0]; // 増減 (+xxx) や達成率 (xx%) を除去
    
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
    if (type === 'money') return `¥${Math.round(num).toLocaleString()}`;
    if (type === 'time') return `${num.toFixed(1)}${unit}`;
    if (type === 'people' || type === 'num') return `${Math.round(num).toLocaleString()}${unit}`;
    return num;
  };

  const formatDiff = (curr, past, type, unit) => {
    const diff = curr - past;
    if (diff === 0) return '±0';
    const sign = diff > 0 ? '+' : '-';
    return `${sign}${formatNum(Math.abs(diff), type, unit)}`;
  };

  let reportText = `[info][title]📊 ${targetMonthStr} 月次実績報告（${targetArea}）[/title]\nお疲れ様です。最新の月次実績をご報告いたします。\n\n`;
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
      reportText += `・${name}\n  └ 当月実績: ${valCurrRaw.replace(/\n/g, ' / ')}\n`;
    } else {
      const c = parseVal(valCurrRaw, type);
      const p = parseVal(valPrevRaw, type);
      const l = parseVal(valLastRaw, type);

      const diffP = formatDiff(c.val, p.val, type, unit);
      const diffL = formatDiff(c.val, l.val, type, unit);
      
      let displayCurr = type === 'mix' ? valCurrRaw : formatNum(c.val, type, unit);

      reportText += `・${name}: ${displayCurr} (前月: ${diffP} / 昨年: ${diffL})\n`;
    }
  });

  reportText += `\n[hr]\n[title]🤖 AI分析レビュー[/title]\n(ここにAIのレビューが挿入されます)[/info]`;
  return reportText;
}

// ==========================================
// 2. Gemini API 呼び出し処理
// ==========================================
function fetchAiReview(reportData) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Gemini APIキーが設定されていません。');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent?key=${apiKey}`;
  
  const prompt = `あなたは医療法人の優秀な経営企画マネージャーです。以下の月次実績データを分析し、Chatworkで報告するためのAIレビューを生成してください。
条件：
・箇条書きで3点以内
・全体の文字数は200文字以内厳守
・専門用語を避け、端的に
・時給、採用、稼働状況の「前月比」「昨年比」の変化にフォーカスして洞察を述べること

データ：
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

// ==========================================
// 3. Chatwork への投稿処理
// ==========================================
function postToChatwork(message) {
  const cwToken = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_KEY');
  const roomId = '410981446'; // ご指定のテストルーム

  if (!cwToken) throw new Error('Chatwork APIキーが設定されていません。');

  const url = `https://api.chatwork.com/v2/rooms/${roomId}/messages`;
  const options = {
    "method": "post",
    "headers": { "X-ChatWorkToken": cwToken },
    "payload": { "body": message }
  };

  UrlFetchApp.fetch(url, options);
  return 'Chatworkへの投稿が完了しました！';
}