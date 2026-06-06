function generateRequirementsPDF() {
  console.log('🚀 仕様書のPDFを生成しています...');

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        body { font-family: 'Noto Sans JP', sans-serif; font-size: 10px; line-height: 1.4; color: #333; }
        h1 { font-size: 16px; text-align: center; border-bottom: 2px solid #333; padding-bottom: 5px; }
        h2 { font-size: 12px; background-color: #f0f0f0; padding: 4px; margin-top: 15px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
        th, td { border: 1px solid #999; padding: 4px; text-align: left; vertical-align: top; }
        th { background-color: #e6e6e6; font-weight: bold; width: 18%; }
        .col-c { width: 12%; }
        .col-logic { width: 40%; }
        .col-source { width: 30%; }
        /* ★追加：注釈用のスタイル */
        .notes-box { border: 1px solid #ccc; padding: 8px 12px; background-color: #fafafa; margin-top: 5px; }
        .notes-box ul { padding-left: 18px; margin: 0; }
        .notes-box li { margin-bottom: 6px; }
      </style>
    </head>
    <body>
      <h1>月次報告自動化スクリプト要件定義書（最終版）</h1>
      <p style="text-align: right;">作成日: 2026年6月6日<br>ステータス: 最終承認（全項目ファクトチェック済）</p>
      
      <p>ベースシートに出力される全項目の計算ロジックと、データ取得元の定義です。</p>

      <h2>■ 1. 時給関係</h2>
      <table>
        <tr><th>中分類(B列)</th><th class="col-c">比較軸(C列)</th><th class="col-logic">計算ロジック(スクリプト処理)</th><th class="col-source">データソース(取得元)</th></tr>
        <tr><td>エリア平均時給</td><td>対象月/前月/昨年</td><td>実際に稼働した全医師の「支払給与合計」を「総勤務時間」で割り戻して算出。</td><td>『確定シフト』<br>時給合計、開始/終了</td></tr>
        <tr><td>エリア別詳細時給</td><td>対象月/前月/昨年</td><td>各エリアごとの詳細な平均時給。</td><td>同上</td></tr>
        <tr><td>依頼手当総額</td><td>対象月/前月/昨年</td><td>特別な勤務依頼等で発生した「依頼手当」の支給総額を合算。</td><td>『確定シフト』<br>追加支給項目・額</td></tr>
      </table>

      <h2>■ 2. 稼働人数・2診状況</h2>
      <table>
        <tr><th>中分類(B列)</th><th class="col-c">比較軸(C列)</th><th class="col-logic">計算ロジック(スクリプト処理)</th><th class="col-source">データソース(取得元)</th></tr>
        <tr><td>稼働人員（全医師）</td><td>対象月/前月/昨年</td><td>対象月に1回以上勤務した医師の実人数（ユニーク数）をカウント。</td><td>『確定シフト』<br>氏名、医籍番号</td></tr>
        <tr><td>稼働人員（常勤除く）</td><td>対象月/前月/昨年</td><td>稼働した全医師から「常勤(所定勤務)」を引いたユニーク人数。</td><td>『確定シフト』<br>＋『医師情報マスタ』</td></tr>
        <tr><td>所定休出医師数</td><td>対象月/前月/昨年</td><td>コメント等に「所定休出」等の記載があり、休日に出勤した医師の実人数。</td><td>『確定シフト』<br>スタッフコメント</td></tr>
        <tr><td>構成比</td><td>対象月/前月/昨年</td><td>各勤務区分（常勤/定期/直応募/紹介/休出）が稼働人数UUに占める割合(%)。</td><td>『確定シフト』等</td></tr>
        <tr><td>２診時間数</td><td>対象月/前月/昨年</td><td>1つの拠点内で、複数の医師の勤務時間が重複した合計時間。</td><td>『確定シフト』<br>＋『休館日シート』</td></tr>
        <tr><td>２診拠点数</td><td>対象月/前月/昨年</td><td>1日に「180分以上」の2診が発生した日が、「月に4日以上」あるレギュラー2診拠点数。</td><td>同上</td></tr>
      </table>

      <h2>■ 3. 拠点数推移・採用人数</h2>
      <table>
        <tr><th>中分類(B列)</th><th class="col-c">比較軸(C列)</th><th class="col-logic">計算ロジック(スクリプト処理)</th><th class="col-source">データソース(取得元)</th></tr>
        <tr><td>対象拠点数</td><td>対象月/前月/昨年</td><td>その月の月末時点で開院日を迎えており、稼働している拠点の総数。</td><td>『正規表現マスタ』<br>開院日</td></tr>
        <tr><td>新規採用（直接）</td><td>対象月/前月/昨年</td><td>小児科のスポット医師のうち、直接応募で採用可となった人数。</td><td>『チェックリスト』</td></tr>
        <tr><td>新規採用（紹介）</td><td>対象月/前月/昨年</td><td>小児科のスポット医師のうち、紹介会社経由で採用可となった人数。</td><td>『チェックリスト』</td></tr>
      </table>

      <h2>■ 4. シフト調整関係</h2>
      <table>
        <tr><th>中分類(B列)</th><th class="col-c">比較軸(C列)</th><th class="col-logic">計算ロジック(スクリプト処理)</th><th class="col-source">データソース(取得元)</th></tr>
        <tr><td>医師勤務時間</td><td>対象月/前月/昨年</td><td>全医師が勤務した総時間の合算。</td><td>『確定シフト』</td></tr>
        <tr><td>勤務時間（固定）</td><td>対象月/前月/昨年</td><td>総勤務時間のうち「常勤」「定期非常勤」の医師が勤務した時間数。</td><td>『確定シフト』等</td></tr>
        <tr><td>勤務時間（募集）</td><td>対象月/前月/昨年</td><td>総勤務時間のうち「スポット」「紹介会社」「休出」の医師が勤務した時間数。</td><td>同上</td></tr>
        <tr><td>医師不在時間</td><td>対象月/前月/昨年</td><td>拠点の基本営業時間内で、誰もシフトに入っていない（穴が空いた）時間数。</td><td>『確定シフト』<br>＋『休館日シート』</td></tr>
        <tr><td>医師残業時間</td><td>対象月/前月/昨年</td><td>13時超過、および基本営業終了時間（20時/21時等）を超過して勤務した時間。</td><td>同上</td></tr>
      </table>

      <h2>■ 5 & 6. 売上・在籍医師数</h2>
      <table>
        <tr><th>中分類(B列)</th><th class="col-c">比較軸(C列)</th><th class="col-logic">計算ロジック(スクリプト処理)</th><th class="col-source">データソース(取得元)</th></tr>
        <tr><td>来院数（達成率）</td><td>対象月/前月/昨年</td><td>実績来院数と、予算（目標）に対する達成率。</td><td>『業績管理シート』</td></tr>
        <tr><td>売上（達成率）</td><td>対象月/前月/昨年</td><td>実績売上金額と、予算（目標）に対する達成率。</td><td>『業績管理シート』</td></tr>
        <tr><td>在籍医師数（常勤）</td><td>対象月/前月/昨年</td><td>対象月に在籍しており、かつ実際にシフトに入った常勤医師の実人数。</td><td>『確定シフト』<br>＋『常勤マスタ』</td></tr>
        <tr><td>在籍数（定期非常勤）</td><td>対象月/前月/昨年</td><td>対象月に在籍しており、かつ実際にシフトに入った定期非常勤の実人数。</td><td>『確定シフト』<br>＋『定期マスタ』</td></tr>
      </table>

      <h2>■ 7. 【重要】集計に関する特記事項・前提条件</h2>
      <div class="notes-box">
        <ul>
          <li><strong>時給・コスト算出の除外項目:</strong> エリア平均時給などの算出には、「紹介会社への紹介手数料（一律20%）」および「インセンティブ支給額」は含まれていません。</li>
          <li><strong>常勤医師の時給除外:</strong> エリア平均時給の算出において、常勤医師の給与・労働時間は計算母数から除外しています。</li>
          <li><strong>内科実績の除外:</strong> 本報告は「小児科単体」の実績となります。内科稼働分（2026年6月現在、常勤2名・定期非常勤27名）はすべての集計から除外しています。</li>
          <li><strong>依頼手当の増加要因:</strong> 依頼手当が昨年と比較して大きく増額していますが、これは昨年10月より導入された「常勤医師追加勤務手当制度」が大きく寄与しているためです。</li>
          <li><strong>２診時間の算出定義:</strong> 確定シフトをベースに、各拠点の基本営業時間（09:00〜21:00 ※北葛西のみ20:00）を基準としています。1拠点あたりの総勤務時間から休診時間（2時間）を差し引いた上で、重複する稼働時間を「2診時間」として算出しています。</li>
          <li><strong>データ抽出のタイミングによる差異:</strong> 各種実績は「確定シフト」の最新データを用いて算出していますが、残業時間などはデータ保存・取得のタイミングにより、最終的な確定給与明細等と微細な差異が生じる場合があります。</li>
        </ul>
      </div>
    </body>
    </html>
  `;

  try {
    const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
    blob.setName('月次報告自動化スクリプト要件定義書_最終版.pdf');
    const file = DriveApp.createFile(blob);
    console.log('🎉 PDFの生成が完了しました！以下のURLからダウンロードできます。');
    console.log(file.getUrl());
  } catch (e) {
    console.log('❌ エラーが発生しました: ' + e.message);
  }
}