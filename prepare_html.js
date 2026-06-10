/**
 * 各HTMLファイルにコンパクトCSSを追加し、グラフ高さを縮小する。
 * その後 generate_pdf.js で PDF を再生成する。
 */
const fs = require('fs');

const COMPACT_CSS = `
/* ===== A4 COMPACT OVERRIDE (自動追加) ===== */
body { font-size: 10px !important; line-height: 1.56 !important; }

/* ページ余白 */
.page { padding: 9mm 11mm 8mm !important; }
.page-header { margin-bottom: 3mm !important; padding-bottom: 1.5mm !important; }
.ph-title { font-size: 14px !important; }
.ph-chapter { font-size: 7px !important; margin-bottom: 0 !important; }
.ph-page { font-size: 8px !important; }

/* インサイトボックス */
.insight { padding: 2mm 3mm !important; margin-bottom: 2.5mm !important; }
.insight .ins-head { font-size: 10px !important; margin-bottom: 0.8mm !important; }

/* セクションタイトル */
.section-title { font-size: 10.5px !important; margin-bottom: 1.5mm !important; }

/* NVF / CMP */
.nvf-wrap, .cmp-row { gap: 2mm !important; margin-bottom: 2.5mm !important; }
.nvf-box, .cmp-box { padding: 2mm 2.5mm !important; }
.nvf-box .nvf-head, .cmp-box .cb-head { font-size: 9.5px !important; margin-bottom: 1mm !important; }
.nvf-box li, .cmp-box li { margin-bottom: 0.2mm !important; }
.nvf-box ul, .cmp-box ul { padding-left: 3mm !important; }

/* グリッド */
.two-col, .three-col { gap: 2mm !important; margin-bottom: 2.5mm !important; }
.strat-grid { gap: 2mm !important; margin-bottom: 2.5mm !important; }
.scenario-3 { gap: 2mm !important; margin-bottom: 2.5mm !important; }
.matrix { gap: 2mm !important; margin-bottom: 2.5mm !important; }
.driver-grid { gap: 1.5mm !important; margin-bottom: 2.5mm !important; }
.hl-num-row, .hl-row { gap: 2mm !important; margin-bottom: 2.5mm !important; }

/* チャートボックス */
.chart-box { padding: 1.5mm 2mm !important; margin-bottom: 2mm !important; }
.chart-label { margin-bottom: 0.8mm !important; font-size: 7.5px !important; }

/* テーブル */
.data-table { margin-bottom: 2.5mm !important; font-size: 9.5px !important; }
.data-table th { padding: 1.2mm 1.8mm !important; font-size: 8.5px !important; }
.data-table td { padding: 1.2mm 1.8mm !important; }

/* 数値ハイライト */
.hl-num, .hn { padding: 1.8mm 2mm !important; }
.hl-num .hn-val { font-size: 15px !important; }
.hl-num .hn-label, .hl-num .hn-l { font-size: 7.5px !important; }
.hl-num .hn-sub, .hl-num .hn-s { font-size: 7.5px !important; margin-top: 0.5mm !important; }

/* 戦略カード */
.strat-card .sc-head { padding: 1.5mm 2mm !important; font-size: 9px !important; }
.strat-card .sc-body { padding: 2mm !important; }
.strat-card .sc-body li { margin-bottom: 0.2mm !important; }
.strat-card .sc-body ul { padding-left: 3mm !important; }

/* シナリオ・マトリクス・ドライバー */
.sc3-box { padding: 1.8mm 2.5mm !important; }
.sc3-box h4 { font-size: 9.5px !important; margin-bottom: 1mm !important; }
.sc3-box li { margin-bottom: 0.2mm !important; }
.sc3-box ul { padding-left: 3mm !important; }
.matrix-cell { padding: 2mm 2.5mm !important; }
.matrix-cell h4 { font-size: 9.5px !important; margin-bottom: 1mm !important; }
.matrix-cell li { margin-bottom: 0.2mm !important; }
.matrix-cell ul { padding-left: 3mm !important; }
.driver-item { padding: 1.5mm 2mm !important; }
.driver-item strong { font-size: 9.5px !important; margin-bottom: 0.2mm !important; }

/* タイムライン */
.timeline-row { margin-bottom: 2mm !important; }
.tl-item { font-size: 7.5px !important; padding: 1.5mm 1mm !important; }
.tl-item .tl-year { font-size: 9px !important; }
.tl-item::after {
  border-top-width: 14px !important;
  border-bottom-width: 14px !important;
  right: -5px !important;
}

/* カバーページ */
.cover { padding: 12mm 12mm !important; }
.cover-title { font-size: 21px !important; margin-bottom: 2.5mm !important; }
.cover-subtitle { font-size: 11px !important; margin-bottom: 5mm !important; }
.cover-kpi-row { gap: 2mm !important; margin-top: 4mm !important; }
.cover-kpi { padding: 2mm 3mm !important; }
.cover-kpi .ck-val { font-size: 14px !important; }
.cover-kpi .ck-label, .cover-kpi .ck-sub { font-size: 7px !important; }
.cover-split { gap: 3mm !important; margin-bottom: 4mm !important; }
.cover-half { padding: 3mm !important; }
.cover-half h3 { font-size: 11px !important; margin-bottom: 1.5mm !important; }
.cover-half .cv-kpi { font-size: 10px !important; margin-bottom: 1mm !important; }
.cover-half .cv-kpi span { font-size: 16px !important; }
.cover-toc { gap: 2mm 4mm !important; }
.cover-toc-item { font-size: 9px !important; }
.cover-msg { padding: 2.5mm 3mm !important; margin-bottom: 3mm !important; font-size: 10px !important; }
.cover-divider { margin: 3mm 0 !important; }
.cover-eyebrow { margin-bottom: 2.5mm !important; }

/* 対比レポート固有 */
.tl-both { margin-bottom: 2mm !important; }
.tl-shin, .tl-ref { padding: 1.5mm 2mm !important; font-size: 8.5px !important; }
.tl-year-cell { font-size: 8.5px !important; padding: 1mm !important; }
.tl-content-col { gap: 1.5mm !important; }
.tl-row { gap: 1.5mm !important; }

/* フッター */
.page-footer { font-size: 7px !important; padding-top: 1mm !important; }
`;

const files = [
  'housing_report_a3.html',
  'reform_report_a3.html',
  'compare_report.html',
];

const CHART_SCALE = 0.78; // グラフ高さを 78% に縮小

files.forEach(filename => {
  let html = fs.readFileSync(filename, 'utf8');

  // 既に追加済みのオーバーライドブロックを除去（べき等性）
  html = html.replace(/\n\/\* ={5} A4 COMPACT OVERRIDE[\s\S]*?={5} \*\/[\s\S]*?(?=\n<\/style>)/, '');

  // コンパクト CSS を </style> の直前に挿入
  html = html.replace('</style>', COMPACT_CSS + '\n</style>');

  // chart-wrap の inline height を縮小
  html = html.replace(
    /(<div class="chart-wrap" style="height:)(\d+)(px">)/g,
    (m, pre, h, post) => pre + Math.round(parseInt(h) * CHART_SCALE) + post
  );

  fs.writeFileSync(filename, html, 'utf8');
  console.log(`✓ ${filename} を更新しました`);
});

console.log('\n準備完了。node generate_pdf.js を実行してください。');
