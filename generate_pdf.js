const puppeteer = require('puppeteer');
const path = require('path');

const files = [
  { html: 'housing_report_a3.html', pdf: 'housing_report.pdf',  label: '新築戸建て業界レポート' },
  { html: 'reform_report_a3.html',  pdf: 'reform_report.pdf',   label: 'リフォーム業界レポート' },
  { html: 'compare_report.html',    pdf: 'compare_report.pdf',  label: '業界対比レポート' },
];

// PDF マージンを 0 にし、CSS @page の 8mm マージンも上書きする。
// .page の padding (14mm 上 + 12mm 下) が視覚的な余白になる。
// これで .page(min-height:297mm) が PDF の A4 1 ページ(297mm)にぴったり収まる。
const MARGIN_RESET_CSS = `
  @page { margin: 0 !important; }
`;

const BASE = path.resolve(__dirname);

(async () => {
  const browser = await puppeteer.launch({
    executablePath: 'C:/Program Files/Google/Chrome/Application/chrome.exe',
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  for (const f of files) {
    console.log(`生成中: ${f.label} ...`);
    const page = await browser.newPage();

    const url = `file:///${BASE.replace(/\\/g, '/')}/${f.html}`;
    await page.goto(url, { waitUntil: 'networkidle0', timeout: 60000 });

    // @page マージンをゼロにリセット（二重マージン防止）
    await page.addStyleTag({ content: MARGIN_RESET_CSS });

    // Chart.js の描画完了を待機
    await new Promise(r => setTimeout(r, 2500));

    await page.pdf({
      path: path.join(BASE, f.pdf),
      format: 'A4',
      printBackground: true,
      // マージンは 0。視覚的余白は HTML 側の .page padding で確保
      margin: { top: '0', bottom: '0', left: '0', right: '0' },
    });

    await page.close();
    console.log(`  → ${f.pdf} 完了`);
  }

  await browser.close();
  console.log('\n✅ 全PDFの生成が完了しました。');
})();
