/**
 * generate_grommet_xlsx.js
 * grommet_1300.js のデータを、既存の「海外便利グッズリスト」シートと同じ列構成のExcelファイルに書き出す
 * 列: No. / カテゴリ / 商品名 / メーカー名 / 販売ECサイト / 商品URL / メーカーHP / 連絡先メール / 日本市場需要 / 需要マッチコメント
 * node generate_grommet_xlsx.js
 */

const XLSX = require('xlsx');
const GROMMET_DATA = require('./grommet_1300.js');

const OUTPUT = "C:\\Users\\hiro\\Desktop\\LEAGUE\\Grommet_製品リスト_1300件_評価付.xlsx";

const HEADER = ['No.', 'カテゴリ', '商品名', 'メーカー名', '販売ECサイト', '商品URL', 'メーカーHP', '連絡先メール', '日本市場需要', '需要マッチコメント'];

// grommet_1300.js の行は [番号, カテゴリ, 製品名, メーカー, ECサイト, URL, メーカーHP, メール, ★, コメント] の並びなので
// そのままの列順で使える
const rows = [HEADER, ...GROMMET_DATA];

const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet(rows);

ws['!cols'] = [
  { wch: 6 }, { wch: 18 }, { wch: 40 }, { wch: 22 },
  { wch: 14 }, { wch: 52 }, { wch: 35 }, { wch: 35 },
  { wch: 12 }, { wch: 45 },
];

// ヘッダー行を太字に
for (let c = 0; c < HEADER.length; c++) {
  const addr = XLSX.utils.encode_cell({ r: 0, c });
  if (ws[addr]) ws[addr].s = { font: { bold: true } };
}

// 商品URL列(F)をハイパーリンク化
for (let r = 1; r < rows.length; r++) {
  const row = rows[r];
  const url = String(row[5] || '');
  if (url.startsWith('http')) {
    const addr = XLSX.utils.encode_cell({ r, c: 5 });
    ws[addr] = {
      v: url, t: 's',
      l: { Target: url },
      s: { font: { color: { rgb: '1155CC' }, underline: true } }
    };
  }
}

// メーカーHP列(G)をハイパーリンク化
for (let r = 1; r < rows.length; r++) {
  const row = rows[r];
  const hp = String(row[6] || '');
  if (hp.startsWith('http')) {
    const addr = XLSX.utils.encode_cell({ r, c: 6 });
    ws[addr] = {
      v: hp, t: 's',
      l: { Target: hp },
      s: { font: { color: { rgb: '1155CC' }, underline: true } }
    };
  }
}

// 連絡先メール列(H)をハイパーリンク化（mailto）
for (let r = 1; r < rows.length; r++) {
  const row = rows[r];
  const mail = String(row[7] || '');
  if (mail.includes('@')) {
    const addr = XLSX.utils.encode_cell({ r, c: 7 });
    ws[addr] = {
      v: mail, t: 's',
      l: { Target: `mailto:${mail}` },
      s: { font: { color: { rgb: '1155CC' }, underline: true } }
    };
  }
}

XLSX.utils.book_append_sheet(wb, ws, 'Grommetリスト');
XLSX.writeFile(wb, OUTPUT);

console.log(`完了: ${OUTPUT}`);
console.log(`件数: ${GROMMET_DATA.length} 件`);
