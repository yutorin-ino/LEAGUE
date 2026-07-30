/**
 * generate_summary_xlsx.js
 * 製品リスト一覧スプレッドシート生成
 * node generate_summary_xlsx.js
 */

const XLSX = require('xlsx');
const path = require('path');

const INPUT_FILE  = path.join(__dirname, '海外便利グッズリスト_日本未上陸3105件_評価付.xlsx');
const OUTPUT_FILE = path.join(__dirname, 'LEAGUE_製品リスト一覧.xlsx');
const SENDER      = 'yutorin.ino@gmail.com';

function buildGmailUrl(email, maker, product) {
  const subject = 'Potential Distribution Partnership for Japan';
  const body = `Dear ${maker} Team,

My name is Hiroyuki Inoguchi from
Sumai pluS Co., Ltd. in Japan

I hope this message finds you all well.

Our company focuses on promoting products that enrich people's daily lives. We are currently seeking unique international brands to introduce to the Japanese market and support their growth.

I recently had the opportunity to review your product (${product}) and was very impressed by its potential in the Japanese market.

I am confident that your product will strongly appeal to Japanese customers, and I would like to explore the possibility of building a partnership with your company to help ensure its success.

Furthermore, our company has a proven track record of promoting overseas products through our proprietary sales network and Japanese distribution channels.

Would it be possible to schedule a brief online meeting sometime next week to discuss this matter?

I look forward to hearing from you.


Sincerely,

Hiroyuki Inoguchi
Sumai pluS Co., Ltd.
 (LEAGUE Co., Ltd. Agent )

================================================

Sumai pluS Co., Ltd.
 (LEAGUE Co., Ltd. Agent )
Hiroyuki Inoguchi

E-mail:${SENDER}

Address: 49-5 Kitazakuno, Higashi-Itsushiro, Ichinomiya City, Aichi Prefecture, Japan

=================================================`;
  const enc = (s) => encodeURIComponent(s);
  return `https://mail.google.com/mail/?authuser=${enc(SENDER)}&view=cm&fs=1&to=${enc(email)}&su=${enc(subject)}&body=${enc(body)}`;
}

// ===== データ読み込み =====
const wb_in = XLSX.readFile(INPUT_FILE);
const ws_in = wb_in.Sheets[wb_in.SheetNames[0]];
const rows  = XLSX.utils.sheet_to_json(ws_in, { header: 1, defval: '' });

// ヘッダー行をスキップ（1行目がヘッダーの場合）
// データ行（番号が数値のもの）だけ抽出
const data = rows.filter(r => typeof r[0] === 'number' || (r[0] && !isNaN(Number(r[0])) && Number(r[0]) > 0));

// フィールド定義
const COL = { NO:0, CAT:1, NAME:2, MAKER:3, EC:4, URL:5, HP:6, MAIL:7, STAR:8, COMMENT:9 };

const CF_KEYWORDS   = ['Kickstarter','Indiegogo','BackerKit','Crowdfunder','Crowd Supply'];
const SHOP_KEYWORDS = ['Amazon','eBay','Walmart','Target','Best Buy','Costco','Shopify'];

function getSiteType(ec) {
  for (const kw of CF_KEYWORDS)   if (String(ec).includes(kw)) return 'クラファン';
  for (const kw of SHOP_KEYWORDS) if (String(ec).includes(kw)) return '通販';
  return 'その他';
}

function hasMail(mail) {
  const m = String(mail);
  return m.includes('@') && !m.startsWith('http');
}

// ===== スタイルヘルパー =====
const NAVY   = '1A2D5A';
const GOLD   = 'C8A96E';
const WHITE  = 'FFFFFF';
const LGRAY  = 'F5F5F5';
const DGRAY  = 'DDDDDD';
const GREEN  = 'E6F4EA';
const ORANGE = 'FFF3E0';
const BLUE   = 'E8F0FE';

function hStyle(bgColor, fontColor, bold, sz) {
  return {
    font:  { bold: !!bold, color: { rgb: fontColor || WHITE }, sz: sz || 10, name: 'Meiryo UI' },
    fill:  { fgColor: { rgb: bgColor } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: {
      top:    { style: 'thin', color: { rgb: DGRAY } },
      bottom: { style: 'thin', color: { rgb: DGRAY } },
      left:   { style: 'thin', color: { rgb: DGRAY } },
      right:  { style: 'thin', color: { rgb: DGRAY } },
    }
  };
}
function dStyle(bgColor, center) {
  return {
    font:  { sz: 9, name: 'Meiryo UI', color: { rgb: '333333' } },
    fill:  bgColor ? { fgColor: { rgb: bgColor } } : undefined,
    alignment: { horizontal: center ? 'center' : 'left', vertical: 'center', wrapText: true },
    border: {
      top:    { style: 'hair', color: { rgb: DGRAY } },
      bottom: { style: 'hair', color: { rgb: DGRAY } },
      left:   { style: 'hair', color: { rgb: DGRAY } },
      right:  { style: 'hair', color: { rgb: DGRAY } },
    }
  };
}

function setCell(ws, r, c, v, style) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const t = typeof v === 'number' ? 'n' : 's';
  ws[addr] = { v, t, s: style };
}

// ===== Sheet1: 全件リスト =====
function buildSheet1(data) {
  const ws = {};
  const headers = ['No.','カテゴリ','製品名','メーカー','ECサイト','製品URL','メーカーHP','メール','★評価','コメント','サイト種別','メールあり'];
  const colW = [6,20,38,22,16,32,28,28,8,50,10,10];

  // ヘッダー
  headers.forEach((h, c) => setCell(ws, 0, c, h, hStyle(NAVY, WHITE, true, 10)));

  // データ行
  data.forEach((row, i) => {
    const r = i + 1;
    const bg = i % 2 === 0 ? WHITE : LGRAY;
    const siteType = getSiteType(row[COL.EC]);
    const mailOk   = hasMail(row[COL.MAIL]);

    setCell(ws, r, 0,  Number(row[COL.NO]) || 0,  dStyle(bg, true));
    setCell(ws, r, 1,  String(row[COL.CAT]   || ''), dStyle(bg));
    setCell(ws, r, 2,  String(row[COL.NAME]  || ''), dStyle(bg));
    setCell(ws, r, 3,  String(row[COL.MAKER] || ''), dStyle(bg));
    setCell(ws, r, 4,  String(row[COL.EC]    || ''), dStyle(bg, true));
    setCell(ws, r, 5,  String(row[COL.URL]   || ''), dStyle(bg));
    setCell(ws, r, 6,  String(row[COL.HP]    || ''), dStyle(bg));
    // メール（元データのセルvalueを直接）
    const mailRaw = ws_in[XLSX.utils.encode_cell({ r: i + 1, c: COL.MAIL })];
    // メール列：プレーンテキスト（VBAがダブルクリックでChromeを起動）
    const mailAddr = String(row[COL.MAIL] || '');
    setCell(ws, r, 7, mailAddr, mailOk
      ? { ...dStyle(bg), font: { sz:9, name:'Meiryo UI', color:{ rgb:'1155CC' }, underline: true } }
      : dStyle(bg));
    setCell(ws, r, 8,  String(row[COL.STAR]  || ''), dStyle(bg, true));
    setCell(ws, r, 9,  String(row[COL.COMMENT]|| ''), dStyle(bg));
    setCell(ws, r, 10, siteType, dStyle(
      siteType === 'クラファン' ? 'FFE0B2' : siteType === '通販' ? 'E8F5E9' : LGRAY, true
    ));
    setCell(ws, r, 11, mailOk ? '✓' : '', dStyle(mailOk ? 'E8F5E9' : bg, true));
    // 隠し列 M (index 12)：Gmail完全URL（VBAが参照）
    if (mailOk) {
      const gmailUrl = buildGmailUrl(mailAddr, String(row[COL.MAKER]||''), String(row[COL.NAME]||''));
      setCell(ws, r, 12, gmailUrl, { font:{ sz:8 }, alignment:{ horizontal:'left' } });
    }
  });

  ws['!ref'] = XLSX.utils.encode_range({ r: 0, c: 0 }, { r: data.length, c: 11 });
  ws['!cols'] = colW.map(w => ({ wch: w }));
  ws['!rows'] = [{ hpt: 22 }];
  ws['!freeze'] = { xSplit: 0, ySplit: 1 };
  return ws;
}

// ===== Sheet2: カテゴリ別集計 =====
function buildSheet2(data) {
  const ws = {};

  // カテゴリ集計
  const catMap = {};
  data.forEach(row => {
    const cat      = String(row[COL.CAT] || 'その他');
    const siteType = getSiteType(row[COL.EC]);
    const mailOk   = hasMail(row[COL.MAIL]);
    if (!catMap[cat]) catMap[cat] = { total:0, cf:0, shop:0, other:0, mail:0 };
    catMap[cat].total++;
    if (siteType === 'クラファン') catMap[cat].cf++;
    else if (siteType === '通販')  catMap[cat].shop++;
    else                           catMap[cat].other++;
    if (mailOk) catMap[cat].mail++;
  });

  const headers = ['カテゴリ','総件数','クラファン','通販','その他','メールあり','メール率'];
  headers.forEach((h, c) => setCell(ws, 0, c, h, hStyle(NAVY, WHITE, true, 10)));

  const cats = Object.keys(catMap).sort();
  cats.forEach((cat, i) => {
    const r  = i + 1;
    const d  = catMap[cat];
    const bg = i % 2 === 0 ? WHITE : LGRAY;
    const rate = d.total > 0 ? Math.round(d.mail / d.total * 100) + '%' : '0%';
    setCell(ws, r, 0, cat,      dStyle(bg));
    setCell(ws, r, 1, d.total,  dStyle(bg, true));
    setCell(ws, r, 2, d.cf,     dStyle(d.cf   > 0 ? 'FFE0B2' : bg, true));
    setCell(ws, r, 3, d.shop,   dStyle(d.shop > 0 ? 'E8F5E9' : bg, true));
    setCell(ws, r, 4, d.other,  dStyle(bg, true));
    setCell(ws, r, 5, d.mail,   dStyle(bg, true));
    setCell(ws, r, 6, rate,     dStyle(bg, true));
  });

  // 合計行
  const total   = data.length;
  const cfTotal = data.filter(r => getSiteType(r[COL.EC]) === 'クラファン').length;
  const shTotal = data.filter(r => getSiteType(r[COL.EC]) === '通販').length;
  const otTotal = data.filter(r => getSiteType(r[COL.EC]) === 'その他').length;
  const mailTot = data.filter(r => hasMail(r[COL.MAIL])).length;
  const tr = cats.length + 1;
  setCell(ws, tr, 0, '合計',                      hStyle(GOLD, NAVY, true));
  setCell(ws, tr, 1, total,                        hStyle(GOLD, NAVY, true));
  setCell(ws, tr, 2, cfTotal,                      hStyle(GOLD, NAVY, true));
  setCell(ws, tr, 3, shTotal,                      hStyle(GOLD, NAVY, true));
  setCell(ws, tr, 4, otTotal,                      hStyle(GOLD, NAVY, true));
  setCell(ws, tr, 5, mailTot,                      hStyle(GOLD, NAVY, true));
  setCell(ws, tr, 6, Math.round(mailTot/total*100)+'%', hStyle(GOLD, NAVY, true));

  ws['!ref']  = XLSX.utils.encode_range({ r:0, c:0 }, { r: tr, c: 6 });
  ws['!cols'] = [{ wch:24 },{ wch:8 },{ wch:12 },{ wch:8 },{ wch:8 },{ wch:10 },{ wch:8 }];
  return ws;
}

// ===== Sheet3: ECサイト別集計 =====
function buildSheet3(data) {
  const ws = {};

  const ecMap = {};
  data.forEach(row => {
    const ec   = String(row[COL.EC] || 'その他');
    const mail = hasMail(row[COL.MAIL]);
    if (!ecMap[ec]) ecMap[ec] = { total:0, mail:0 };
    ecMap[ec].total++;
    if (mail) ecMap[ec].mail++;
  });

  const headers = ['ECサイト','件数','メールあり','メール率','種別'];
  headers.forEach((h, c) => setCell(ws, 0, c, h, hStyle(NAVY, WHITE, true, 10)));

  const sorted = Object.entries(ecMap).sort((a,b) => b[1].total - a[1].total);
  sorted.forEach(([ec, d], i) => {
    const r    = i + 1;
    const bg   = i % 2 === 0 ? WHITE : LGRAY;
    const rate = d.total > 0 ? Math.round(d.mail/d.total*100)+'%' : '0%';
    const type = getSiteType(ec);
    setCell(ws, r, 0, ec,      dStyle(bg));
    setCell(ws, r, 1, d.total, dStyle(bg, true));
    setCell(ws, r, 2, d.mail,  dStyle(bg, true));
    setCell(ws, r, 3, rate,    dStyle(bg, true));
    setCell(ws, r, 4, type,    dStyle(
      type === 'クラファン' ? 'FFE0B2' : type === '通販' ? 'E8F5E9' : LGRAY, true
    ));
  });

  ws['!ref']  = XLSX.utils.encode_range({ r:0, c:0 }, { r: sorted.length, c: 4 });
  ws['!cols'] = [{ wch:18 },{ wch:8 },{ wch:10 },{ wch:8 },{ wch:12 }];
  return ws;
}

// ===== Sheet4: ★評価別集計 =====
function buildSheet4(data) {
  const ws = {};

  const starMap = {};
  data.forEach(row => {
    const star = String(row[COL.STAR] || '未評価');
    const mail = hasMail(row[COL.MAIL]);
    if (!starMap[star]) starMap[star] = { total:0, mail:0 };
    starMap[star].total++;
    if (mail) starMap[star].mail++;
  });

  const headers = ['★評価','件数','メールあり','メール率'];
  headers.forEach((h, c) => setCell(ws, 0, c, h, hStyle(NAVY, WHITE, true, 10)));

  const sorted = Object.entries(starMap).sort((a,b) => b[0].localeCompare(a[0]));
  sorted.forEach(([star, d], i) => {
    const r    = i + 1;
    const bg   = i % 2 === 0 ? WHITE : LGRAY;
    const rate = d.total > 0 ? Math.round(d.mail/d.total*100)+'%' : '0%';
    setCell(ws, r, 0, star,    dStyle(bg));
    setCell(ws, r, 1, d.total, dStyle(bg, true));
    setCell(ws, r, 2, d.mail,  dStyle(bg, true));
    setCell(ws, r, 3, rate,    dStyle(bg, true));
  });

  ws['!ref']  = XLSX.utils.encode_range({ r:0, c:0 }, { r: sorted.length, c: 3 });
  ws['!cols'] = [{ wch:14 },{ wch:8 },{ wch:10 },{ wch:8 }];
  return ws;
}

// ===== Sheet5: クラファン限定リスト =====
function buildSheet5(data) {
  const ws = {};
  const cfData = data.filter(r => getSiteType(r[COL.EC]) === 'クラファン' && hasMail(r[COL.MAIL]));

  const headers = ['No.','カテゴリ','製品名','メーカー','ECサイト','メール','★評価','コメント'];
  headers.forEach((h, c) => setCell(ws, 0, c, h, hStyle('D84315', WHITE, true, 10)));

  cfData.forEach((row, i) => {
    const r  = i + 1;
    const bg = i % 2 === 0 ? WHITE : 'FFF8F4';
    setCell(ws, r, 0, Number(row[COL.NO]) || 0,      dStyle(bg, true));
    setCell(ws, r, 1, String(row[COL.CAT]  || ''),   dStyle(bg));
    setCell(ws, r, 2, String(row[COL.NAME] || ''),   dStyle(bg));
    setCell(ws, r, 3, String(row[COL.MAKER]|| ''),   dStyle(bg));
    setCell(ws, r, 4, String(row[COL.EC]   || ''),   dStyle(bg, true));
    // メール列：Gmailリンク付き
    const cfMail = String(row[COL.MAIL] || '');
    const cfMailCell = XLSX.utils.encode_cell({ r, c: 5 });
    if (cfMail.includes('@') && !cfMail.startsWith('http')) {
      const gmailUrl = buildGmailUrl(cfMail, String(row[COL.MAKER]||''), String(row[COL.NAME]||''));
      ws[cfMailCell] = { v: cfMail, t: 's', l: { Target: gmailUrl }, s: { ...dStyle(bg), font: { sz:9, name:'Meiryo UI', color:{ rgb:'1155CC' }, underline: true } } };
    } else {
      setCell(ws, r, 5, cfMail, dStyle(bg));
    }
    setCell(ws, r, 6, String(row[COL.STAR] || ''),   dStyle(bg, true));
    setCell(ws, r, 7, String(row[COL.COMMENT]|| ''), dStyle(bg));
  });

  ws['!ref']  = XLSX.utils.encode_range({ r:0, c:0 }, { r: cfData.length, c: 7 });
  ws['!cols'] = [{ wch:6 },{ wch:20 },{ wch:38 },{ wch:22 },{ wch:14 },{ wch:28 },{ wch:10 },{ wch:50 }];
  return ws;
}

// ===== ワークブック組み立て =====
const wb = XLSX.utils.book_new();

console.log('データ読み込み中...');
console.log(`  総件数: ${data.length}件`);

console.log('Sheet1: 全件リスト を生成中...');
XLSX.utils.book_append_sheet(wb, buildSheet1(data), '全件リスト');

console.log('Sheet2: カテゴリ別集計 を生成中...');
XLSX.utils.book_append_sheet(wb, buildSheet2(data), 'カテゴリ別集計');

console.log('Sheet3: ECサイト別集計 を生成中...');
XLSX.utils.book_append_sheet(wb, buildSheet3(data), 'ECサイト別集計');

console.log('Sheet4: ★評価別集計 を生成中...');
XLSX.utils.book_append_sheet(wb, buildSheet4(data), '★評価別集計');

console.log('Sheet5: クラファン限定リスト を生成中...');
XLSX.utils.book_append_sheet(wb, buildSheet5(data), 'クラファン限定(メールあり)');

XLSX.writeFile(wb, OUTPUT_FILE, { bookSST: false, type: 'buffer', cellStyles: true });

console.log(`\n✓ 完了: ${OUTPUT_FILE}`);
console.log(`  総件数: ${data.length}件`);
const cfCount   = data.filter(r => getSiteType(r[COL.EC]) === 'クラファン').length;
const mailCount = data.filter(r => hasMail(r[COL.MAIL])).length;
const cfMail    = data.filter(r => getSiteType(r[COL.EC]) === 'クラファン' && hasMail(r[COL.MAIL])).length;
console.log(`  クラファン: ${cfCount}件`);
console.log(`  メールあり: ${mailCount}件`);
console.log(`  クラファン×メールあり: ${cfMail}件`);
