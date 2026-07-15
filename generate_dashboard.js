const XLSX = require('xlsx');
const fs = require('fs');

const INPUT = "C:\\Users\\hiro\\Desktop\\LEAGUE\\海外便利グッズリスト_日本未上陸3105件_評価付.xlsx";
const OUTPUT = "C:\\Users\\hiro\\Desktop\\LEAGUE\\mail_dashboard.html";

console.log("Excelファイル読み込み中...");
const wb = XLSX.readFile(INPUT);
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

function encodeURIFull(str) {
  return encodeURIComponent(String(str));
}

function buildGmailUrl(email, makerName, productName) {
  const subject = "Potential Distribution Partnership for Japan";
  const body = `Dear ${makerName} Team,

My name is Hiroyuki Inoguchi from
Sumai pluS Co., Ltd. in Japan

I hope this message finds you all well.

Our company focuses on promoting products that enrich people's daily lives. We are currently seeking unique international brands to introduce to the Japanese market and support their growth.

I recently had the opportunity to review your product (${productName}) and was very impressed by its potential in the Japanese market.

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

E-mail:yutorin.ino@gmail.com

Address: 49-5 Kitazakuno, Higashi-Itsushiro, Ichinomiya City, Aichi Prefecture, Japan

=================================================`;

  const sender = encodeURIFull("yutorin.ino@gmail.com");
  return `https://mail.google.com/mail/?authuser=${sender}&view=cm&fs=1&to=${encodeURIFull(email)}&su=${encodeURIFull(subject)}&body=${encodeURIFull(body)}`;
}

const products = [];
for (let i = 1; i < rows.length; i++) {
  const row = rows[i];
  if (!row[0]) continue;
  const no       = row[0];
  const category = String(row[1] || "");
  const name     = String(row[2] || "");
  const maker    = String(row[3] || "");
  const ec_site  = String(row[4] || "");
  const prod_url = String(row[5] || "");
  const maker_hp = String(row[6] || "");
  const contact  = String(row[7] || "").trim();
  const stars    = String(row[8] || "");
  const comment  = String(row[9] || "");

  const hasEmail = contact.includes("@") && !contact.startsWith("http") && !contact.includes("問い合わせ");
  const gmailUrl = hasEmail ? buildGmailUrl(contact, maker, name) : "";

  products.push({ no, category, name, maker, ec_site, prod_url, maker_hp, contact, has_email: hasEmail, gmail_url: gmailUrl, stars, comment });
}

const categories = [...new Set(products.map(p => p.category).filter(Boolean))].sort();
const emailCount = products.filter(p => p.has_email).length;

console.log(`商品数: ${products.length} 件 / メールあり: ${emailCount} 件`);

const dataJson = JSON.stringify(products);
const catOptions = categories.map(c => `<option value="${c}">${c}</option>`).join('\n    ');

const html = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>海外便利グッズ 一括メール送信ダッシュボード</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Segoe UI", "Yu Gothic", sans-serif; background: #f0f4f8; color: #222; }
  header { background: #1a3a6b; color: white; padding: 18px 24px; display: flex; align-items: center; gap: 16px; }
  header h1 { font-size: 1.25rem; font-weight: 700; }
  header .count { background: #f0a500; color: #1a3a6b; border-radius: 20px; padding: 3px 12px; font-weight: 700; font-size: 0.9rem; }
  .toolbar { background: white; padding: 14px 24px; display: flex; flex-wrap: wrap; gap: 12px; align-items: center; border-bottom: 1px solid #ddd; position: sticky; top: 0; z-index: 100; box-shadow: 0 2px 6px rgba(0,0,0,0.08); }
  .toolbar input[type=text] { border: 1.5px solid #bbb; border-radius: 8px; padding: 7px 12px; font-size: 0.95rem; width: 240px; outline: none; }
  .toolbar input[type=text]:focus { border-color: #1a3a6b; }
  .toolbar select { border: 1.5px solid #bbb; border-radius: 8px; padding: 7px 10px; font-size: 0.95rem; outline: none; }
  .btn { padding: 8px 18px; border-radius: 8px; border: none; cursor: pointer; font-size: 0.92rem; font-weight: 600; transition: background 0.15s; }
  .btn-primary { background: #1a3a6b; color: white; }
  .btn-primary:hover { background: #2a5299; }
  .btn-primary:disabled { background: #aaa; cursor: not-allowed; }
  .btn-success { background: #28a745; color: white; }
  .btn-success:hover { background: #218838; }
  .btn-outline { background: white; color: #1a3a6b; border: 1.5px solid #1a3a6b; }
  .btn-outline:hover { background: #eef2ff; }
  .sel-count { font-size: 0.9rem; color: #555; white-space: nowrap; }
  .sel-count span { font-weight: 700; color: #1a3a6b; }
  .filter-info { font-size: 0.85rem; color: #777; }

  .table-wrap { padding: 16px 24px; overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,0.1); font-size: 0.85rem; }
  thead { background: #1a3a6b; color: white; }
  thead th { padding: 11px 10px; text-align: left; font-weight: 600; white-space: nowrap; }
  thead th:first-child { width: 40px; text-align: center; }
  tbody tr { border-bottom: 1px solid #eee; }
  tbody tr:hover { background: #f5f8ff; }
  tbody tr.selected { background: #dceeff; }
  tbody tr.no-email { opacity: 0.55; }
  tbody td { padding: 9px 10px; vertical-align: middle; }
  tbody td:first-child { text-align: center; }
  .no-col { color: #888; font-size: 0.8rem; width: 40px; text-align: center; }
  .cat-badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 0.75rem; font-weight: 600; white-space: nowrap; }
  .name-col { max-width: 200px; font-weight: 500; }
  .maker-col { max-width: 130px; color: #444; }
  a.link { color: #1a6bc5; text-decoration: none; font-size: 0.82rem; }
  a.link:hover { text-decoration: underline; }
  .email-col { max-width: 180px; font-size: 0.82rem; color: #555; word-break: break-all; }
  .no-email-tag { color: #aaa; font-size: 0.8rem; }
  input[type=checkbox] { width: 16px; height: 16px; cursor: pointer; accent-color: #1a3a6b; }
  .stars-5 { color: #16a34a; font-weight: 700; }
  .stars-4 { color: #ca8a04; font-weight: 700; }
  .stars-3 { color: #ea580c; }
  .stars-2 { color: #dc2626; }
  .stars-1 { color: #991b1b; }
  .comment-col { max-width: 260px; font-size: 0.78rem; color: #555; line-height: 1.4; }

  .modal { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 200; align-items: center; justify-content: center; }
  .modal.show { display: flex; }
  .modal-box { background: white; border-radius: 12px; padding: 28px; max-width: 480px; width: 90%; box-shadow: 0 8px 32px rgba(0,0,0,0.2); }
  .modal-box h2 { font-size: 1.1rem; margin-bottom: 12px; color: #1a3a6b; }
  .modal-box p { font-size: 0.92rem; color: #444; margin-bottom: 8px; line-height: 1.6; }
  .modal-box .warn { background: #fff3cd; border: 1px solid #f0ad4e; border-radius: 6px; padding: 10px 14px; font-size: 0.85rem; color: #856404; margin: 14px 0; }
  .modal-box .btns { display: flex; gap: 10px; justify-content: flex-end; margin-top: 18px; }
  .progress { display: none; margin-top: 14px; }
  .progress-bar { background: #eee; border-radius: 8px; height: 10px; overflow: hidden; }
  .progress-fill { background: #28a745; height: 100%; transition: width 0.2s; }
  .progress-text { font-size: 0.82rem; color: #555; margin-top: 6px; text-align: center; }

  .cat-キッチン・調理器具 { background: #dbeafe; color: #1e40af; }
  .cat-スマートホーム・インテリア・照明 { background: #d1fae5; color: #065f46; }
  .cat-ウェアラブル・ヘルス・フィットネス { background: #ede9fe; color: #4c1d95; }
  .cat-アウトドア・スポーツ・旅行 { background: #fee2e2; color: #991b1b; }
  .cat-ペット用品 { background: #fef3c7; color: #78350f; }
  .cat-テクノロジー・ガジェット { background: #e0f2fe; color: #0c4a6e; }
  .cat-美容・スキンケア { background: #fce7f3; color: #831843; }
  .cat-子供・教育 { background: #d1fae5; color: #14532d; }
  .cat-ファッション・アクセサリー { background: #fdf2f8; color: #701a75; }
  .cat-クリーニング・収納・整理 { background: #f3f4f6; color: #374151; }
</style>
</head>
<body>

<header>
  <h1>🌐 海外便利グッズ 一括メール送信ダッシュボード</h1>
  <span class="count" id="totalCount"></span>
</header>

<div class="toolbar">
  <input type="text" id="searchBox" placeholder="🔍 商品名・メーカー名で検索..." oninput="applyFilter()">
  <select id="catFilter" onchange="applyFilter()">
    <option value="">すべてのカテゴリ</option>
    ${catOptions}
  </select>
  <select id="emailFilter" onchange="applyFilter()">
    <option value="">メール有無：すべて</option>
    <option value="yes">メールあり</option>
    <option value="no">メールなし</option>
  </select>
  <select id="starsFilter" onchange="applyFilter()">
    <option value="">需要評価：すべて</option>
    <option value="5">★★★★★ のみ</option>
    <option value="4">★★★★☆ 以上</option>
    <option value="3">★★★☆☆ 以上</option>
  </select>
  <button class="btn btn-outline" onclick="selectAllVisible()">表示中を全選択</button>
  <button class="btn btn-outline" onclick="clearSelection()">選択解除</button>
  <span class="sel-count">選択中: <span id="selCount">0</span> 件</span>
  <span class="filter-info" id="filterInfo"></span>
  <button class="btn btn-success" id="sendBtn" onclick="confirmSend()" disabled>📧 選択した相手に一括送信</button>
</div>

<div class="table-wrap">
  <table id="mainTable">
    <thead>
      <tr>
        <th><input type="checkbox" id="headerCheck" onchange="toggleHeaderCheck(this)"></th>
        <th>No.</th>
        <th>カテゴリ</th>
        <th>商品名</th>
        <th>メーカー名</th>
        <th>販売ECサイト</th>
        <th>商品URL</th>
        <th>メーカーHP</th>
        <th>連絡先メール</th>
        <th>日本需要評価</th>
        <th>需要マッチコメント</th>
      </tr>
    </thead>
    <tbody id="tableBody"></tbody>
  </table>
</div>

<div class="modal" id="confirmModal">
  <div class="modal-box">
    <h2>📧 一括メール送信の確認</h2>
    <p id="confirmText"></p>
    <div class="warn">
      ⚠️ ブラウザのポップアップブロックを許可してください。<br>
      複数のGmail作成画面が新しいタブで開きます。<br>
      一度に大量送信する場合は、数十件ずつ分けることを推奨します。
    </div>
    <div class="progress" id="progressArea">
      <div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>
      <div class="progress-text" id="progressText"></div>
    </div>
    <div class="btns">
      <button class="btn btn-outline" onclick="closeModal()">キャンセル</button>
      <button class="btn btn-success" id="execBtn" onclick="executeSend()">送信開始</button>
    </div>
  </div>
</div>

<script>
const products = ${dataJson};

let filtered = [...products];
let selected = new Set();

function starsClass(stars) {
  if (!stars) return '';
  const n = (stars.match(/★/g)||[]).length - (stars.match(/☆/g)||[]).length + (stars.match(/☆/g)||[]).length;
  const filled = (stars.match(/★/g)||[]).length;
  return 'stars-' + filled;
}

function init() {
  document.getElementById('totalCount').textContent = products.length + ' 件';
  renderTable(filtered);
}

function applyFilter() {
  const q = document.getElementById('searchBox').value.toLowerCase();
  const cat = document.getElementById('catFilter').value;
  const em = document.getElementById('emailFilter').value;
  const st = document.getElementById('starsFilter').value;
  filtered = products.filter(p => {
    const matchQ = !q || p.name.toLowerCase().includes(q) || p.maker.toLowerCase().includes(q);
    const matchCat = !cat || p.category === cat;
    const matchEm = !em || (em === 'yes' ? p.has_email : !p.has_email);
    let matchSt = true;
    if (st && p.stars) {
      const filled = (p.stars.match(/★/g)||[]).length;
      matchSt = filled >= parseInt(st);
    }
    return matchQ && matchCat && matchEm && matchSt;
  });
  document.getElementById('filterInfo').textContent = '表示中: ' + filtered.length + ' 件';
  renderTable(filtered);
  updateSendBtn();
}

function renderTable(rows) {
  const tbody = document.getElementById('tableBody');
  tbody.innerHTML = '';
  rows.forEach(p => {
    const tr = document.createElement('tr');
    if (selected.has(p.no)) tr.classList.add('selected');
    if (!p.has_email) tr.classList.add('no-email');
    const sc = starsClass(p.stars);
    tr.innerHTML =
      '<td><input type="checkbox"' + (!p.has_email ? ' disabled title="メールアドレスなし"' : '') + (selected.has(p.no) ? ' checked' : '') + ' onchange="toggleRow(' + JSON.stringify(p.no) + ', this)"></td>' +
      '<td class="no-col">' + p.no + '</td>' +
      '<td><span class="cat-badge cat-' + p.category + '">' + p.category + '</span></td>' +
      '<td class="name-col" title="' + esc(p.name) + '">' + truncate(p.name, 30) + '</td>' +
      '<td class="maker-col" title="' + esc(p.maker) + '">' + truncate(p.maker, 20) + '</td>' +
      '<td style="font-size:0.8rem;color:#555">' + p.ec_site + '</td>' +
      '<td>' + (p.prod_url ? '<a class="link" href="' + p.prod_url + '" target="_blank">🔗 商品ページ</a>' : '') + '</td>' +
      '<td>' + (p.maker_hp ? '<a class="link" href="' + p.maker_hp + '" target="_blank">🏠 HP</a>' : '') + '</td>' +
      '<td class="email-col">' + (p.has_email ? '<a class="link" href="' + p.gmail_url + '" target="_blank">✉️ ' + p.contact + '</a>' : '<span class="no-email-tag">メールなし</span>') + '</td>' +
      '<td class="' + sc + '" style="white-space:nowrap;font-size:1rem">' + (p.stars || '') + '</td>' +
      '<td class="comment-col">' + esc(p.comment) + '</td>';
    tbody.appendChild(tr);
  });
  document.getElementById('headerCheck').checked = false;
}

function esc(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function truncate(str, n) {
  return str.length > n ? str.slice(0, n) + '…' : str;
}

function toggleRow(no, cb) {
  if (cb.checked) selected.add(no);
  else selected.delete(no);
  cb.closest('tr').classList.toggle('selected', cb.checked);
  updateSendBtn();
  document.getElementById('selCount').textContent = selected.size;
}

function selectAllVisible() {
  filtered.forEach(p => { if (p.has_email) selected.add(p.no); });
  renderTable(filtered);
  updateSendBtn();
  document.getElementById('selCount').textContent = selected.size;
}

function clearSelection() {
  selected.clear();
  renderTable(filtered);
  updateSendBtn();
  document.getElementById('selCount').textContent = 0;
}

function toggleHeaderCheck(cb) {
  if (cb.checked) selectAllVisible();
  else clearSelection();
}

function updateSendBtn() {
  document.getElementById('sendBtn').disabled = selected.size === 0;
  document.getElementById('selCount').textContent = selected.size;
}

function confirmSend() {
  const sel = products.filter(p => selected.has(p.no) && p.has_email);
  document.getElementById('confirmText').textContent =
    '選択した ' + sel.length + ' 件の相手にGmailの作成画面を開きます。よろしいですか？';
  document.getElementById('progressArea').style.display = 'none';
  document.getElementById('execBtn').disabled = false;
  document.getElementById('confirmModal').classList.add('show');
}

function closeModal() {
  document.getElementById('confirmModal').classList.remove('show');
}

async function executeSend() {
  const sel = products.filter(p => selected.has(p.no) && p.has_email);
  document.getElementById('execBtn').disabled = true;
  const prog = document.getElementById('progressArea');
  prog.style.display = 'block';
  const fill = document.getElementById('progressFill');
  const text = document.getElementById('progressText');
  for (let i = 0; i < sel.length; i++) {
    window.open(sel[i].gmail_url, '_blank');
    const pct = Math.round(((i + 1) / sel.length) * 100);
    fill.style.width = pct + '%';
    text.textContent = (i + 1) + ' / ' + sel.length + ' 件 開きました...';
    await new Promise(r => setTimeout(r, 400));
  }
  text.textContent = '✅ ' + sel.length + ' 件のGmail作成画面を開きました！';
  setTimeout(() => closeModal(), 2000);
}

init();
</script>
</body>
</html>`;

fs.writeFileSync(OUTPUT, html, 'utf8');
console.log(`\n完了: ${OUTPUT}`);
console.log(`商品数: ${products.length} 件 / メールあり: ${emailCount} 件`);
