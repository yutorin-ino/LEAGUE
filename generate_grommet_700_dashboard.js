/**
 * generate_grommet_700_dashboard.js
 * The Grommet専用・No.1〜700独立ダッシュボード生成（スタンドアロンHTML・XLSX不要）
 * node generate_grommet_700_dashboard.js
 */

const fs = require('fs');
const GROMMET_700 = require('./grommet_700.js');

const OUTPUT = "C:\\Users\\hiro\\Desktop\\LEAGUE\\grommet_700_dashboard.html";
const SENDER = 'yutorin.ino@gmail.com';

function buildGmailUrl(email, maker, product) {
  const enc = s => encodeURIComponent(s);
  const subject = 'Potential Distribution Partnership for Japan';
  const body = `Dear ${maker} Team,

My name is Hiroyuki Inoguchi from
Sumai pluS in Japan

I hope this message finds you all well.

Our company focuses on promoting products that enrich people's daily lives. We are currently seeking unique international brands to introduce to the Japanese market and support their growth.

I recently had the opportunity to review your product (${product}) on The Grommet and was very impressed by its potential in the Japanese market.

I am confident that your product will strongly appeal to Japanese customers, and I would like to explore the possibility of building a partnership with your company to help ensure its success.

Furthermore, our company has a proven track record of promoting overseas products through our proprietary sales network and Japanese distribution channels.

Would it be possible to schedule a brief online meeting sometime next week to discuss this matter?

I look forward to hearing from you.


Sincerely,

Hiroyuki Inoguchi
Sumai pluS
 (LEAGUE Co., Ltd. Agent )

================================================

Sumai pluS
 (LEAGUE Co., Ltd. Agent )
Hiroyuki Inoguchi

E-mail:${SENDER}

Address: 49-5 Kitazakuno, Higashi-Itsushiro, Ichinomiya City, Aichi Prefecture, Japan

=================================================`;
  return `https://mail.google.com/mail/?authuser=${enc(SENDER)}&view=cm&fs=1&to=${enc(email)}&su=${enc(subject)}&body=${enc(body)}`;
}

console.log("grommet_700.js 読み込み中...");

const products = [];
for (const row of GROMMET_700) {
  const no       = row[0];
  const category = String(row[1] || "");
  const name     = String(row[2] || "");
  const maker    = String(row[3] || "");
  const prod_url = String(row[5] || "");
  const maker_hp = String(row[6] || "");
  const contact  = String(row[7] || "").trim();
  const stars    = String(row[8] || "");
  const comment  = String(row[9] || "");

  const hasEmail = contact.includes("@") && !contact.startsWith("http");
  const gmailUrl = hasEmail ? buildGmailUrl(contact, maker, name) : "";

  products.push({ no, category, name, maker, prod_url, maker_hp, contact, has_email: hasEmail, gmail_url: gmailUrl, stars, comment });
}

const categories = [...new Set(products.map(p => p.category).filter(Boolean))].sort();
const emailCount = products.filter(p => p.has_email).length;

console.log(`The Grommet 700 製品数: ${products.length} 件 / メールあり: ${emailCount} 件`);

const dataJson   = JSON.stringify(products);
const catOptions = categories.map(c => `<option value="${c}">${c}</option>`).join('\n    ');

const html = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>🟢 The Grommet No.1-700 専用ダッシュボード</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Segoe UI", "Yu Gothic", sans-serif; background: #f0f4f8; color: #222; }

  header {
    background: linear-gradient(135deg, #1a6b2e 0%, #2d9e4a 100%);
    color: white; padding: 16px 24px;
  }
  header .h-top { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }
  header h1 { font-size: 1.2rem; font-weight: 700; }
  .badge { display: inline-block; border-radius: 20px; padding: 3px 12px; font-weight: 700; font-size: 0.85rem; }
  .badge-total { background: #f0a500; color: #1a2d0a; }
  .badge-email { background: #6366f1; color: white; }
  header .h-stats { display: flex; gap: 10px; margin-top: 10px; flex-wrap: wrap; }
  .grommet-logo { font-size: 1.5rem; }

  .toolbar {
    background: white; padding: 12px 24px; display: flex; flex-wrap: wrap;
    gap: 10px; align-items: center; border-bottom: 1px solid #ddd;
    position: sticky; top: 0; z-index: 100; box-shadow: 0 2px 8px rgba(0,0,0,0.1);
  }
  .toolbar input[type=text] {
    border: 1.5px solid #bbb; border-radius: 8px; padding: 7px 12px;
    font-size: 0.9rem; width: 220px; outline: none;
  }
  .toolbar input[type=text]:focus { border-color: #1a6b2e; }
  .toolbar select {
    border: 1.5px solid #bbb; border-radius: 8px; padding: 7px 10px;
    font-size: 0.9rem; outline: none;
  }
  .btn { padding: 7px 16px; border-radius: 8px; border: none; cursor: pointer; font-size: 0.88rem; font-weight: 600; transition: background 0.15s; }
  .btn-primary { background: #1a6b2e; color: white; }
  .btn-primary:hover { background: #145922; }
  .btn-primary:disabled { background: #aaa; cursor: not-allowed; }
  .btn-success { background: #28a745; color: white; }
  .btn-success:hover { background: #218838; }
  .btn-outline { background: white; color: #1a6b2e; border: 1.5px solid #1a6b2e; }
  .btn-outline:hover { background: #e8f5e9; }
  .sel-count { font-size: 0.88rem; color: #555; white-space: nowrap; }
  .sel-count span { font-weight: 700; color: #1a6b2e; }
  .filter-info { font-size: 0.82rem; color: #777; }

  .table-wrap { padding: 14px 24px; overflow-x: auto; }
  table {
    width: 100%; border-collapse: collapse; background: white;
    border-radius: 10px; overflow: hidden;
    box-shadow: 0 1px 4px rgba(0,0,0,0.1); font-size: 0.83rem;
  }
  thead { background: #1a6b2e; color: white; }
  thead th { padding: 10px 9px; text-align: left; font-weight: 600; white-space: nowrap; }
  thead th:first-child { width: 40px; text-align: center; }
  tbody tr { border-bottom: 1px solid #eee; }
  tbody tr:hover { background: #f0f9f2; }
  tbody tr.selected { background: #d4edda; }
  tbody tr.no-email { opacity: 0.55; }
  tbody td { padding: 8px 9px; vertical-align: middle; }
  tbody td:first-child { text-align: center; }
  .no-col { color: #888; font-size: 0.78rem; }
  .cat-badge {
    display: inline-block; padding: 2px 7px; border-radius: 12px;
    font-size: 0.72rem; font-weight: 600; white-space: nowrap;
  }
  .grommet-badge {
    display: inline-block; padding: 2px 8px; border-radius: 10px;
    font-size: 0.72rem; font-weight: 700; background: #e8f5e9;
    color: #1a6b2e; border: 1px solid #2d9e4a; white-space: nowrap;
  }
  .name-col { max-width: 190px; font-weight: 500; }
  .maker-col { max-width: 120px; color: #444; }
  a.link { color: #1a6b2e; text-decoration: none; font-size: 0.8rem; }
  a.link:hover { text-decoration: underline; }
  .email-col { max-width: 170px; font-size: 0.8rem; color: #555; word-break: break-all; }
  .no-email-tag { color: #aaa; font-size: 0.78rem; }
  input[type=checkbox] { width: 16px; height: 16px; cursor: pointer; accent-color: #1a6b2e; }
  .stars-5 { color: #16a34a; font-weight: 700; }
  .stars-4 { color: #ca8a04; font-weight: 700; }
  .stars-3 { color: #ea580c; }
  .comment-col { max-width: 240px; font-size: 0.76rem; color: #555; line-height: 1.4; }

  .upvote-info { font-size: 0.72rem; color: #2d9e4a; font-weight: 600; }

  .modal { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.5); z-index: 200; align-items: center; justify-content: center; }
  .modal.show { display: flex; }
  .modal-box { background: white; border-radius: 12px; padding: 26px; max-width: 480px; width: 90%; box-shadow: 0 8px 32px rgba(0,0,0,0.2); }
  .modal-box h2 { font-size: 1.05rem; margin-bottom: 10px; color: #1a6b2e; }
  .modal-box p { font-size: 0.9rem; color: #444; margin-bottom: 8px; line-height: 1.6; }
  .modal-box .warn { background: #d4edda; border: 1px solid #28a745; border-radius: 6px; padding: 10px 14px; font-size: 0.83rem; color: #155724; margin: 12px 0; }
  .modal-box .btns { display: flex; gap: 10px; justify-content: flex-end; margin-top: 16px; }
  .progress { display: none; margin-top: 12px; }
  .progress-bar { background: #eee; border-radius: 8px; height: 10px; overflow: hidden; }
  .progress-fill { background: #28a745; height: 100%; transition: width 0.2s; }
  .progress-text { font-size: 0.8rem; color: #555; margin-top: 6px; text-align: center; }

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
  <div class="h-top">
    <span class="grommet-logo">🟢</span>
    <h1>The Grommet No.1-700 専用ダッシュボード — 独立採番・厳選アメリカ製品キュレーション</h1>
  </div>
  <div class="h-stats">
    <span class="badge badge-total">総数 ${products.length} 件</span>
    <span class="badge badge-email">✉ メールあり ${emailCount} 件</span>
    <span class="badge" style="background:#fff;color:#1a6b2e;">🏆 全件 thegrommet.com 掲載製品のみ</span>
  </div>
</header>

<div class="toolbar">
  <input type="text" id="searchBox" placeholder="🔍 製品名・メーカー名で検索..." oninput="applyFilter()">
  <select id="catFilter" onchange="applyFilter()">
    <option value="">すべてのカテゴリ</option>
    ${catOptions}
  </select>
  <select id="emailFilter" onchange="applyFilter()">
    <option value="">メール：すべて</option>
    <option value="yes">メールあり</option>
    <option value="no">メールなし</option>
  </select>
  <select id="starsFilter" onchange="applyFilter()">
    <option value="">評価：すべて</option>
    <option value="5">★★★★★ のみ</option>
    <option value="4">★★★★ 以上</option>
  </select>
  <button class="btn btn-outline" onclick="selectAllVisible()">表示中を全選択</button>
  <button class="btn btn-outline" onclick="clearSelection()">選択解除</button>
  <span class="sel-count">選択中: <span id="selCount">0</span> 件</span>
  <span class="filter-info">💡 Shift+クリックで範囲選択</span>
  <span class="filter-info" id="filterInfo"></span>
  <button class="btn btn-success" id="sendBtn" onclick="confirmSend()" disabled>📧 一括メール送信</button>
</div>

<div class="table-wrap">
  <table id="mainTable">
    <thead>
      <tr>
        <th><input type="checkbox" id="headerCheck" onchange="toggleHeaderCheck(this)"></th>
        <th>No.</th>
        <th>カテゴリ</th>
        <th>製品名</th>
        <th>メーカー</th>
        <th>Grommet</th>
        <th>メーカーHP</th>
        <th>メール</th>
        <th>評価</th>
        <th>コメント</th>
      </tr>
    </thead>
    <tbody id="tableBody"></tbody>
  </table>
</div>

<div class="modal" id="sendModal">
  <div class="modal-box">
    <h2>📧 一括メール送信の確認</h2>
    <p>以下の件数のメーカーにGmailで連絡します：</p>
    <p><strong id="modalCount" style="font-size:1.3rem;color:#1a6b2e;"></strong></p>
    <div class="warn">
      ✅ The Grommet掲載製品のメーカーへの日本展開提案メールです。<br>
      送信前に内容をご確認ください。
    </div>
    <div class="progress" id="progressArea">
      <div class="progress-bar"><div class="progress-fill" id="progressFill" style="width:0%"></div></div>
      <div class="progress-text" id="progressText">準備中...</div>
    </div>
    <div class="btns" id="modalBtns">
      <button class="btn btn-outline" onclick="closeModal()">キャンセル</button>
      <button class="btn btn-primary" id="goSendBtn" onclick="startSend()">送信開始</button>
    </div>
  </div>
</div>

<script>
const ALL_PRODUCTS = ${dataJson};
let visibleIds = [];
let selectedIds = new Set();
let currentList = [];
let lastClickedNo = null;

const CAT_COLORS = {
  'キッチン・調理器具':         'cat-キッチン・調理器具',
  'スマートホーム・インテリア・照明': 'cat-スマートホーム・インテリア・照明',
  'ウェアラブル・ヘルス・フィットネス': 'cat-ウェアラブル・ヘルス・フィットネス',
  'アウトドア・スポーツ・旅行':   'cat-アウトドア・スポーツ・旅行',
  'ペット用品':               'cat-ペット用品',
  'テクノロジー・ガジェット':     'cat-テクノロジー・ガジェット',
  '美容・スキンケア':           'cat-美容・スキンケア',
  '子供・教育':               'cat-子供・教育',
  'ファッション・アクセサリー':   'cat-ファッション・アクセサリー',
  'クリーニング・収納・整理':     'cat-クリーニング・収納・整理',
};

function starsClass(s) {
  if (s.includes('★★★★★')) return 'stars-5';
  if (s.includes('★★★★'))  return 'stars-4';
  return 'stars-3';
}

// コメントからupvote数を抽出して表示
function extractUpvote(comment) {
  const m = comment.match(/\\(([0-9,]+)\\s*upvotes?\\)/i);
  return m ? '<br><span class="upvote-info">👍 ' + m[1] + ' upvotes</span>' : '';
}

function renderTable(list) {
  const tbody = document.getElementById('tableBody');
  currentList = list;
  visibleIds = list.map(p => p.no);
  tbody.innerHTML = list.map(p => {
    const catCls  = CAT_COLORS[p.category] || '';
    const sCls    = starsClass(p.stars);
    const rowCls  = [selectedIds.has(p.no) ? 'selected' : '', !p.has_email ? 'no-email' : ''].filter(Boolean).join(' ');
    const upvote  = extractUpvote(p.comment);
    const cleanComment = p.comment.replace(/\\s*\\([0-9,]+ upvotes?\\)/i, '');
    const emailCell = p.has_email
      ? \`<a class="link" href="\${p.gmail_url}" target="_blank">\${p.contact}</a>\`
      : '<span class="no-email-tag">メールなし</span>';
    const nameCell = p.prod_url
      ? \`<a class="link" href="\${p.prod_url}" target="_blank">\${p.name}</a>\${upvote}\`
      : p.name + upvote;
    const hpCell = p.maker_hp
      ? \`<a class="link" href="\${p.maker_hp}" target="_blank">🔗 公式</a>\`
      : '';
    return \`<tr class="\${rowCls}" data-id="\${p.no}">
      <td><input type="checkbox" \${p.has_email ? '' : 'disabled'} \${selectedIds.has(p.no) ? 'checked' : ''} onclick="toggleRow(this, '\${p.no}', event)"></td>
      <td class="no-col">\${p.no}</td>
      <td><span class="cat-badge \${catCls}">\${p.category}</span></td>
      <td class="name-col">\${nameCell}</td>
      <td class="maker-col">\${p.maker}</td>
      <td><span class="grommet-badge">🟢 Grommet</span></td>
      <td>\${hpCell}</td>
      <td class="email-col">\${emailCell}</td>
      <td class="\${sCls}">\${p.stars}</td>
      <td class="comment-col">\${cleanComment}</td>
    </tr>\`;
  }).join('');
  document.getElementById('filterInfo').textContent = \`表示: \${list.length} 件\`;
  updateSendBtn();
}

function applyFilter() {
  const q    = document.getElementById('searchBox').value.toLowerCase();
  const cat  = document.getElementById('catFilter').value;
  const mail = document.getElementById('emailFilter').value;
  const star = document.getElementById('starsFilter').value;
  const list = ALL_PRODUCTS.filter(p => {
    if (cat  && p.category !== cat) return false;
    if (mail === 'yes' && !p.has_email) return false;
    if (mail === 'no'  &&  p.has_email) return false;
    if (star === '5' && !p.stars.includes('★★★★★')) return false;
    if (star === '4' && !p.stars.includes('★★★★'))  return false;
    if (q && !(p.name.toLowerCase().includes(q) || p.maker.toLowerCase().includes(q) || p.comment.toLowerCase().includes(q))) return false;
    return true;
  });
  renderTable(list);
}

function toggleRow(cb, id, evt) {
  const numId = Number(id);
  const shouldCheck = cb.checked;

  if (evt && evt.shiftKey && lastClickedNo !== null) {
    const lastIdx = visibleIds.indexOf(lastClickedNo);
    const curIdx  = visibleIds.indexOf(numId);
    if (lastIdx !== -1 && curIdx !== -1) {
      const start = Math.min(lastIdx, curIdx);
      const end   = Math.max(lastIdx, curIdx);
      for (let i = start; i <= end; i++) {
        const no = visibleIds[i];
        const p = ALL_PRODUCTS.find(x => x.no === no);
        if (!p || !p.has_email) continue;
        if (shouldCheck) selectedIds.add(no); else selectedIds.delete(no);
      }
      lastClickedNo = numId;
      renderTable(currentList);
      return;
    }
  }

  if (shouldCheck) selectedIds.add(numId); else selectedIds.delete(numId);
  lastClickedNo = numId;
  const row = cb.closest('tr');
  row.classList.toggle('selected', shouldCheck);
  updateSendBtn();
  document.getElementById('headerCheck').checked = false;
}

function toggleHeaderCheck(hcb) {
  visibleIds.forEach(id => {
    const p = ALL_PRODUCTS.find(x => x.no === id);
    if (!p || !p.has_email) return;
    if (hcb.checked) selectedIds.add(id); else selectedIds.delete(id);
  });
  applyFilter();
}

function selectAllVisible() {
  visibleIds.forEach(id => {
    const p = ALL_PRODUCTS.find(x => x.no === id);
    if (p && p.has_email) selectedIds.add(id);
  });
  applyFilter();
}

function clearSelection() {
  selectedIds.clear();
  applyFilter();
}

function updateSendBtn() {
  const cnt = selectedIds.size;
  document.getElementById('selCount').textContent = cnt;
  document.getElementById('sendBtn').disabled = cnt === 0;
}

function confirmSend() {
  document.getElementById('modalCount').textContent = selectedIds.size + ' 社';
  document.getElementById('progressArea').style.display = 'none';
  document.getElementById('progressFill').style.width = '0%';
  document.getElementById('modalBtns').style.display = 'flex';
  document.getElementById('sendModal').classList.add('show');
}

function closeModal() {
  document.getElementById('sendModal').classList.remove('show');
}

async function startSend() {
  const ids  = [...selectedIds];
  const total = ids.length;
  document.getElementById('modalBtns').style.display = 'none';
  document.getElementById('progressArea').style.display = 'block';
  for (let i = 0; i < total; i++) {
    const p = ALL_PRODUCTS.find(x => x.no === ids[i]);
    if (p && p.gmail_url) window.open(p.gmail_url, '_blank');
    const pct = Math.round((i + 1) / total * 100);
    document.getElementById('progressFill').style.width = pct + '%';
    document.getElementById('progressText').textContent = \`\${i + 1} / \${total} 件送信中...\`;
    await new Promise(r => setTimeout(r, 600));
  }
  document.getElementById('progressText').textContent = \`✅ \${total} 件の送信が完了しました！\`;
  setTimeout(() => closeModal(), 2000);
}

// 初期表示
applyFilter();
</script>
</body>
</html>`;

fs.writeFileSync(OUTPUT, html, 'utf8');
console.log(`完了: ${OUTPUT}`);
console.log(`The Grommet 700 製品数: ${products.length} 件 / メールあり: ${emailCount} 件`);
