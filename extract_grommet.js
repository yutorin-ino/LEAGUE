/**
 * extract_grommet.js
 * add_200_items_14.js / add_200_items_15.js から The Grommet 掲載製品のみ抽出し、
 * No.1〜 で連番を振り直した中間JSONを出力する。
 */
const fs = require('fs');
const path = require('path');

function extractProducts(filename) {
  const text = fs.readFileSync(path.join(__dirname, filename), 'utf8');
  const startIdx = text.indexOf('const NEW_PRODUCTS = [');
  if (startIdx === -1) throw new Error('NEW_PRODUCTS not found in ' + filename);
  const arrStart = text.indexOf('[', startIdx);
  // 対応する閉じ角括弧を探す（ネストしたブラケットに対応）
  let depth = 0, i = arrStart, end = -1;
  for (; i < text.length; i++) {
    if (text[i] === '[') depth++;
    else if (text[i] === ']') {
      depth--;
      if (depth === 0) { end = i; break; }
    }
  }
  const arrText = text.slice(arrStart, end + 1);
  const arr = new Function('return ' + arrText)();
  return arr;
}

const p14 = extractProducts('add_200_items_14.js');
const p15 = extractProducts('add_200_items_15.js');

const g14 = p14.filter(p => p[4] === 'The Grommet');
const g15 = p15.filter(p => p[4] === 'The Grommet');

console.log(`add_200_items_14.js: 総${p14.length}件 / Grommet ${g14.length}件`);
console.log(`add_200_items_15.js: 総${p15.length}件 / Grommet ${g15.length}件`);
console.log(`合計 Grommet件数: ${g14.length + g15.length}件`);

// No.1〜 に振り直し
const renumbered = [];
let no = 1;
for (const p of [...g14, ...g15]) {
  const row = [...p];
  row[0] = no;
  renumbered.push(row);
  no++;
}

fs.writeFileSync(
  path.join(__dirname, 'grommet_existing_293.json'),
  JSON.stringify(renumbered, null, 0),
  'utf8'
);
console.log(`出力: grommet_existing_293.json (${renumbered.length}件, 次の番号 = ${no})`);
