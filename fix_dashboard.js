const fs = require('fs');

const newInput = 'C:\\\\Users\\\\hiro\\\\Desktop\\\\LEAGUE\\\\海外便利グッズリスト_日本未上陸3305件_評価付.xlsx';

function fixFile(filename) {
  let c = fs.readFileSync(filename, 'utf8');
  const lines = c.split('\n');
  for (let i = 0; i < lines.length; i++) {
    // Fix INPUT path (any corrupted Japanese filename)
    if (lines[i].startsWith('const INPUT') && lines[i].includes('xlsx')) {
      lines[i] = `const INPUT = "${newInput}";`;
      console.log(`[${filename}] Fixed INPUT line ${i+1}`);
    }
    // Fix broken 問い合わせ string
    if (lines[i].includes('contact.includes') && lines[i].includes('問い合わ') && !lines[i].trim().endsWith('"));')) {
      lines[i] = '  const hasEmail = contact.includes("@") && !contact.startsWith("http") && !contact.includes("問い合わせ");';
      console.log(`[${filename}] Fixed hasEmail line ${i+1}`);
    }
  }
  fs.writeFileSync(filename, lines.join('\n'), 'utf8');
  console.log(`[${filename}] Done`);
}

fixFile('generate_dashboard.js');
fixFile('generate_cf_dashboard.js');
