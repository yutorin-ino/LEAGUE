const fs = require('fs');
const OLD = '海外便利グッズリスト_日本未上陸3705件_評価付.xlsx';
const NEW = '海外便利グッズリスト_日本未上陸3905件_評価付.xlsx';

const files = [
  'generate_dashboard.js',
  'generate_cf_dashboard.js',
  'generate_grommet_dashboard.js',
];

for (const f of files) {
  let c = fs.readFileSync(f, 'utf8');
  const updated = c.split(OLD).join(NEW);
  fs.writeFileSync(f, updated, 'utf8');
  console.log(`Updated: ${f} (${c.split(OLD).length - 1} replacement(s))`);
}
