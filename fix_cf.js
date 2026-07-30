const fs = require('fs');
let c = fs.readFileSync('generate_cf_dashboard.js', 'utf8');
const lines = c.split('\n');
for (let i = 0; i < lines.length; i++) {
  // Fix the corrupted CF_KEYWORDS line
  if (lines[i].includes('CF_KEYWORDS') && lines[i].includes('Kickstarter') && !lines[i].startsWith('const CF_KEYWORDS')) {
    lines[i] = "const CF_KEYWORDS    = ['Kickstarter', 'Indiegogo', 'BackerKit', 'Crowdfunder', 'Crowd Supply'];";
    console.log('Fixed line', i+1, ':', lines[i]);
  }
}
fs.writeFileSync('generate_cf_dashboard.js', lines.join('\n'), 'utf8');
console.log('Done');
