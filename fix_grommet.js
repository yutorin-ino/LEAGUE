const fs = require('fs');
let c = fs.readFileSync('generate_grommet_dashboard.js', 'utf8');

// The problem: inside the outer html = `...` template literal,
// the embedded JS has unescaped backtick template literals.
// Fix by replacing them with string concatenation.

// Fix extractUpvote function
c = c.replace(
  "return m ? `<br><span class=\"upvote-info\">👍 ${m[1]} upvotes</span>` : '';",
  "return m ? '<br><span class=\"upvote-info\">👍 ' + m[1] + ' upvotes</span>' : '';"
);

fs.writeFileSync('generate_grommet_dashboard.js', c, 'utf8');
console.log('Fixed. Lines around fix:');
const lines = c.split('\n');
for (let i = 305; i < 312; i++) console.log(i+1 + ':', lines[i]);
