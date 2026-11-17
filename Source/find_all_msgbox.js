const fs = require('fs');

const vba = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8');

const regex = /MsgBox\s+[^\n]+/g;
let m;
let matches = [];
while ((m = regex.exec(vba)) !== null) {
  matches.push(m[0]);
}

console.log(`Total MsgBox calls found: ${matches.length}`);
matches.forEach((mb, i) => {
  if (i < 30) console.log(`${i + 1}: ${mb.substring(0, 100)}`);
});
