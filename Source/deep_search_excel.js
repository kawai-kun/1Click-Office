const fs = require('fs');

const vba = fs.readFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', 'utf8');

const lines = vba.split('\n');
const suspicious = ['thanh phu', 'thạnh phú', 'cuongbc', 'tác giả', 'tac gia', 'created by', 'phat trien', 'lien he', 'zalo', 'facebook', 'github.com/cuong'];

lines.forEach((line, lineNo) => {
  let lower = line.toLowerCase();
  for (let s of suspicious) {
    if (lower.includes(s)) {
      if (!lower.includes('tran quoc duoc') && !lower.includes('trần quốc được') && !lower.includes('tan khai') && !lower.includes('tân khai')) {
        console.log(`Excel Line ${lineNo + 1}: ${line.trim()}`);
      }
    }
  }
});
console.log('Excel scan complete.');
