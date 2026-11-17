const fs = require('fs');

let vba = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8');

const target = '"Created by VPDU xa Thanh Phu"';
const replacement = '"1Click Office - Tran Quoc Duoc (TTDVTH Phuong Tan Khai)"';

if (vba.includes(target)) {
  vba = vba.split(target).join(replacement);
  fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', vba, 'utf8');
  console.log('Successfully replaced all occurrences of ' + target);
} else {
  console.log('Target not found!');
}

// Verify no more occurrences
if (!vba.includes('Thanh Phu')) {
  console.log('VERIFIED: Zero occurrences of Thanh Phu remain in Module1_clean.vba!');
} else {
  console.log('WARNING: Some occurrences of Thanh Phu still exist!');
}
