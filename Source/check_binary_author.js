const fs = require('fs');
const path = require('path');

const dotmStartup = path.join(process.env.APPDATA, 'Microsoft\\Word\\STARTUP\\1Click.dotm');
const dotmOutput = 'G:\\1Click Build\\Output\\1Click.dotm';

[dotmStartup, dotmOutput].forEach(f => {
  const buf = fs.readFileSync(f);
  const str = buf.toString('latin1');
  const hasOld = str.includes('Thanh Phu') || str.includes('VPDU xa');
  const hasNew = str.includes('Tran Quoc Duoc');
  console.log(`File ${f}:`);
  console.log(`  Has old 'Thanh Phu': ${hasOld}`);
  console.log(`  Has new 'Tran Quoc Duoc': ${hasNew}`);
});
