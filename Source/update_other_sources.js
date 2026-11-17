const fs = require('fs');

for (const fn of ['Module1_clean.bas', 'Module1_noattr.vba']) {
  const p = 'G:\\1Click Build\\Source\\' + fn;
  if (fs.existsSync(p)) {
    let c = fs.readFileSync(p, 'utf8');
    c = c.split('"Created by VPDU xa Thanh Phu"').join('"1Click Office - Tran Quoc Duoc (TTDVTH Phuong Tan Khai)"');
    fs.writeFileSync(p, c, 'utf8');
    console.log('Updated ' + fn);
  }
}
