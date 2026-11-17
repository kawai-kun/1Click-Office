const fs = require('fs');
const path = require('path');

function scanDir(dir) {
  const files = fs.readdirSync(dir);
  for (const f of files) {
    const full = path.join(dir, f);
    const stat = fs.statSync(full);
    if (stat.isDirectory()) {
      scanDir(full);
    } else if (f.endsWith('.vba') || f.endsWith('.bas') || f.endsWith('.xml') || f.endsWith('.iss') || f.endsWith('.txt') || f.endsWith('.bat')) {
      const content = fs.readFileSync(full, 'utf8');
      const patterns = ['Thanh Phu', 'Thạnh Phú', 'VPDU xa', 'CuongBC', 'Cuong', 'Cường', 'Created by'];
      for (const p of patterns) {
        let lowerC = content.toLowerCase();
        let lowerP = p.toLowerCase();
        let idx = 0;
        while ((idx = lowerC.indexOf(lowerP, idx)) !== -1) {
          const start = Math.max(0, idx - 60);
          const end = Math.min(content.length, idx + 100);
          console.log(`MATCH [${p}] in ${full} at ${idx}:`);
          console.log('   ' + JSON.stringify(content.substring(start, end)));
          idx += p.length;
        }
      }
    }
  }
}

scanDir('G:\\1Click Build');
