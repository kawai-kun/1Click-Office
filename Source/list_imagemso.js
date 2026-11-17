const fs = require('fs');
const content = fs.readFileSync('G:\\1Click Build\\Source\\customUI.xml', 'utf8');
const regex = /imageMso="([^"]+)"/g;
let match;
const set = new Set();
while ((match = regex.exec(content)) !== null) {
  set.add(match[1]);
}
console.log('Unique imageMso in customUI.xml:');
for (const item of set) {
  console.log(item);
}
