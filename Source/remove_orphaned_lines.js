const fs = require('fs');

let content = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8');
let lines = content.split(/\r?\n/);

console.log("Total lines before:", lines.length);

// Look for TT_RibbonRefresh and remove the orphaned MsgBoxUni right after its End Sub
for (let i = 0; i < lines.length; i++) {
  if (lines[i].includes('Private Sub TT_RibbonRefresh()')) {
    console.log("Found TT_RibbonRefresh at line", i + 1);
    // Scan ahead for End Sub
    for (let j = i; j < i + 10; j++) {
      if (lines[j].trim() === 'End Sub') {
        console.log("Found End Sub at line", j + 1);
        // Check next lines for orphaned MsgBoxUni
        if (lines[j+1] && lines[j+1].includes('MsgBoxUni')) {
          console.log("Removing orphaned lines starting at line", j + 2);
          // Count how many lines until the next Sub
          let k = j + 1;
          while (k < lines.length && !lines[k].trim().startsWith('Public Sub') && !lines[k].trim().startsWith('Private Sub') && !lines[k].trim().startsWith('Sub')) {
            console.log("Deleting:", lines[k].substring(0, 60));
            k++;
          }
          lines.splice(j + 1, k - (j + 1));
        }
        break;
      }
    }
    break;
  }
}

console.log("Total lines after:", lines.length);
fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', lines.join('\r\n'), 'utf8');
console.log("Successfully fixed Module1_clean.vba.");
