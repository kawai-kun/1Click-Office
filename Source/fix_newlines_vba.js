const fs = require('fs');

function toChrW(str) {
  if (!str) return '""';
  let chunks = [];
  let curAscii = '';
  for (let i = 0; i < str.length; i++) {
    let ch = str[i];
    let code = str.charCodeAt(i);
    if (code === 34) {
      if (curAscii.length > 0) {
        chunks.push(JSON.stringify(curAscii));
        curAscii = '';
      }
      chunks.push('""""');
    } else if (code > 127 || code < 32) {
      if (curAscii.length > 0) {
        chunks.push(JSON.stringify(curAscii));
        curAscii = '';
      }
      if (code === 10) chunks.push('vbLf');
      else if (code === 13) chunks.push('vbCr');
      else chunks.push(`ChrW(${code})`);
    } else {
      curAscii += ch;
    }
  }
  if (curAscii.length > 0) chunks.push(JSON.stringify(curAscii));
  return chunks.length > 0 ? chunks.join(' & ') : '""';
}

let content = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8');

const newND30Msg = `    Dim msgND30 As String
    msgND30 = ${toChrW("Đã đánh số trang chuẩn Nghị định 30/2020/NĐ-CP thành công!")} & vbCrLf & _
              ${toChrW("- Vị trí: Ở giữa lề trên (Header)")} & vbCrLf & _
              ${toChrW("- Bắt đầu hiển thị: Từ trang 2 trở đi (ẩn trang 1)")} & vbCrLf & _
              ${toChrW("- Định dạng: Times New Roman, cỡ 13.5 pt.")}
    MsgBoxUni msgND30, vbInformation, "1Click Word"`;

// Replace broken MsgBoxUni in Word_DanhSoTrangND30
content = content.replace(
  /MsgBoxUni ChrW\(272\) & ChrW\(227\)[\s\S]*?"1Click Word"\s*End Sub\s*' === CHỨC NĂNG MỚI: ĐÁNH SỐ TỪ TRANG BẤT KỲ ===/,
  `${newND30Msg}\nEnd Sub\n\n' === CHỨC NĂNG MỚI: ĐÁNH SỐ TỪ TRANG BẤT KỲ ===`
);

// Replace MsgBoxUni in Word_DanhSoTrangBatKy
const newBatKyMsg = `    Dim msgBatKy As String
    msgBatKy = ${toChrW("Đã hoàn thành đánh số trang thành công!")} & vbCrLf & _
               ${toChrW("- Bắt đầu từ trang: ")} & trangBD & vbCrLf & _
               ${toChrW("- Số hiển thị bắt đầu: ")} & soBD & vbCrLf & _
               ${toChrW("- Các trang phía trước: Đã ẩn toàn bộ số trang.")}
    MsgBoxUni msgBatKy, vbInformation, "1Click Word"`;

content = content.replace(
  /MsgBoxUni ChrW\(272\) & ChrW\(227\) & " ho"[\s\S]*?"1Click Word"\s*Exit Sub\s*ThoatLoi:/,
  `${newBatKyMsg}\n    Exit Sub\nThoatLoi:`
);

// Also check lines around xoay ngang trang and bang bieu where \n was used:
// Line: ans = MsgBoxUni("H" & ChrW(7879)... \n...
content = content.replace(
  /ans = MsgBoxUni\("H" & ChrW\(7879\) & " th"[\s\S]*?vbYesNo \+ vbQuestion, "1Click Word"\)/,
  `Dim msgXoay As String
    msgXoay = ${toChrW("Hệ thống sẽ xoay NGANG trang hiện tại đang chọn để phù hợp chèn bảng biểu rộng.")} & vbCrLf & _
              ${toChrW("Các trang trước và sau vẫn giữ nguyên khổ DỌC.")} & vbCrLf & vbCrLf & _
              ${toChrW("Bạn có muốn tiếp tục không?")}
    ans = MsgBoxUni(msgXoay, vbYesNo + vbQuestion, "1Click Word")`
);

// Line: MsgBoxUni ChrW(272) & ChrW(227) & " chu" & ChrW(7849) & "n h" & ChrW(243) & "a to" & ChrW(224) & "n b" & ChrW(7897) & " " & doc.Tables.Count & " b" & ChrW(7843) & "ng bi" & ChrW(7875) & "u trong v" & ChrW(259) & "n b" & ChrW(7843) & "n:\n- T"
content = content.replace(
  /MsgBoxUni ChrW\(272\) & ChrW\(227\) & " chu"[\s\S]*?vbInformation, "1Click Word"/,
  `Dim msgTable As String
    msgTable = ${toChrW("Đã chuẩn hóa toàn bộ ")} & doc.Tables.Count & " " & ${toChrW("bảng biểu trong văn bản:")} & vbCrLf & _
               ${toChrW("- Tự động lặp lại dòng tiêu đề khi bảng dài qua nhiều trang.")} & vbCrLf & _
               ${toChrW("- Chống tình trạng gãy dòng dữ liệu giữa 2 trang.")} & vbCrLf & _
               ${toChrW("- Căn giữa bảng biểu theo lề trang in chuẩn.")}
    MsgBoxUni msgTable, vbInformation, "1Click Word"`
);

fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', content, 'utf8');
console.log("Successfully fixed all VBA multiline strings.");
