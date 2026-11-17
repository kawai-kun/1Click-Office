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

// 1. Word Module1_clean.vba
let wLines = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8').split('\r\n');
if (wLines.length === 1) wLines = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8').split('\n');

// 2269-2271 (0-indexed: 2269)
wLines[2269] = `MsgBoxUni ${toChrW("Đơn vị ban hành đang chọn của khối này là: ")} & TT_UnitLabelAt(grp, TT_UnitSlot(grp)) & vbCrLf & vbCrLf & _`;
wLines[2270] = `${toChrW("Tên này có vẻ thuộc hệ thống khác với khối vừa chọn (")} & TT_TenKhoiCoDau(TT_GetSetting()) & ").\" & vbCrLf & _`;
wLines[2271] = `${toChrW("Mỗi khối lưu danh sách đơn vị riêng. Hãy bấm 'Sửa thông tin' để nhập đúng cơ quan, hoặc 'Thêm đơn vị'.")}, vbInformation, ${toChrW("Đơn Vị Ban Hành")}`;

// 4392-4393
wLines[4392] = `MsgBoxUni ${toChrW("Đã lưu file cấu hình (")} & soDong & " " & ${toChrW("mục):")} & vbCrLf & d.FullName & vbCrLf & vbCrLf & _`;
wLines[4393] = `${toChrW("Gửi file này sang máy khác. Trên máy đó: vào tab THỂ THỨC VĂN BẢN, bấm 'Nhập cấu hình'.")}, vbInformation, ${toChrW("Xuất Cấu Hình")}`;

// 4395-4397
wLines[4395] = `MsgBoxUni ${toChrW("Chưa lưu file. Văn bản cấu hình vẫn đang mở (")} & soDong & " " & ${toChrW("mục).")} & vbCrLf & vbCrLf & _`;
wLines[4396] = `${toChrW("Hãy LƯU văn bản này thành file rồi gửi sang máy khác.")} & vbCrLf & _`;
wLines[4397] = `${toChrW("Trên máy đó: vào tab THỂ THỨC VĂN BẢN, bấm 'Nhập cấu hình'.")}, vbInformation, ${toChrW("Xuất Cấu Hình")}`;

// 9041-9042
wLines[9041] = `MsgBoxUni ${toChrW("Văn bản thứ hai trong file đã được chuẩn hóa nhưng không ghép lại được.")} & vbCrLf & _`;
wLines[9042] = `${toChrW("Nội dung đang nằm ở cửa sổ văn bản mới, hãy chép thủ công vào cuối văn bản.")}, vbExclamation, ${toChrW("Chuẩn Hóa Thể Thức")}`;

fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', wLines.join('\r\n'), 'utf8');
console.log("Word Module1_clean.vba line replacements completed.");

// 2. Excel Excel_Module1_final.vba
let eLines = fs.readFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', 'utf8').split('\r\n');
if (eLines.length === 1) eLines = fs.readFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', 'utf8').split('\n');

// 193
eLines[193] = `MsgBoxUni ${toChrW("Không có lần chuẩn hóa nào để hoàn tác (hoặc Excel đã đóng/mở lại sau lần chuẩn hóa).")}, vbInformation, ${toChrW("Hoàn Tác")}`;

// 211
eLines[211] = `If Not silentMode Then MsgBoxUni ${toChrW("Bảng tính (Sheet) đang bị khóa bảo vệ. Vui lòng bỏ khóa Sheet trước khi chuẩn hóa trang in.")}, vbExclamation, ${toChrW("Chuẩn Hóa Excel")}`;

// 458
eLines[458] = `MsgBoxUni ${toChrW("Có lỗi khi chuẩn hóa: ")} & errMsg & vbCrLf & ${toChrW("Đã trả bảng tính về như trước khi chuẩn hóa.")}, vbExclamation, ${toChrW("Chuẩn Hóa Excel")}`;

fs.writeFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', eLines.join('\r\n'), 'utf8');
console.log("Excel Excel_Module1_final.vba line replacements completed.");
