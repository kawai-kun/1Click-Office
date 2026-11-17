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

// 1. Word
let w = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8');
let wLines = w.split(/\r?\n/);
for (let i = 0; i < wLines.length; i++) {
  if (wLines[i].includes('MsgBox "Don vi ban hanh dang chon cua khoi nay la:')) {
    wLines[i] = `MsgBoxUni ${toChrW("Đơn vị ban hành đang chọn của khối này là: ")} & TT_UnitLabelAt(grp, TT_UnitSlot(grp)) & vbCrLf & vbCrLf & _`;
    wLines[i+1] = `${toChrW("Tên này có vẻ thuộc hệ thống khác với khối vừa chọn (")} & TT_TenKhoiCoDau(TT_GetSetting()) & ").\" & vbCrLf & _`;
    wLines[i+2] = `${toChrW("Mỗi khối lưu danh sách đơn vị riêng. Hãy bấm 'Sửa thông tin' để nhập đúng cơ quan, hoặc 'Thêm đơn vị'.")}, vbInformation, ${toChrW("Đơn Vị Ban Hành")}`;
  }
  if (wLines[i].includes('MsgBox "Da luu file cau hinh (" & soDong & " muc):')) {
    wLines[i] = `MsgBoxUni ${toChrW("Đã lưu file cấu hình (")} & soDong & " " & ${toChrW("mục):")} & vbCrLf & d.FullName & vbCrLf & vbCrLf & _`;
    wLines[i+1] = `${toChrW("Gửi file này sang máy khác. Trên máy đó: vào tab THỂ THỨC VĂN BẢN, bấm 'Nhập cấu hình'.")}, vbInformation, ${toChrW("Xuất Cấu Hình")}`;
  }
  if (wLines[i].includes('MsgBox "Chua luu file. Van ban cau hinh van dang mo ("')) {
    wLines[i] = `MsgBoxUni ${toChrW("Chưa lưu file. Văn bản cấu hình vẫn đang mở (")} & soDong & " " & ${toChrW("mục).")} & vbCrLf & vbCrLf & _`;
    wLines[i+1] = `${toChrW("Hãy LƯU văn bản này thành file rồi gửi sang máy khác.")} & vbCrLf & _`;
    wLines[i+2] = `${toChrW("Trên máy đó: vào tab THỂ THỨC VĂN BẢN, bấm 'Nhập cấu hình'.")}, vbInformation, ${toChrW("Xuất Cấu Hình")}`;
  }
  if (wLines[i].includes('MsgBox "Van ban thu hai trong file da duoc chuan hoa nhung khong ghep lai duoc."')) {
    wLines[i] = `MsgBoxUni ${toChrW("Văn bản thứ hai trong file đã được chuẩn hóa nhưng không ghép lại được.")} & vbCrLf & _`;
    wLines[i+1] = `${toChrW("Nội dung đang nằm ở cửa sổ văn bản mới, hãy chép thủ công vào cuối văn bản.")}, vbExclamation, ${toChrW("Chuẩn Hóa Thể Thức")}`;
  }
}
fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', wLines.join('\r\n'), 'utf8');

// 2. Excel
let e = fs.readFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', 'utf8');
let eLines = e.split(/\r?\n/);
for (let i = 0; i < eLines.length; i++) {
  if (eLines[i].includes('MsgBox "Khong co lan chuan hoa nao de hoan tac')) {
    eLines[i] = `MsgBoxUni ${toChrW("Không có lần chuẩn hóa nào để hoàn tác (hoặc Excel đã đóng/mở lại sau lần chuẩn hóa).")}, vbInformation, ${toChrW("Hoàn Tác")}`;
  }
  if (eLines[i].includes('If Not silentMode Then MsgBox "Sheet dang bi khoa')) {
    eLines[i] = `If Not silentMode Then MsgBoxUni ${toChrW("Bảng tính (Sheet) đang bị khóa bảo vệ. Vui lòng bỏ khóa Sheet trước khi chuẩn hóa trang in.")}, vbExclamation, ${toChrW("Chuẩn Hóa Excel")}`;
  }
  if (eLines[i].includes('MsgBox "Co loi khi chuan hoa: " & errMsg')) {
    eLines[i] = `MsgBoxUni ${toChrW("Có lỗi khi chuẩn hóa: ")} & errMsg & vbCrLf & ${toChrW("Đã trả bảng tính về như trước khi chuẩn hóa.")}, vbExclamation, ${toChrW("Chuẩn Hóa Excel")}`;
  }
}
fs.writeFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', eLines.join('\r\n'), 'utf8');

console.log("Replaced perfectly by search.");
