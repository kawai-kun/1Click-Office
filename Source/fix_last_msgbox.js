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

// Fix Word
let w = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8');

// Line 2270
w = w.replace(
  'MsgBox "Don vi ban hanh dang chon cua khoi nay la: " & ToAscii(TT_UnitLabelAt(grp, TT_UnitSlot(grp))) & vbCrLf & vbCrLf & _\r\n"Ten nay co ve thuoc he thong khac voi khoi vua chon (" & TT_ProfileLabelAscii(TT_GetSetting()) & ")." & vbCrLf & _\r\n"Hay vao nhom \'Thong tin don vi\' tren Ribbon de chon dung don vi ban hanh cua khoi nay truoc khi tao van ban.", vbExclamation, "Luu y ve don vi ban hanh"',
  `MsgBoxUni ${toChrW("Đơn vị ban hành đang chọn của khối này là: ")} & TT_UnitLabelAt(grp, TT_UnitSlot(grp)) & vbCrLf & vbCrLf & _\r\n${toChrW("Tên này có vẻ thuộc hệ thống khác với khối vừa chọn (")} & TT_TenKhoiCoDau(TT_GetSetting()) & ").\" & vbCrLf & _\r\n${toChrW("Hãy vào nhóm 'Đơn vị ban hành' trên Ribbon để chọn đúng đơn vị trước khi tạo văn bản.")}, vbExclamation, ${toChrW("Lưu Ý Về Đơn Vị Ban Hành")}`
);
// Also without \r:
w = w.replace(
  'MsgBox "Don vi ban hanh dang chon cua khoi nay la: " & ToAscii(TT_UnitLabelAt(grp, TT_UnitSlot(grp))) & vbCrLf & vbCrLf & _\n"Ten nay co ve thuoc he thong khac voi khoi vua chon (" & TT_ProfileLabelAscii(TT_GetSetting()) & ")." & vbCrLf & _\n"Hay vao nhom \'Thong tin don vi\' tren Ribbon de chon dung don vi ban hanh cua khoi nay truoc khi tao van ban.", vbExclamation, "Luu y ve don vi ban hanh"',
  `MsgBoxUni ${toChrW("Đơn vị ban hành đang chọn của khối này là: ")} & TT_UnitLabelAt(grp, TT_UnitSlot(grp)) & vbCrLf & vbCrLf & _\n${toChrW("Tên này có vẻ thuộc hệ thống khác với khối vừa chọn (")} & TT_TenKhoiCoDau(TT_GetSetting()) & ").\" & vbCrLf & _\n${toChrW("Hãy vào nhóm 'Đơn vị ban hành' trên Ribbon để chọn đúng đơn vị trước khi tạo văn bản.")}, vbExclamation, ${toChrW("Lưu Ý Về Đơn Vị Ban Hành")}`
);

// Line 4393 & 4396
w = w.replace(
  'MsgBox "Da luu file cau hinh (" & soDong & " muc):" & vbCrLf & d.FullName & vbCrLf & vbCrLf & _\r\n"Dung file nay de \'Nhap cau hinh\' sang may khac.", vbInformation, "Luu cau hinh"',
  `MsgBoxUni ${toChrW("Đã lưu file cấu hình (")} & soDong & " " & ${toChrW("mục):")} & vbCrLf & d.FullName & vbCrLf & vbCrLf & _\r\n${toChrW("Dùng file này để 'Nhập cấu hình' sang máy khác.")}, vbInformation, ${toChrW("Lưu Cấu Hình")}`
);
w = w.replace(
  'MsgBox "Da luu file cau hinh (" & soDong & " muc):" & vbCrLf & d.FullName & vbCrLf & vbCrLf & _\n"Dung file nay de \'Nhap cau hinh\' sang may khac.", vbInformation, "Luu cau hinh"',
  `MsgBoxUni ${toChrW("Đã lưu file cấu hình (")} & soDong & " " & ${toChrW("mục):")} & vbCrLf & d.FullName & vbCrLf & vbCrLf & _\n${toChrW("Dùng file này để 'Nhập cấu hình' sang máy khác.")}, vbInformation, ${toChrW("Lưu Cấu Hình")}`
);

w = w.replace(
  'MsgBox "Chua luu file. Van ban cau hinh van dang mo (" & soDong & " muc)." & vbCrLf & vbCrLf & _\r\n"Luu file do roi hay dong cua so.", vbExclamation, "Luu cau hinh"',
  `MsgBoxUni ${toChrW("Chưa lưu file. Văn bản cấu hình vẫn đang mở (")} & soDong & " " & ${toChrW("mục).")} & vbCrLf & vbCrLf & _\r\n${toChrW("Lưu file đó rồi hãy đóng cửa sổ.")}, vbExclamation, ${toChrW("Lưu Cấu Hình")}`
);
w = w.replace(
  'MsgBox "Chua luu file. Van ban cau hinh van dang mo (" & soDong & " muc)." & vbCrLf & vbCrLf & _\n"Luu file do roi hay dong cua so.", vbExclamation, "Luu cau hinh"',
  `MsgBoxUni ${toChrW("Chưa lưu file. Văn bản cấu hình vẫn đang mở (")} & soDong & " " & ${toChrW("mục).")} & vbCrLf & vbCrLf & _\n${toChrW("Lưu file đó rồi hãy đóng cửa sổ.")}, vbExclamation, ${toChrW("Lưu Cấu Hình")}`
);

// Line 9042
w = w.replace(
  'MsgBox "Van ban thu hai trong file da duoc chuan hoa nhung khong ghep lai duoc." & vbCrLf & _\r\n"Hay kiem tra lai phan sau cua van ban.", vbExclamation, "Tach ghep van ban"',
  `MsgBoxUni ${toChrW("Văn bản thứ hai trong file đã được chuẩn hóa nhưng không ghép lại được.")} & vbCrLf & _\r\n${toChrW("Hãy kiểm tra lại phần sau của văn bản.")}, vbExclamation, ${toChrW("Tách Ghép Văn Bản")}`
);
w = w.replace(
  'MsgBox "Van ban thu hai trong file da duoc chuan hoa nhung khong ghep lai duoc." & vbCrLf & _\n"Hay kiem tra lai phan sau cua van ban.", vbExclamation, "Tach ghep van ban"',
  `MsgBoxUni ${toChrW("Văn bản thứ hai trong file đã được chuẩn hóa nhưng không ghép lại được.")} & vbCrLf & _\n${toChrW("Hãy kiểm tra lại phần sau của văn bản.")}, vbExclamation, ${toChrW("Tách Ghép Văn Bản")}`
);

fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', w, 'utf8');

// Fix Excel
let e = fs.readFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', 'utf8');

e = e.replace(
  'MsgBox "Khong co lan chuan hoa nao de hoan tac (hoac Excel da dong / mo lai sau lan chuan hoa).", vbInformation, "Hoan tac Chuan hoa"',
  `MsgBoxUni ${toChrW("Không có lần chuẩn hóa nào để hoàn tác (hoặc Excel đã đóng/mở lại sau lần chuẩn hóa).")}, vbInformation, ${toChrW("Hoàn Tác Chuẩn Hóa")}`
);

e = e.replace(
  'If Not silentMode Then MsgBox "Sheet dang bi khoa (Protect Sheet). Hay bo khoa sheet truoc khi chuan hoa.", vbExclamation, "Chuan Hoa Excel"',
  `If Not silentMode Then MsgBoxUni ${toChrW("Bảng tính (Sheet) đang bị khóa bảo vệ. Vui lòng bỏ khóa Sheet trước khi chuẩn hóa.")}, vbExclamation, ${toChrW("Chuẩn Hóa Excel")}`
);

e = e.replace(
  'MsgBox "Co loi khi chuan hoa: " & errMsg & vbCrLf & "Da tra bang tinh ve nhu truoc khi chuan hoa.", vbCritical, "Chuan Hoa Excel"',
  `MsgBoxUni ${toChrW("Có lỗi khi chuẩn hóa: ")} & errMsg & vbCrLf & ${toChrW("Đã trả bảng tính về như trước khi chuẩn hóa.")}, vbCritical, ${toChrW("Chuẩn Hóa Excel")}`
);

fs.writeFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', e, 'utf8');

console.log("Finished fix_last_msgbox.");
