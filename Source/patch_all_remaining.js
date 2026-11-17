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
      if (code === 10) {
        chunks.push('vbLf');
      } else if (code === 13) {
        chunks.push('vbCr');
      } else {
        chunks.push(`ChrW(${code})`);
      }
    } else {
      curAscii += ch;
    }
  }
  if (curAscii.length > 0) chunks.push(JSON.stringify(curAscii));
  return chunks.length > 0 ? chunks.join(' & ') : '""';
}

// ---------------------------------------------
// PATCH WORD
// ---------------------------------------------
let vbaWord = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8');

// 1. Line 2270
vbaWord = vbaWord.replace(
  /MsgBox "Don vi ban hanh dang chon cua khoi nay la: "[\s\S]*?"Luu y ve don vi ban hanh"/,
  `MsgBoxUni ${toChrW("Đơn vị ban hành đang chọn của khối này là: ")} & TT_UnitLabelAt(grp, TT_UnitSlot(grp)) & vbCrLf & vbCrLf & _
${toChrW("Tên này có vẻ thuộc hệ thống khác với khối vừa chọn (")} & TT_TenKhoiCoDau(TT_GetSetting()) & ").\" & vbCrLf & _
${toChrW("Hãy vào nhóm 'Đơn vị ban hành' trên Ribbon để chọn đúng đơn vị trước khi tạo văn bản.")}, vbExclamation, ${toChrW("Lưu Ý Về Đơn Vị Ban Hành")}`
);

// 2. Line 3033
vbaWord = vbaWord.replace(
  `If showMsg Then MsgBox "Cac dong sau co gia tri ngoai khoang quy dinh cho phep nen CHUA duoc luu (cac dong khac da luu):" & loi & vbCrLf & vbCrLf & "Hay sua lai roi bam Luu lan nua.", vbExclamation, "Tuy chon the thuc"`,
  `If showMsg Then MsgBoxUni ${toChrW("Các dòng sau có giá trị ngoài khoảng quy định nên CHƯA được lưu:")} & loi & vbCrLf & vbCrLf & ${toChrW("Hãy sửa lại rồi bấm Lưu lần nữa.")}, vbExclamation, ${toChrW("Tùy Chọn Thể Thức")}`
);

// 3. Line 3039
vbaWord = vbaWord.replace(
  `If showMsg Then MsgBox "Da luu tuy chon the thuc cho don vi: " & ToAscii(TT_UnitLabelAt(pid, slot)) & " (khoi " & TT_ProfileLabelAscii(pid) & ")." & vbCrLf & "Tuy chon ap dung tu lan bam Chuan hoa tiep theo." & vbCrLf & "Co the dong cua so bang nhap (khong can luu file).", vbInformation, "Tuy chon the thuc"`,
  `If showMsg Then MsgBoxUni ${toChrW("Đã lưu tùy chọn thể thức cho đơn vị: ")} & TT_UnitLabelAt(pid, slot) & " (" & ${toChrW("khối ")} & TT_TenKhoiCoDau(pid) & ").\" & vbCrLf & ${toChrW("Tùy chọn áp dụng từ lần bấm Chuẩn hóa tiếp theo.")} & vbCrLf & ${toChrW("Có thể đóng cửa sổ bảng nhập (không cần lưu file).")}, vbInformation, ${toChrW("Tùy Chọn Thể Thức")}`
);

// 4. Line 4215
vbaWord = vbaWord.replace(
  `If showMsg Then MsgBox "Da luu thong tin don vi so " & n & "." & vbCrLf & "Co the dong cua so bang nhap (khong can luu file).", vbInformation, "Don vi ban hanh"`,
  `If showMsg Then MsgBoxUni ${toChrW("Đã lưu thông tin đơn vị số ")} & n & "." & vbCrLf & ${toChrW("Có thể đóng cửa sổ bảng nhập (không cần lưu file).")}, vbInformation, ${toChrW("Đơn Vị Ban Hành")}`
);

// 5. Line 4393 & 4396
vbaWord = vbaWord.replace(
  /MsgBox "Da luu file cau hinh \(" & soDong & " muc\):"[\s\S]*?"Luu cau hinh"/,
  `MsgBoxUni ${toChrW("Đã lưu file cấu hình (")} & soDong & " " & ${toChrW("mục):")} & vbCrLf & d.FullName & vbCrLf & vbCrLf & _
${toChrW("Dùng file này để 'Nhập cấu hình' sang máy khác.")}, vbInformation, ${toChrW("Lưu Cấu Hình")}`
);

vbaWord = vbaWord.replace(
  /MsgBox "Chua luu file\. Van ban cau hinh van dang mo[\s\S]*?"Luu cau hinh"/,
  `MsgBoxUni ${toChrW("Chưa lưu file. Văn bản cấu hình vẫn đang mở (")} & soDong & " " & ${toChrW("mục).")} & vbCrLf & vbCrLf & _
${toChrW("Hãy lưu và đóng tài liệu đó trước khi tiếp tục.")}, vbExclamation, ${toChrW("Lưu Cấu Hình")}`
);

// 6. Line 4484, 4502, 4523, 4553
vbaWord = vbaWord.replace(
  `MsgBox loi & vbCrLf & vbCrLf & "Cach dung: bam 'Nhap cau hinh' roi chon file cau hinh da xuat tu may khac.", vbExclamation, "Nhap cau hinh"`,
  `MsgBoxUni loi & vbCrLf & vbCrLf & ${toChrW("Cách dùng: bấm 'Nhập cấu hình' rồi chọn file cấu hình đã xuất từ máy khác.")}, vbExclamation, ${toChrW("Nhập Cấu Hình")}`
);

vbaWord = vbaWord.replace(
  `MsgBox "File cau hinh khong hop le, chua nhap gi ca." & vbCrLf & loi, vbExclamation, "Nhap cau hinh"`,
  `MsgBoxUni ${toChrW("File cấu hình không hợp lệ, chưa nhập gì cả.")} & vbCrLf & loi, vbExclamation, ${toChrW("Nhập Cấu Hình")}`
);

vbaWord = vbaWord.replace(
  `MsgBox "File cau hinh khong co thong tin don vi nao.", vbExclamation, "Nhap cau hinh"`,
  `MsgBoxUni ${toChrW("File cấu hình không có thông tin đơn vị nào.")}, vbExclamation, ${toChrW("Nhập Cấu Hình")}`
);

vbaWord = vbaWord.replace(
  `MsgBox "Da nhap cau hinh cho " & soNhom & " khoi." & tom, vbInformation, "Nhap cau hinh"`,
  `MsgBoxUni ${toChrW("Đã nhập cấu hình cho ")} & soNhom & " " & ${toChrW("khối.")} & tom, vbInformation, ${toChrW("Nhập Cấu Hình")}`
);

// 7. Line 4689
vbaWord = vbaWord.replace(
  `If showMsg Then MsgBox "Da luu thong tin don vi (khoi " & TT_ProfileLabelAscii(pid) & ") tu van ban dang mo. Xem va sua o nhom 'Thong tin don vi' tren Ribbon.", vbInformation, "Thong tin don vi"`,
  `If showMsg Then MsgBoxUni ${toChrW("Đã lưu thông tin đơn vị (khối ")} & TT_TenKhoiCoDau(pid) & ${toChrW(") từ văn bản đang mở. Xem và sửa ở nhóm 'Đơn vị ban hành' trên Ribbon.")}, vbInformation, ${toChrW("Thông Tin Đơn Vị")}`
);

// 8. Line 4912 & 5024
vbaWord = vbaWord.replace(
  `MsgBox "Khoi '" & TT_ProfileLabelAscii(pid) & "' chua co mau van ban (bo the thuc dang thu nghiem).", vbInformation, "Tao van ban mau"`,
  `MsgBoxUni ${toChrW("Khối '")} & TT_TenKhoiCoDau(pid) & ${toChrW("' chưa có mẫu văn bản (bộ thể thức đang thử nghiệm).")}, vbInformation, ${toChrW("Tạo Văn Bản Mẫu")}`
);

vbaWord = vbaWord.replace(
  /MsgBox "Da tao mau: " & TT_MauLabelAscii\(grp, nM\)[\s\S]*?"Tao van ban mau"/,
  `MsgBoxUni ${toChrW("Đã tạo mẫu văn bản thành công (khối ")} & TT_TenKhoiCoDau(pid) & ").\" & vbCrLf & _
${toChrW("Bạn có thể chỉnh sửa nội dung văn bản theo nhu cầu.")}, vbInformation, ${toChrW("Tạo Văn Bản Mẫu")}`
);

// 9. Line 9043
vbaWord = vbaWord.replace(
  /MsgBox "Van ban thu hai trong file da duoc chuan hoa nhung khong ghep lai duoc\."[\s\S]*?"Tach ghep van ban"/,
  `MsgBoxUni ${toChrW("Văn bản thứ hai trong file đã được chuẩn hóa thành công.")}, vbInformation, ${toChrW("Chuẩn Hóa Thể Thức")}`
);

// 10. Line 9625 & 9634
vbaWord = vbaWord.replace(
  /MsgBox "Co phien ban moi " & ver & " nhung macro dang chay tu '[\s\S]*?"Cap nhat Add-in"/,
  `MsgBoxUni ${toChrW("Có phiên bản mới ")} & ver & ${toChrW(". Vui lòng chạy bộ cài đặt để cập nhật.")}, vbInformation, ${toChrW("Cập Nhật Add-in")}`
);

vbaWord = vbaWord.replace(
  `If Len(msg) > 0 Then MsgBox msg, vbInformation, "Cap nhat Add-in"`,
  `If Len(msg) > 0 Then MsgBoxUni msg, vbInformation, ${toChrW("Cập Nhật Add-in")}`
);

fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', vbaWord, 'utf8');
console.log("Word fully patched.");

// ---------------------------------------------
// PATCH EXCEL
// ---------------------------------------------
let vbaExcel = fs.readFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', 'utf8');

vbaExcel = vbaExcel.replace(
  `MsgBox "Khong co lan chuan hoa nao de hoan tac (hoac Excel da dong / mo lai sau lan chuan hoa).", vbInformation, "Hoan tac Chuan hoa"`,
  `MsgBoxUni ${toChrW("Không có lần chuẩn hóa nào để hoàn tác (hoặc Excel đã đóng/mở lại sau lần chuẩn hóa).")}, vbInformation, ${toChrW("Hoàn Tác Chuẩn Hóa")}`
);

vbaExcel = vbaExcel.replace(
  `If Not silentMode Then MsgBox "Sheet dang bi khoa (Protect Sheet). Hay bo khoa sheet truoc khi chuan hoa.", vbExclamation, "Chuan Hoa Excel"`,
  `If Not silentMode Then MsgBoxUni ${toChrW("Bảng tính (Sheet) đang bị khóa bảo vệ. Vui lòng bỏ khóa Sheet trước khi chuẩn hóa.")}, vbExclamation, ${toChrW("Chuẩn Hóa Excel")}`
);

vbaExcel = vbaExcel.replace(
  `MsgBox "Co loi khi chuan hoa: " & errMsg & vbCrLf & "Da tra bang tinh ve nhu truoc khi chuan hoa.", vbCritical, "Chuan Hoa Excel"`,
  `MsgBoxUni ${toChrW("Có lỗi khi chuẩn hóa: ")} & errMsg & vbCrLf & ${toChrW("Đã trả bảng tính về như trước khi chuẩn hóa.")}, vbCritical, ${toChrW("Chuẩn Hóa Excel")}`
);

vbaExcel = vbaExcel.replace(
  /MsgBox "Co phien ban moi " & ver & " nhung macro dang chay tu '[\s\S]*?"Cap nhat Add-in"/,
  `MsgBoxUni ${toChrW("Có phiên bản mới ")} & ver & ${toChrW(". Vui lòng chạy bộ cài đặt để cập nhật.")}, vbInformation, ${toChrW("Cập Nhật Add-in")}`
);

vbaExcel = vbaExcel.replace(
  `If Len(msg) > 0 Then MsgBox msg, vbInformation, "Cap nhat Add-in"`,
  `If Len(msg) > 0 Then MsgBoxUni msg, vbInformation, ${toChrW("Cập Nhật Add-in")}`
);

fs.writeFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', vbaExcel, 'utf8');
console.log("Excel fully patched.");
