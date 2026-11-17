const fs = require('fs');
const path = require('path');

function toChrW(str) {
  if (!str) return '""';
  let chunks = [];
  let curAscii = '';
  for (let i = 0; i < str.length; i++) {
    let ch = str[i];
    let code = str.charCodeAt(i);
    if (code === 34) { // quote
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

function toVbaLines(str, varName = "s") {
  let lines = str.split('\n');
  let out = [];
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].replace(/\r$/, '');
    let piece = toChrW(line);
    if (i === 0) {
      out.push(`    ${varName} = ${piece}`);
    } else {
      out.push(`    ${varName} = ${varName} & vbCrLf & ${piece}`);
    }
  }
  return out.join('\n');
}

// ---------------------------------------------------------
// 1. UPDATE Word Module1_clean.vba
// ---------------------------------------------------------
console.log("Reading Module1_clean.vba...");
let vbaWord = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8');

// A. Update RibbonThongTinTacGia in Word to include phone 0984.724.131
const authorWordText = `1CLICK OFFICE 3.0 ENTERPRISE
Bộ Tiện Ích Chuẩn Hóa Văn Bản & Bảng Tính Hành Chính

- Tác giả: Trần Quốc Được
- Đơn vị: Trung tâm Dịch vụ Tổng hợp Phường Tân Khai
- Điện thoại / Zalo: 0984.724.131
- Trạng thái bản quyền: VĨNH VIỄN (Đã kích hoạt)

Hỗ trợ kỹ thuật & Đóng góp ý kiến:
Vui lòng liên hệ Tác giả Trần Quốc Được - ĐT: 0984.724.131 (UBND / TTDVTH Phường Tân Khai).`;

const newSubAuthorWord = `Public Sub RibbonThongTinTacGia(control As Object)
    Dim s As String
${toVbaLines(authorWordText, "s")}
    MsgBoxUni s, vbInformation, ${toChrW("1Click Office - Thông Tin Tác Giả")}
End Sub`;

vbaWord = vbaWord.replace(/Public Sub RibbonThongTinTacGia\(control As Object\)[\s\S]*?End Sub/, newSubAuthorWord);

// B. Update TT_TenKhoiCoDau and TT_LegalBasisCoDau
const newTenKhoiCoDau = `Private Function TT_TenKhoiCoDau(ByVal pid As String) As String
Select Case pid
Case "DANG": TT_TenKhoiCoDau = ${toChrW("Đảng")}
Case "CHINH_QUYEN": TT_TenKhoiCoDau = ${toChrW("Chính quyền")}
Case "CONG_AN": TT_TenKhoiCoDau = ${toChrW("Công an")}
Case "QUAN_DOI": TT_TenKhoiCoDau = ${toChrW("Quân đội")}
Case "DOAN": TT_TenKhoiCoDau = ${toChrW("Đoàn TNCS Hồ Chí Minh")}
Case "MTTQ": TT_TenKhoiCoDau = ${toChrW("Mặt trận Tổ quốc & Đoàn thể")}
Case "AUTO": TT_TenKhoiCoDau = ${toChrW("Tự động nhận diện")}
Case Else: TT_TenKhoiCoDau = pid
End Select
End Function

Public Function TT_LegalBasisCoDau(ByVal pid As String) As String
Select Case pid
Case "DANG": TT_LegalBasisCoDau = ${toChrW("Quy định số 399-QĐ/TW ngày 09/01/2026 của Ban Bí thư; Hướng dẫn số 05-HD/VPTW ngày 27/5/2026 của Văn phòng Trung ương Đảng")}
Case "CHINH_QUYEN": TT_LegalBasisCoDau = ${toChrW("Nghị định số 30/2020/NĐ-CP ngày 05/3/2020 của Chính phủ về công tác văn thư (Phụ lục I: thể thức, kỹ thuật trình bày văn bản hành chính)")}
Case "CONG_AN": TT_LegalBasisCoDau = ${toChrW("Thông tư số 26/2019/TT-BCA về công tác văn thư trong Công an nhân dân; Thể thức theo Điều 8, 9 Nghị định 30/2020/NĐ-CP")}
Case "QUAN_DOI": TT_LegalBasisCoDau = ${toChrW("Thể thức, kỹ thuật trình bày theo Nghị định 30/2020/NĐ-CP; Quy chế công tác văn thư trong Bộ Quốc phòng (Thông tư 68/2023/TT-BQP)")}
Case "DOAN": TT_LegalBasisCoDau = ${toChrW("Hướng dẫn số 29-HD/TWĐTN-VP ngày 29/10/2013 của Ban Bí thư Trung ương Đoàn về thể thức văn bản của Đoàn TNCS Hồ Chí Minh")}
Case "MTTQ": TT_LegalBasisCoDau = ${toChrW("Quyết định số 207/QĐ-MTTW-BTT ngày 29/8/2025 của Ban Thường trực Ủy ban Trung ương MTTQ Việt Nam")}
Case Else: TT_LegalBasisCoDau = ""
End Select
End Function`;

vbaWord = vbaWord.replace(/Private Function TT_TenKhoiCoDau\(ByVal pid As String\)[\s\S]*?End Function/, newTenKhoiCoDau);

// C. Update TT_ReportHeader
const newReportHeader = `Public Function TT_ReportHeader() As String
Dim s As String
s = ${toChrW("Khối thể thức: ")} & TT_TenKhoiCoDau(TT_Profile())
Select Case mTTSource
Case "TU_NHAN_DIEN": s = s & " " & ${toChrW("(tự động nhận diện)")}
Case "MAC_DINH": s = s & " " & ${toChrW("(không nhận diện được, dùng khối mặc định)")}
Case Else: s = s & " " & ${toChrW("(người dùng chọn)")}
End Select
s = s & vbCrLf & ${toChrW("Căn cứ: ")} & TT_LegalBasisCoDau(TT_Profile())
Select Case TT_S("STATUS")
Case "experimental": s = s & vbCrLf & ${toChrW("Lưu ý: bộ thể thức khối này đang THỬ NGHIỆM, chỉ kiểm tra nới lỏng.")}
Case "needs_review": s = s & vbCrLf & ${toChrW("Lưu ý: bộ thể thức khối này CẦN ĐỐI CHIẾU THÊM văn bản gốc.")}
End Select
TT_ReportHeader = s
End Function`;

vbaWord = vbaWord.replace(/Public Function TT_ReportHeader\(\) As String[\s\S]*?End Function/, newReportHeader);

// D. Update TT_MismatchWarning
const newMismatchWarning = `Public Function TT_MismatchWarning() As String
TT_MismatchWarning = ""
If mTTSource <> "CHON" Or Not mTTDetectedStrong Or Len(mTTDetected) = 0 Then Exit Function
If mTTDetected = mTTProfile Then Exit Function
If (TT_RawS(mTTDetected, "MOTTO_KIND") = "DANG") <> (TT_RawS(mTTProfile, "MOTTO_KIND") = "DANG") Then
TT_MismatchWarning = ${toChrW("[Khối thể thức] Đang chọn khối '")} & TT_TenKhoiCoDau(mTTProfile) & _
${toChrW("' nhưng dòng tiêu ngữ trên văn bản thuộc khối '")} & TT_TenKhoiCoDau(mTTDetected) & _
${toChrW("'. Kiểm tra lại mục Khối trên Ribbon (hoặc chọn Tự động nhận diện).")}
End If
End Function`;

vbaWord = vbaWord.replace(/Public Function TT_MismatchWarning\(\) As String[\s\S]*?End Function/, newMismatchWarning);

// E. Update KiemTra_LogicTheThuc warnings to 100% Accented Vietnamese
// Let's replace each warning in KiemTra_LogicTheThuc
vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Du lieu trong] Phat hien con dau '...' chua dien day du thong tin."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Dữ liệu trống] Phát hiện còn dấu '...' chưa điền đầy đủ thông tin.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Cho trong] Con " & phCount & " cho can dien trong dau [ ] (vi du: [Ho va ten])."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Chỗ trống] Còn ")} & phCount & " " & ${toChrW("chỗ cần điền trong dấu [ ] (ví dụ: [Họ và tên]).")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Thu tu Dieu] Nhay coc hoac sai thu tu: Ky vong 'Dieu " & nextExpectedDieu & "' nhung gap 'Dieu " & actualDieu & "'."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Thứ tự Điều] Nhảy cóc hoặc sai thứ tự: Kỳ vọng 'Điều ")} & nextExpectedDieu & "' " & ${toChrW("nhưng gặp 'Điều ")} & actualDieu & "'."`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [De muc La Ma] Nhay coc thu tu: Ky vong muc '" & TT_SoLaMa(nextExpectedRoman) & "' nhung gap muc '" & romanName & "'."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Đề mục La Mã] Nhảy cóc thứ tự: Kỳ vọng mục '")} & TT_SoLaMa(nextExpectedRoman) & "' " & ${toChrW("nhưng gặp mục '")} & romanName & "'."`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [De muc So] Nhay coc thu tu khoan: Ky vong muc '" & nextExpectedArabic & ".' nhung gap '" & actualArab & ".' (" & Left(pAsc, 35) & "...)."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Đề mục Số] Nhảy cóc thứ tự khoản: Kỳ vọng mục '")} & nextExpectedArabic & ".' " & ${toChrW("nhưng gặp '")} & actualArab & ".' (" & Left(pAsc, 35) & "...)."`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [So van ban] Dang de trong 3 khoang trang ('" & foundSoRaw & "') cho van thu cap so."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Số văn bản] Đang để trống khoảng trắng ('")} & foundSoRaw & ${toChrW("') cho văn thư cấp số.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [So, ky hieu] Ky hieu dang ghi theo kieu Nha nuoc ('" & foundSoRaw & "'); van ban Dang ghi dang 'So 12-QD/DU'."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Số, ký hiệu] Ký hiệu đang ghi theo kiểu Nhà nước ('")} & foundSoRaw & ${toChrW("'); văn bản Đảng ghi dạng 'Số 12-QĐ/ĐU'.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [So, ky hieu] Ky hieu dang ghi theo kieu Dang ('" & foundSoRaw & "'); van ban hanh chinh ghi dang 'So: 12/QD-UBND' (ND 30/2020)."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Số, ký hiệu] Ký hiệu đang ghi theo kiểu Đảng ('")} & foundSoRaw & ${toChrW("'); văn bản hành chính ghi dạng 'Số: 12/QĐ-UBND' (NĐ 30/2020).")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [So van ban] Khong tim thay dong So, ky hieu van ban o phan Tieu de."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Số văn bản] Không tìm thấy dòng Số, ký hiệu văn bản ở phần Tiêu đề.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Ngay thang] Dang de trong ngay/thang ban hanh ('" & foundDateRaw & "') cho ngay ky duyet."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Ngày tháng] Đang để trống ngày/tháng ban hành ('")} & foundDateRaw & ${toChrW("') cho ngày ký duyệt.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Ngay thang] Ngay thang khong hop le: ngay " & dInt & " thang " & mInt & "."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Ngày tháng] Ngày tháng không hợp lệ: ngày ")} & dInt & " " & ${toChrW("tháng ")} & mInt & "."`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Ngay thang] Thang " & mInt & " chi co toi da 30 ngay (dang ghi ngay " & dInt & ")."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Ngày tháng] Tháng ")} & mInt & " " & ${toChrW("chỉ có tối đa 30 ngày (đang ghi ngày ")} & dInt & ")."`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Ngay thang] Thang 02 khong the co ngay " & dInt & "."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Ngày tháng] Tháng 02 không thể có ngày ")} & dInt & "."`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Loai van ban] Ky hieu la Cong van ('CV') nhung lai co tieu de '" & foundTitle & "' o giua bai."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Loại văn bản] Ký hiệu là Công văn ('CV') nhưng lại có tiêu đề '")} & foundTitle & ${toChrW("' ở giữa bài.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Trich yeu Cong van] Cong van dang thieu dong trich yeu 'V/v...' duoi so van ban."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Trích yếu Công văn] Công văn đang thiếu dòng trích yếu 'V/v...' dưới số văn bản.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Kinh gui] Cong van dang thieu muc 'Kinh gui:'."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Kính gửi] Công văn đang thiếu mục 'Kính gửi:'.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Thieu tieu de] Ky hieu van ban la '" & foundCode & "' nhung thieu dong ten loai van ban ('" & Replace(expectTitle, "|", " / ") & "') o giua trang."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Thiếu tiêu đề] Ký hiệu văn bản là '")} & foundCode & "' " & ${toChrW("nhưng thiếu dòng tên loại văn bản ('")} & Replace(expectTitle, "|", " / ") & ${toChrW("') ở giữa trang.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Mau thuan loai VB] Ky hieu ghi la '" & foundCode & "' nhung tieu de van ban lai ghi la '" & foundTitle & "'."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Mâu thuẫn loại VB] Ký hiệu ghi là '")} & foundCode & "' " & ${toChrW("nhưng tiêu đề văn bản lại ghi là '")} & foundTitle & "'."`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Cau truc Quyet dinh] Van ban Quyet dinh nhung thieu dong 'QUYET DINH:' hoac thieu cac 'Dieu 1.', 'Dieu 2.'."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Cấu trúc Quyết định] Văn bản Quyết định nhưng thiếu dòng 'QUYẾT ĐỊNH:' hoặc thiếu các 'Điều 1.', 'Điều 2.'.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Logic Noi nhan] Van ban Quyet dinh co Dieu khoan thi hanh nhung phan Noi nhan thieu dong '- Nhu Dieu ...;'."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Logic Nơi nhận] Văn bản Quyết định có Điều khoản thi hành nhưng phần Nơi nhận thiếu dòng '- Như Điều ...;'.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Tham quyen ky] Co quan ban hanh o O 1 la tap the ('" & foundIssuingOrgan & "') nhung chu ky thieu tham quyen 'T/M ...'."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Thẩm quyền ký] Cơ quan ban hành ở Ô 1 là tập thể ('")} & foundIssuingOrgan & ${toChrW("') nhưng chữ ký thiếu thẩm quyền 'T/M ...'.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Tham quyen ky] Co quan ban hanh o O 1 la '" & foundIssuingOrgan & "' (che do thu truong) nen khong duoc ky 'T/M' (" & foundSignAuth & ")."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Thẩm quyền ký] Cơ quan ban hành ở Ô 1 là '")} & foundIssuingOrgan & ${toChrW("' (chế độ thủ trưởng) nên không được ký 'T/M' (")} & foundSignAuth & ")."`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Phan 'Noi nhan:' dang thieu dong luu o cuoi (vi du theo HD 29-HD/TWDTN-VP: '- Luu VP.')."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Nơi nhận] Phần 'Nơi nhận:' đang thiếu dòng lưu ở cuối (ví dụ theo HD 29-HD/TWĐTN-VP: '- Lưu VP.').")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Phan 'Noi nhan:' dang thieu dong luu o cuoi (vi du: '- Luu Van phong Tinh uy.')."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Nơi nhận] Phần 'Nơi nhận:' đang thiếu dòng lưu ở cuối (ví dụ: '- Lưu Văn phòng Tỉnh ủy.').")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Phan 'Noi nhan:' dang thieu dong '- Luu: VT...' o cuoi."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Nơi nhận] Phần 'Nơi nhận:' đang thiếu dòng '- Lưu: VT...' ở cuối.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Dong luu nen ghi noi luu (vi du theo HD 05-HD/VPTW: '- Luu Van phong Tinh uy.')."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Nơi nhận] Dòng lưu nên ghi nơi lưu (ví dụ theo HD 05-HD/VPTW: '- Lưu Văn phòng Tỉnh ủy.').")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Noi nhan] Dong luu nen bo sung ky hieu 'VT' (Van thu) (vi du: '- Luu: VT, ...')."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Nơi nhận] Dòng lưu nên bổ sung ký hiệu 'VT' (Văn thư) (ví dụ: '- Lưu: VT, ...').")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Ten co quan] '" & foundIssuingOrgan & "' la co quan tham muu, giup viec: dong tren phai ghi ten cap uy ma co quan do truc thuoc (HD 05-HD/VPTW muc 2.1.3), vi du 'DANG UY XA ...'."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Tên cơ quan] '")} & foundIssuingOrgan & ${toChrW("' là cơ quan tham mưu, giúp việc: dòng trên phải ghi tên cấp ủy mà cơ quan đó trực thuộc (HD 05-HD/VPTW mục 2.1.3), ví dụ 'ĐẢNG ỦY XÃ ...'.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Ten co quan] Van ban cua '" & foundIssuingOrgan & "' phai ghi ten dang bo cap tren truc tiep o dong tren (HD 05-HD/VPTW muc 2.1.2), vi du 'DANG BO TINH ...' / 'DANG BO THANH PHO ...'."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Tên cơ quan] Văn bản của '")} & foundIssuingOrgan & ${toChrW("' phải ghi tên đảng bộ cấp trên trực tiếp ở dòng trên (HD 05-HD/VPTW mục 2.1.2), ví dụ 'ĐẢNG BỘ TỈNH ...' / 'ĐẢNG BỘ THÀNH PHỐ ...'.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Ten co quan] Dong tren dang ghi '" & foundSuperiorOrgan & "' (mo hinh cap huyen cu). Dang bo xa, phuong, dac khu truc thuoc truc tiep Tinh uy, Thanh uy: ghi 'DANG BO TINH ...' / 'DANG BO THANH PHO ...'."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Tên cơ quan] Dòng trên đang ghi '")} & foundSuperiorOrgan & ${toChrW("' (mô hình cấp huyện cũ). Đảng bộ xã, phường trực thuộc trực tiếp Tỉnh ủy/Thành ủy: ghi 'ĐẢNG BỘ TỈNH ...' / 'ĐẢNG BỘ THÀNH PHỐ ...'.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Ten co quan] Dong tren ghi ten DANG BO cap tren truc tiep (vi du 'DANG BO TINH ...'), khong ghi '" & foundSuperiorOrgan & "' (HD 05-HD/VPTW muc 2.1.2)."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Tên cơ quan] Dòng trên ghi tên ĐẢNG BỘ cấp trên trực tiếp (ví dụ 'ĐẢNG BỘ TỈNH ...'), không ghi '")} & foundSuperiorOrgan & ${toChrW("' (HD 05-HD/VPTW mục 2.1.2).")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Tieu ngu] Van ban " & TT_ProfileLabelAscii(TT_Profile()) & " chua co dong '" & UCase(ToAscii(TT_TextTieuNguDang())) & "' o dau trang."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Tiêu ngữ] Văn bản ")} & TT_TenKhoiCoDau(TT_Profile()) & " " & ${toChrW("chưa có dòng '")} & UCase(TT_TextTieuNguDang()) & ${toChrW("' ở đầu trang.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Quoc hieu] Van ban hanh chinh chua co Quoc hieu 'CONG HOA XA HOI CHU NGHIA VIET NAM' va Tieu ngu."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Quốc hiệu] Văn bản hành chính chưa có Quốc hiệu 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM' và Tiêu ngữ.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Quoc hieu] Thieu dong Tieu ngu 'Doc lap - Tu do - Hanh phuc' duoi Quoc hieu."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Quốc hiệu] Thiếu dòng Tiêu ngữ 'Độc lập - Tự do - Hạnh phúc' dưới Quốc hiệu.")}`
);

vbaWord = vbaWord.replace(
  `warnings = warnings & vbCrLf & warnCount & ". [Nguoi ky] Ho ten nguoi ky ('" & ToAscii(signerName) & "') chua ghi cap bac ham phia truoc (vi du: 'Dai ta Nguyen Van A'). Kiem tra lai neu don vi co quy dinh ghi cap bac."`,
  `warnings = warnings & vbCrLf & warnCount & ". " & ${toChrW("[Người ký] Họ tên người ký ('")} & signerName & ${toChrW("') chưa ghi cấp bậc hàm phía trước (ví dụ: 'Đại tá Nguyễn Văn A'). Kiểm tra lại nếu đơn vị có quy định ghi cấp bậc.")}`
);

// F. Update Report Dialog at lines 6740-6750
const oldReportBlockRegex = /If Not silentMode Then\s*If Len\(strWarnings\) > 0 Then\s*MsgBox TT_HienThiXuongDong\([\s\S]*?End If\s*End If/;
const newReportBlock = `If Not silentMode Then
        Dim strTieuDeBaoCao As String
        Dim strFooterTacGia As String
        strFooterTacGia = ${toChrW("1Click Office - Tác giả: Trần Quốc Được (TTDVTH Phường Tân Khai) - ĐT: 0984.724.131")}
        If Len(strWarnings) > 0 Then
            strTieuDeBaoCao = ${toChrW("Báo Cáo Rà Soát Logic & Thể Thức")}
            MsgBoxUni ${toChrW("ĐÃ CHUẨN HÓA THỂ THỨC VĂN BẢN!")} & vbCrLf & _
                      TT_ReportHeader() & vbCrLf & vbCrLf & _
                      ${toChrW("Phát hiện các điểm bất hợp lý về LOGIC cần kiểm tra trước khi trình ký:")} & vbCrLf & _
                      strWarnings & vbCrLf & vbCrLf & _
                      strFooterTacGia, vbExclamation, strTieuDeBaoCao
        Else
            strTieuDeBaoCao = ${toChrW("Chuẩn Hóa Thể Thức Thành Công")}
            MsgBoxUni ${toChrW("ĐÃ CHUẨN HÓA THỂ THỨC VĂN BẢN!")} & vbCrLf & _
                      TT_ReportHeader() & vbCrLf & vbCrLf & _
                      ${toChrW("Văn bản đạt chuẩn 100%, không phát hiện lỗi logic.")} & vbCrLf & vbCrLf & _
                      strFooterTacGia, vbInformation, strTieuDeBaoCao
        End If
    End If`;

vbaWord = vbaWord.replace(oldReportBlockRegex, newReportBlock);

// G. Replace other unaccented MsgBox in Word with MsgBoxUni & accented Vietnamese:
vbaWord = vbaWord.replace(
  `MsgBox "Da chon khoi: " & TT_ProfileLabelAscii(TT_GetSetting()), vbInformation, "Khoi the thuc van ban"`,
  `MsgBoxUni ${toChrW("Đã chọn khối: ")} & TT_TenKhoiCoDau(TT_GetSetting()), vbInformation, ${toChrW("Khối Thể Thức Văn Bản")}`
);

vbaWord = vbaWord.replace(
  `If showMsg Then MsgBox "Bang tuy chon da bi sua mat dong. Hay dong cua so nay va mo lai bang tuy chon.", vbExclamation, "Tuy chon the thuc"`,
  `If showMsg Then MsgBoxUni ${toChrW("Bảng tùy chọn đã bị sửa mất dòng. Hãy đóng cửa sổ này và mở lại bảng tùy chọn.")}, vbExclamation, ${toChrW("Tùy Chọn Thể Thức")}`
);

vbaWord = vbaWord.replace(
  `MsgBox "Da du " & TT_UnitMax() & " don vi cho khoi nay. Hay xoa bot don vi khong dung truoc khi them.", vbExclamation, "Don vi ban hanh"`,
  `MsgBoxUni ${toChrW("Đã đủ ")} & TT_UnitMax() & " " & ${toChrW("đơn vị cho khối này. Hãy xóa bớt đơn vị không dùng trước khi thêm.")}, vbExclamation, ${toChrW("Đơn Vị Ban Hành")}`
);

vbaWord = vbaWord.replace(
  `MsgBox "Chi con 1 don vi, khong xoa duoc. Muon doi thong tin thi sua truc tiep cac o.", vbInformation, "Don vi ban hanh"`,
  `MsgBoxUni ${toChrW("Chỉ còn 1 đơn vị, không thể xóa. Muốn đổi thông tin thì sửa trực tiếp các ô.")}, vbInformation, ${toChrW("Đơn Vị Ban Hành")}`
);

vbaWord = vbaWord.replace(
  `If showMsg Then MsgBox "Chua mo bang thong tin don vi. Bam nut 'Sua thong tin' truoc, dien xong roi bam 'Luu don vi'.", vbInformation, "Don vi ban hanh"`,
  `If showMsg Then MsgBoxUni ${toChrW("Chưa mở bảng thông tin đơn vị. Bấm nút 'Sửa thông tin' trước, điền xong rồi bấm 'Lưu đơn vị'.")}, vbInformation, ${toChrW("Đơn Vị Ban Hành")}`
);

vbaWord = vbaWord.replace(
  `If showMsg Then MsgBox "Bang thong tin don vi da bi sua mat dong. Hay dong cua so nay va bam 'Sua thong tin' lai.", vbExclamation, "Don vi ban hanh"`,
  `If showMsg Then MsgBoxUni ${toChrW("Bảng thông tin đơn vị đã bị sửa mất dòng. Hãy đóng cửa sổ này và bấm 'Sửa thông tin' lại.")}, vbExclamation, ${toChrW("Đơn Vị Ban Hành")}`
);

vbaWord = vbaWord.replace(
  `If showMsg Then MsgBox "Chua nhap ten co quan ban hanh (dong 2). Hay dien roi bam Luu, hoac dong cua so nay neu khong muon them don vi.", vbExclamation, "Don vi ban hanh"`,
  `If showMsg Then MsgBoxUni ${toChrW("Chưa nhập tên cơ quan ban hành (dòng 2). Hãy điền rồi bấm Lưu, hoặc đóng cửa sổ này nếu không muốn thêm.")}, vbExclamation, ${toChrW("Đơn Vị Ban Hành")}`
);

vbaWord = vbaWord.replace(
  `If showMsg Then MsgBox "Da du so don vi cho khoi nay, khong them duoc.", vbExclamation, "Don vi ban hanh"`,
  `If showMsg Then MsgBoxUni ${toChrW("Đã đủ số đơn vị cho khối này, không thể thêm được nữa.")}, vbExclamation, ${toChrW("Đơn Vị Ban Hành")}`
);

vbaWord = vbaWord.replace(
  `If showMsg Then MsgBox "Khong nhan dien duoc khoi cua van ban dang mo (can co bang tieu de da chuan hoa).", vbExclamation, "Thong tin don vi"`,
  `If showMsg Then MsgBoxUni ${toChrW("Không nhận diện được khối của văn bản đang mở (cần có bảng tiêu đề đã chuẩn hóa).")}, vbExclamation, ${toChrW("Thông Tin Đơn Vị")}`
);

vbaWord = vbaWord.replace(
  `If showMsg Then MsgBox "Van ban dang mo chua co bang tieu de (Ten co quan, So ky hieu). Hay Chuan hoa the thuc truoc.", vbExclamation, "Thong tin don vi"`,
  `If showMsg Then MsgBoxUni ${toChrW("Văn bản đang mở chưa có bảng tiêu đề (Tên cơ quan, Số ký hiệu). Hãy Chuẩn hóa thể thức trước.")}, vbExclamation, ${toChrW("Thông Tin Đơn Vị")}`
);

vbaWord = vbaWord.replace(
  `MsgBox "Thong tin don vi nam o tab 'THE THUC VAN BAN' tren thanh Ribbon (nhom 'Thong Tin Don Vi').", vbInformation, "Thong tin don vi"`,
  `MsgBoxUni ${toChrW("Thông tin đơn vị nằm ở tab 'THỂ THỨC VĂN BẢN' trên thanh Ribbon (nhóm 'Đơn vị ban hành').")}, vbInformation, ${toChrW("Thông Tin Đơn Vị")}`
);

vbaWord = vbaWord.replace(
  `If hoiKhiKhongCo Then MsgBox "Khong ket noi duoc may chu cap nhat (GitHub). Kiem tra mang roi thu lai." & vbCrLf & "Phien ban dang dung: " & CN_PhienBan(), vbExclamation, "Cap nhat Add-in"`,
  `If hoiKhiKhongCo Then MsgBoxUni ${toChrW("Không kết nối được máy chủ cập nhật (GitHub). Vui lòng kiểm tra mạng Internet rồi thử lại.")} & vbCrLf & ${toChrW("Phiên bản đang dùng: ")} & CN_PhienBan(), vbExclamation, ${toChrW("Cập Nhật Add-in")}`
);

vbaWord = vbaWord.replace(
  `If hoiKhiKhongCo Then MsgBox "Ban dang dung phien ban moi nhat: " & CN_PhienBan(), vbInformation, "Cap nhat Add-in"`,
  `If hoiKhiKhongCo Then MsgBoxUni ${toChrW("Bạn đang dùng phiên bản mới nhất: ")} & CN_PhienBan() & " Enterprise.", vbInformation, ${toChrW("Kiểm Tra Cập Nhật")}`
);

fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', vbaWord, 'utf8');
console.log("Updated Module1_clean.vba successfully.");

// ---------------------------------------------------------
// 2. UPDATE Excel Excel_Module1_final.vba
// ---------------------------------------------------------
console.log("Reading Excel_Module1_final.vba...");
let vbaExcel = fs.readFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', 'utf8');

// A. Update RibbonThongTinTacGia in Excel to include phone 0984.724.131
const authorExcelText = `1CLICK OFFICE 3.0 ENTERPRISE
Bộ Tiện Ích Chuẩn Hóa Văn Bản & Bảng Tính Hành Chính

- Tác giả: Trần Quốc Được
- Đơn vị: Trung tâm Dịch vụ Tổng hợp Phường Tân Khai
- Điện thoại / Zalo: 0984.724.131
- Trạng thái bản quyền: VĨNH VIỄN (Đã kích hoạt)

Hỗ trợ kỹ thuật & Đóng góp ý kiến:
Vui lòng liên hệ Tác giả Trần Quốc Được - ĐT: 0984.724.131 (UBND / TTDVTH Phường Tân Khai).`;

const newSubAuthorExcel = `Public Sub RibbonThongTinTacGia(control As Object)
    Dim s As String
${toVbaLines(authorExcelText, "s")}
    MsgBoxUni s, vbInformation, ${toChrW("1Click Office - Thông Tin Tác Giả")}
End Sub`;

vbaExcel = vbaExcel.replace(/Public Sub RibbonThongTinTacGia\(control As Object\)[\s\S]*?End Sub/, newSubAuthorExcel);

// B. Update RibbonCapNhat in Excel
const newSubCapNhatExcel = `Public Sub RibbonCapNhat(control As Object)
    Dim s As String
    s = ${toChrW("Phần mềm 1Click Excel - Chuẩn Hóa Trang In & Bảng Biểu (v3.0 Enterprise).")} & vbCrLf & _
        ${toChrW("- Tác giả: Trần Quốc Được (TTDVTH Phường Tân Khai)")} & vbCrLf & _
        ${toChrW("- Điện thoại / Zalo: 0984.724.131")} & vbCrLf & _
        ${toChrW("- Trạng thái: Đã kích hoạt VĨNH VIỄN trên máy tính này.")}
    MsgBoxUni s, vbInformation, ${toChrW("1Click Excel - Thông Tin Tác Giả")}
End Sub`;

vbaExcel = vbaExcel.replace(/Public Sub RibbonCapNhat\(control As Object\)[\s\S]*?End Sub/, newSubCapNhatExcel);

// C. Update print optimization completion popup in Excel (lines 439-444)
const oldExcelChuanHoaRegex = /MsgBox "Da chuan hoa bang bieu & trang in hoan hao!"[\s\S]*?"Chuan Hoa Trang In Excel - 1Click"/;
const newExcelChuanHoaPopup = `MsgBoxUni ${toChrW("ĐÃ CHUẨN HÓA BẢNG BIỂU & TRANG IN HOÀN HẢO!")} & vbCrLf & _
${toChrW("- Khổ giấy tự động: ")} & orientName & vbCrLf & _
${toChrW("- Đã căn chỉnh Fit trang in A4 & chống che khuất chữ 100%.")} & chuNho & vbCrLf & vbCrLf & _
${toChrW("Bấm Ctrl+Z (hoặc nút Hoàn tác) để trả lại như trước khi chuẩn hóa.")} & vbCrLf & vbCrLf & _
${toChrW("1Click Office - Tác giả: Trần Quốc Được (TTDVTH Phường Tân Khai) - ĐT: 0984.724.131")}, vbInformation, ${toChrW("Chuẩn Hóa Trang In Excel - 1Click")}`;

vbaExcel = vbaExcel.replace(oldExcelChuanHoaRegex, newExcelChuanHoaPopup);

// D. Replace other unaccented MsgBox in Excel with MsgBoxUni
vbaExcel = vbaExcel.replace(
  `MsgBox "Khong co lan chuan hoa nao de hoan tac (hoac Excel da dong / mo lai sau lan chuan hoa).", vbInformation, "Hoan tac Chuan hoa"`,
  `MsgBoxUni ${toChrW("Không có lần chuẩn hóa nào để hoàn tác (hoặc Excel đã đóng/mở lại sau lần chuẩn hóa).")}, vbInformation, ${toChrW("Hoàn Tác Chuẩn Hóa")}`
);

vbaExcel = vbaExcel.replace(
  `If Not silentMode Then MsgBox "Sheet dang bi khoa (Protect Sheet). Hay bo khoa sheet truoc khi chuan hoa.", vbExclamation, "Chuan Hoa Excel"`,
  `If Not silentMode Then MsgBoxUni ${toChrW("Bảng tính (Sheet) đang bị khóa bảo vệ. Vui lòng bỏ khóa Sheet trước khi chuẩn hóa.")}, vbExclamation, ${toChrW("Chuẩn Hóa Excel")}`
);

vbaExcel = vbaExcel.replace(
  `MsgBox "Bang tinh hien tai khong co du lieu!", vbExclamation, "Chuan Hoa Excel"`,
  `MsgBoxUni ${toChrW("Bảng tính hiện tại không có dữ liệu!")}, vbExclamation, ${toChrW("Chuẩn Hóa Excel")}`
);

vbaExcel = vbaExcel.replace(
  `MsgBox "Co loi khi chuan hoa: " & errMsg & vbCrLf & "Da tra bang tinh ve nhu truoc khi chuan hoa.", vbCritical, "Chuan Hoa Excel"`,
  `MsgBoxUni ${toChrW("Có lỗi khi chuẩn hóa: ")} & errMsg & vbCrLf & ${toChrW("Đã trả bảng tính về như trước khi chuẩn hóa.")}, vbCritical, ${toChrW("Chuẩn Hóa Excel")}`
);

vbaExcel = vbaExcel.replace(
  `MsgBox "Co loi khi chuan hoa: " & errMsg, vbExclamation, "Chuan Hoa Excel"`,
  `MsgBoxUni ${toChrW("Có lỗi khi chuẩn hóa: ")} & errMsg, vbExclamation, ${toChrW("Chuẩn Hóa Excel")}`
);

vbaExcel = vbaExcel.replace(
  `If hoiKhiKhongCo Then MsgBox "Khong ket noi duoc may chu cap nhat (GitHub). Kiem tra mang roi thu lai." & vbCrLf & "Phien ban dang dung: " & CN_PhienBan(), vbExclamation, "Cap nhat Add-in"`,
  `If hoiKhiKhongCo Then MsgBoxUni ${toChrW("Không kết nối được máy chủ cập nhật (GitHub). Vui lòng kiểm tra mạng Internet rồi thử lại.")} & vbCrLf & ${toChrW("Phiên bản đang dùng: ")} & CN_PhienBan(), vbExclamation, ${toChrW("Cập Nhật Add-in")}`
);

vbaExcel = vbaExcel.replace(
  `If hoiKhiKhongCo Then MsgBox "Ban dang dung phien ban moi nhat: " & CN_PhienBan(), vbInformation, "Cap nhat Add-in"`,
  `If hoiKhiKhongCo Then MsgBoxUni ${toChrW("Bạn đang dùng phiên bản mới nhất: ")} & CN_PhienBan() & " Enterprise.", vbInformation, ${toChrW("Kiểm Tra Cập Nhật")}`
);

fs.writeFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', vbaExcel, 'utf8');
console.log("Updated Excel_Module1_final.vba successfully.");

// ---------------------------------------------------------
// 3. UPDATE HuongDan_SuDung.txt
// ---------------------------------------------------------
console.log("Updating HuongDan_SuDung.txt...");
const huongDanPath = 'G:\\1Click Build\\BoCaiDat_1Click\\HuongDan_SuDung.txt';
let huongDan = fs.readFileSync(huongDanPath, 'utf8');
if (!huongDan.includes('0984.724.131')) {
  huongDan = huongDan.replace(
    /Tác giả: Trần Quốc Được - TTDVTH Phường Tân Khai/g,
    'Tác giả: Trần Quốc Được - ĐT/Zalo: 0984.724.131 - Trung tâm Dịch vụ Tổng hợp Phường Tân Khai'
  );
  if (!huongDan.includes('0984.724.131')) {
    huongDan += '\r\n\r\nTHÔNG TIN TÁC GIẢ & HỖ TRỢ:\r\n- Tác giả: Trần Quốc Được\r\n- Điện thoại / Zalo: 0984.724.131\r\n- Đơn vị: Trung tâm Dịch vụ Tổng hợp Phường Tân Khai\r\n';
  }
  fs.writeFileSync(huongDanPath, huongDan, 'utf8');
  console.log("Updated HuongDan_SuDung.txt successfully.");
}
