const fs = require('fs');

function toChrW(str) {
  if (!str) return '""';
  let chunks = [];
  let curAscii = '';
  for (let ch of str) {
    let code = ch.charCodeAt(0);
    if (code > 127) {
      if (curAscii.length > 0) {
        chunks.push(JSON.stringify(curAscii));
        curAscii = '';
      }
      chunks.push(`ChrW(${code})`);
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
    let piece = toChrW(lines[i]);
    if (i === 0) {
      out.push(`    ${varName} = ${piece}`);
    } else {
      out.push(`    ${varName} = ${varName} & vbCrLf & ${piece}`);
    }
  }
  return out.join('\n');
}

// ==========================================
// 1. TOP DECLARATIONS (MUST BE AT MODULE TOP)
// ==========================================
const topDeclares = `
#If VBA7 Then
    Private Declare PtrSafe Function MsgBoxW Lib "user32" Alias "MessageBoxW" ( _
        ByVal hWnd As LongPtr, _
        ByVal lpText As LongPtr, _
        ByVal lpCaption As LongPtr, _
        ByVal uType As Long) As Long
#Else
    Private Declare Function MsgBoxW Lib "user32" Alias "MessageBoxW" ( _
        ByVal hWnd As Long, _
        ByVal lpText As Long, _
        ByVal lpCaption As Long, _
        ByVal uType As Long) As Long
#End If
`;

// ==========================================
// 2. UNICODE CASES & HELPERS
// ==========================================
const lowerChars = 'aàáảãạăằắẳẵặâầấẩẫậeèéẻẽẹêềếểễệiìíỉĩịoòóỏõọôồốổỗộơờớởỡợuùúủũụưừứửữựyỳýỷỹỵđ';
const upperChars = 'AÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬEÈÉẺẼẸÊỀẾỂỄỆIÌÍỈĨỊOÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢUÙÚỦŨỤƯỪỨỬỮỰYỲÝỶỸỴĐ';
const lArr = [...lowerChars];
const uArr = [...upperChars];

let uCaseCode = `Public Function UCaseUnicode(ByVal s As String) As String
    Dim i As Long, ch As String, cCode As Long, res As String
    res = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        cCode = AscW(ch)
        If cCode < 0 Then cCode = cCode + 65536
        Select Case cCode
`;
for (let i = 0; i < lArr.length; i++) {
  uCaseCode += `            Case ${lArr[i].charCodeAt(0)}: res = res & ChrW(${uArr[i].charCodeAt(0)})\n`;
}
uCaseCode += `            Case Else
                res = res & UCase(ch)
        End Select
    Next i
    UCaseUnicode = res
End Function
`;

let lCaseCode = `Public Function LCaseUnicode(ByVal s As String) As String
    Dim i As Long, ch As String, cCode As Long, res As String
    res = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        cCode = AscW(ch)
        If cCode < 0 Then cCode = cCode + 65536
        Select Case cCode
`;
for (let i = 0; i < uArr.length; i++) {
  lCaseCode += `            Case ${uArr[i].charCodeAt(0)}: res = res & ChrW(${lArr[i].charCodeAt(0)})\n`;
}
lCaseCode += `            Case Else
                res = res & LCase(ch)
        End Select
    Next i
    LCaseUnicode = res
End Function
`;

const chuanHoaTenRieng = `Public Function ChuanHoa_TenRieng(ByVal s As String) As String
    s = Application.Trim(s)
    If Len(s) = 0 Then
        ChuanHoa_TenRieng = ""
        Exit Function
    End If
    s = LCaseUnicode(s)
    Dim words() As String
    words = Split(s, " ")
    Dim i As Long
    For i = LBound(words) To UBound(words)
        If Len(words(i)) > 0 Then
            words(i) = UCaseUnicode(Left(words(i), 1)) & Mid(words(i), 2)
        End If
    Next i
    ChuanHoa_TenRieng = Join(words, " ")
End Function
`;

const docSoCode = `Public Function DocSo(ByVal nVal As Variant, Optional ByVal DongY As Boolean = True) As String
    On Error Resume Next
    If IsNull(nVal) Or IsEmpty(nVal) Then
        DocSo = ""
        Exit Function
    End If
    Dim strVal As String
    strVal = Trim(CStr(nVal))
    strVal = Replace(strVal, ",", "")
    strVal = Replace(strVal, " ", "")
    If Not IsNumeric(strVal) Then
        DocSo = ""
        Exit Function
    End If
    
    Dim dVal As Double
    dVal = CDbl(strVal)
    If dVal = 0 Then
        If DongY Then
            DocSo = ${toChrW('Không đồng chẵn.')}
        Else
            DocSo = ${toChrW('Không.')}
        End If
        Exit Function
    End If
    
    Dim am As Boolean: am = False
    If dVal < 0 Then
        am = True
        dVal = Abs(dVal)
    End If
    
    Dim intPart As Double
    intPart = Fix(dVal)
    
    Dim sNum As String
    sNum = Format(intPart, "0")
    
    Dim lop As Variant
    lop = Array("", ${toChrW('nghìn')}, ${toChrW('triệu')}, ${toChrW('tỷ')}, ${toChrW('nghìn tỷ')}, ${toChrW('triệu tỷ')})
    
    Dim res As String: res = ""
    Dim countLop As Long: countLop = 0
    
    Do While Len(sNum) > 0
        Dim block As Long, strBlock As String
        If Len(sNum) > 3 Then
            strBlock = Right(sNum, 3)
            sNum = Left(sNum, Len(sNum) - 3)
        Else
            strBlock = sNum
            sNum = ""
        End If
        block = CLng(strBlock)
        
        If block > 0 Or (Len(sNum) > 0 And countLop = 0) Then
            Dim docBlock As String
            docBlock = Doc3So(block, (Len(sNum) > 0))
            If Len(docBlock) > 0 Then
                Dim tenLop As String
                If countLop <= 5 Then
                    tenLop = " " & lop(countLop)
                Else
                    tenLop = ""
                End If
                If countLop = 0 Then tenLop = ""
                res = docBlock & tenLop & " " & res
            End If
        End If
        countLop = countLop + 1
    Loop
    
    res = Application.Trim(res)
    If Len(res) > 0 Then
        res = UCaseUnicode(Left(res, 1)) & Mid(res, 2)
        If DongY Then
            res = res & " " & ${toChrW('đồng chẵn.')}
        Else
            res = res & "."
        End If
    End If
    If am Then res = ${toChrW('Âm ')} & LCaseUnicode(res)
    DocSo = res
End Function

Private Function Doc3So(ByVal n As Long, ByVal coPhiaTruoc As Boolean) As String
    If n = 0 Then
        Doc3So = ""
        Exit Function
    End If
    Dim chuSo As Variant
    chuSo = Array(${toChrW('không')}, ${toChrW('một')}, ${toChrW('hai')}, ${toChrW('ba')}, ${toChrW('bốn')}, ${toChrW('năm')}, ${toChrW('sáu')}, ${toChrW('bảy')}, ${toChrW('tám')}, ${toChrW('chín')})
    Dim h As Long, c As Long, d As Long
    h = n \\ 100
    c = (n Mod 100) \\ 10
    d = n Mod 10
    Dim s As String: s = ""
    
    If h > 0 Or coPhiaTruoc Then
        s = s & chuSo(h) & " " & ${toChrW('trăm')} & " "
    End If
    
    If c = 0 Then
        If h > 0 And d > 0 Then s = s & ${toChrW('lẻ')} & " "
    ElseIf c = 1 Then
        s = s & ${toChrW('mười')} & " "
    Else
        s = s & chuSo(c) & " " & ${toChrW('mươi')} & " "
    End If
    
    If d = 1 Then
        If c > 1 Then s = s & ${toChrW('mốt')} Else s = s & chuSo(1)
    ElseIf d = 5 Then
        If c > 0 Then s = s & ${toChrW('lăm')} Else s = s & chuSo(5)
    ElseIf d > 0 Then
        s = s & chuSo(d)
    End If
    Doc3So = Trim(s)
End Function
`;

const unicodeMsgBoxHelper = `
Public Function MsgBoxUni(ByVal prompt As String, _
                         Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                         Optional ByVal title As String = "1Click Office") As VbMsgBoxResult
    On Error Resume Next
    MsgBoxUni = MsgBoxW(0, StrPtr(prompt), StrPtr(title), buttons)
End Function
`;

const romanNumeralHelper = `
Public Function ToRoman(ByVal n As Long) As String
    If n <= 0 Then
        ToRoman = CStr(n)
        Exit Function
    End If
    Dim vals As Variant, syms As Variant
    vals = Array(1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1)
    syms = Array("M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I")
    Dim res As String: res = ""
    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        Do While n >= vals(i)
            res = res & syms(i)
            n = n - vals(i)
        Loop
    Next i
    ToRoman = res
End Function
`;

const authorInfoText = `1CLICK OFFICE 3.0 ENTERPRISE
Bộ Tiện Ích Chuẩn Hóa Văn Bản & Bảng Tính Hành Chính

- Tác giả: Trần Quốc Được
- Đơn vị: Trung tâm Dịch vụ Tổng hợp Phường Tân Khai
- Trạng thái bản quyền: VĨNH VIỄN (Đã kích hoạt)

Hỗ trợ kỹ thuật & Đóng góp ý kiến:
Vui lòng liên hệ UBND Phường Tân Khai / TTDVTH Phường Tân Khai.`;

const helpText = `HƯỚNG DẪN NHANH 1CLICK EXCEL 3.0:

1. Chuẩn Hóa Trang In: Vừa trang A4, lặp tiêu đề, chống mất chữ.
2. Viết Hoa Đầu Từ: Chuẩn hóa họ tên, địa danh (VD: Nguyễn Văn A).
3. Tách Cột Họ & Tên: Tách riêng [Họ và đệm] và [Tên] để sắp xếp A-Z.
4. Đọc Số Ra Chữ: Nút bấm trên ribbon hoặc công thức =DocSo(ô).
5. Định dạng tiền: Phân cách hàng nghìn (1,000,000) tức thì.
6. Đánh lại STT: Đánh số 1, 2, 3... hoặc STT La Mã I, II, III... cho dự toán.
7. Kẻ viền chuẩn & Xóa dòng trống: Tối ưu bảng biểu trong 1 cú nhấp.`;

const uiHandlersV3 = `
Public Sub RibbonThongTinTacGia(control As Object)
    Dim s As String
${toVbaLines(authorInfoText, "s")}
    MsgBoxUni s, vbInformation, ${toChrW('1Click Office - Thông Tin Tác Giả')}
End Sub

Public Sub RibbonHuongDanSuDung(control As Object)
    On Error Resume Next
    Dim p As String
    p = Environ("APPDATA") & "\\1Click\\HuongDan_SuDung.txt"
    If Dir(p) <> "" Then
        CreateObject("WScript.Shell").Run "notepad.exe """ & p & """", 1, False
    Else
        Dim s As String
${toVbaLines(helpText, "s")}
        MsgBoxUni s, vbInformation, ${toChrW('1Click Excel - Hướng Dẫn')}
    End If
End Sub

Public Sub RibbonKiemTraCapNhat(control As Object)
    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim url As String
    url = "https://raw.githubusercontent.com/duoc-tankhai/1Click-Office/main/version.json"
    
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "1ClickOffice"
    http.setTimeouts 3000, 3000, 5000, 5000
    http.Send
    
    If http.Status = 200 Then
        Dim resp As String: resp = http.responseText
        Dim verRemote As String
        verRemote = ParseJsonField(resp, "version")
        If Len(verRemote) > 0 Then
            If verRemote > "3.0.0" Then
                Dim notes As String
                notes = ParseJsonField(resp, "notes")
                Dim ans As VbMsgBoxResult
                ans = MsgBoxUni(${toChrW('Phát hiện có phiên bản mới: ')} & verRemote & vbCrLf & vbCrLf & _
                                notes & vbCrLf & vbCrLf & _
                                ${toChrW('Bạn có muốn tải và cài đặt bản cập nhật mới ngay bây giờ không?')}, vbYesNo + vbQuestion, ${toChrW('Cập Nhật 1Click Office')})
                If ans = vbYes Then
                    Dim setupUrl As String
                    setupUrl = ParseJsonField(resp, "setup_url")
                    If Len(setupUrl) > 0 Then
                        CreateObject("WScript.Shell").Run setupUrl
                    End If
                End If
                Exit Sub
            End If
        End If
        MsgBoxUni ${toChrW('Bạn đang sử dụng phiên bản mới nhất: 3.0 Enterprise.\nHiện tại không có bản cập nhật nào mới hơn.')}, vbInformation, ${toChrW('Kiểm Tra Cập Nhật')}
    Else
        MsgBoxUni ${toChrW('Bạn đang sử dụng phiên bản: 3.0 Enterprise (Mới nhất).\nKhông thể kết nối đến máy chủ GitHub hoặc bạn đang làm việc Ngoại Tuyến (Offline).')}, vbInformation, ${toChrW('Kiểm Tra Cập Nhật')}
    End If
End Sub

Public Sub Xl_CanGiuaTrang(control As Object)
    On Error Resume Next
    ActiveSheet.PageSetup.CenterHorizontally = True
    MsgBoxUni ${toChrW('Đã căn giữa trang in theo chiều ngang thành công!')}, vbInformation, "1Click Excel"
End Sub

Public Sub Xl_DongBangTieuDe(control As Object)
    On Error Resume Next
    Dim curActive As Window: Set curActive = ActiveWindow
    If curActive.FreezePanes Then
        curActive.FreezePanes = False
        MsgBoxUni ${toChrW('Đã bỏ cố định dòng tiêu đề.')}, vbInformation, "1Click Excel"
    Else
        curActive.FreezePanes = True
        MsgBoxUni ${toChrW('Đã cố định dòng tiêu đề bảng tính thành công!')}, vbInformation, "1Click Excel"
    End If
End Sub

Public Sub Xl_XemTruocBanIn(control As Object)
    On Error Resume Next
    ActiveSheet.PrintPreview
End Sub

Public Sub Xl_XoaDinhDangRac(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.ClearFormats
    MsgBoxUni ${toChrW('Đã xóa sạch các định dạng rác trong vùng chọn!')}, vbInformation, "1Click Excel"
End Sub

Public Sub Xl_DocSoRaChu(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection
        If IsNumeric(cell.Value2) And Len(CStr(cell.Value2)) > 0 Then
            Dim strChu As String
            strChu = DocSo(cell.Value2, True)
            cell.Offset(0, 1).Value = strChu
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_DinhDangTienVND(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection
        If IsNumeric(cell.Value2) And Len(CStr(cell.Value2)) > 0 Then
            cell.NumberFormat = "#,##0"
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_TuTinhThanhTien(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rng As Range: Set rng = Selection
    Application.ScreenUpdating = False
    Dim cell As Range
    For Each cell In rng
        If cell.Row > 1 And cell.Column >= 3 Then
            cell.FormulaR1C1 = "=RC[-2]*RC[-1]"
            cell.NumberFormat = "#,##0"
        End If
    Next cell
    Application.ScreenUpdating = True
    MsgBoxUni ${toChrW('Đã tự động tính thành tiền cho các dòng được chọn!')}, vbInformation, "1Click Excel"
End Sub

Public Sub Xl_VietHoaDauTu(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection
        If Not cell.HasFormula And Len(Trim(CStr(cell.Value2))) > 0 Then
            cell.Value = ChuanHoa_TenRieng(CStr(cell.Value2))
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_VietHoaToanBo(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection
        If Not cell.HasFormula And Len(Trim(CStr(cell.Value2))) > 0 Then
            cell.Value = UCaseUnicode(CStr(cell.Value2))
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_ChuThuong(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection
        If Not cell.HasFormula And Len(Trim(CStr(cell.Value2))) > 0 Then
            cell.Value = LCaseUnicode(CStr(cell.Value2))
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_XoaKhoangTrang(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection
        If Not cell.HasFormula And Len(CStr(cell.Value2)) > 0 Then
            cell.Value = Application.Trim(CStr(cell.Value2))
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_TachHoVaTen(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rng As Range: Set rng = Selection
    If rng.Columns.Count > 1 Then
        MsgBoxUni ${toChrW('Vui lòng chỉ chọn 1 cột chứa Họ và Tên cần tách!')}, vbExclamation, "1Click Excel"
        Exit Sub
    End If
    
    Dim ans As VbMsgBoxResult
    ans = MsgBoxUni(${toChrW('Hệ thống sẽ chèn thêm 2 cột mới bên cạnh để chứa [Họ và đệm] và [Tên].\nBạn có muốn tiếp tục không?')}, vbYesNo + vbQuestion, "1Click Excel")
    If ans <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim colIdx As Long: colIdx = rng.Column
    rng.Worksheet.Columns(colIdx + 1).Insert Shift:=xlToRight
    rng.Worksheet.Columns(colIdx + 2).Insert Shift:=xlToRight
    rng.Worksheet.Cells(rng.Row, colIdx + 1).Value = ${toChrW('Họ và đệm')}
    rng.Worksheet.Cells(rng.Row, colIdx + 2).Value = ${toChrW('Tên')}
    
    Dim cell As Range
    For Each cell In rng
        Dim val As String
        val = Application.Trim(CStr(cell.Value2))
        If Len(val) > 0 Then
            val = ChuanHoa_TenRieng(val)
            Dim lastSpace As Long
            lastSpace = InStrRev(val, " ")
            If lastSpace > 0 Then
                cell.Offset(0, 1).Value = Left(val, lastSpace - 1)
                cell.Offset(0, 2).Value = Mid(val, lastSpace + 1)
            Else
                cell.Offset(0, 1).Value = ""
                cell.Offset(0, 2).Value = val
            End If
        End If
    Next cell
    Application.ScreenUpdating = True
    MsgBoxUni ${toChrW('Đã tách Họ và Tên thành công!')}, vbInformation, "1Click Excel"
End Sub

Public Sub Xl_DanhLaiSTT(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range, i As Long: i = 1
    Application.ScreenUpdating = False
    For Each cell In Selection
        cell.Value = i
        i = i + 1
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_DanhSTT_LaMa(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range, i As Long: i = 1
    Application.ScreenUpdating = False
    For Each cell In Selection
        cell.Value = ToRoman(i)
        i = i + 1
    Next cell
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_KeKhungChuan(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rng As Range
    Set rng = Selection
    Application.ScreenUpdating = False
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    rng.Borders(xlInsideHorizontal).Weight = xlThin
    rng.Borders(xlInsideVertical).LineStyle = xlContinuous
    rng.Borders(xlInsideVertical).Weight = xlThin
    rng.Borders(xlEdgeTop).LineStyle = xlContinuous
    rng.Borders(xlEdgeTop).Weight = xlThin
    rng.Borders(xlEdgeBottom).LineStyle = xlContinuous
    rng.Borders(xlEdgeBottom).Weight = xlThin
    rng.Borders(xlEdgeLeft).LineStyle = xlContinuous
    rng.Borders(xlEdgeLeft).Weight = xlThin
    rng.Borders(xlEdgeRight).LineStyle = xlContinuous
    rng.Borders(xlEdgeRight).Weight = xlThin
    Application.ScreenUpdating = True
End Sub

Public Sub Xl_XoaDongTrong(control As Object)
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim r As Long, rng As Range
    Set rng = Selection
    Application.ScreenUpdating = False
    For r = rng.Rows.Count To 1 Step -1
        If Application.CountA(rng.Rows(r)) = 0 Then
            rng.Rows(r).Delete Shift:=xlUp
        End If
    Next r
    Application.ScreenUpdating = True
    MsgBoxUni ${toChrW('Đã xóa các dòng hoàn toàn trống trong vùng chọn!')}, vbInformation, "1Click Excel"
End Sub
`;

// Read original base Excel VBA
const origVbaPath = 'G:\\1Click Build\\Source\\Excel_Module1_clean.vba';
let origContent = fs.readFileSync(origVbaPath, 'utf8');
let cutIdx = origContent.indexOf("' === EXTENDED FEATURES FOR 1CLICK EXCEL");
if (cutIdx === -1) cutIdx = origContent.indexOf("Public Function DocSo(");
let baseVba = (cutIdx !== -1) ? origContent.substring(0, cutIdx).trim() : origContent.trim();

// Insert topDeclares right before the first procedure in baseVba
const firstSubMarker = "Public Sub RibbonChuanHoaExcel";
const markerPos = baseVba.indexOf(firstSubMarker);
let baseWithTopDeclares;
if (markerPos !== -1) {
  baseWithTopDeclares = baseVba.substring(0, markerPos) + topDeclares + '\n' + baseVba.substring(markerPos);
} else {
  baseWithTopDeclares = topDeclares + '\n' + baseVba;
}

let fullNewVba = baseWithTopDeclares + '\n\n' + 
  "' === EXTENDED FEATURES FOR 1CLICK EXCEL 3.0 ENTERPRISE ===\n" +
  "' Author: Tran Quoc Duoc - TTDVTH Phuong Tan Khai\n" +
  "' Zero Telemetry - 100% Offline\n\n" +
  unicodeMsgBoxHelper + '\n' +
  romanNumeralHelper + '\n' +
  uCaseCode + '\n' +
  lCaseCode + '\n' +
  chuanHoaTenRieng + '\n' +
  docSoCode + '\n' +
  uiHandlersV3;

fs.writeFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', fullNewVba, 'utf8');
console.log('Saved G:\\1Click Build\\Source\\Excel_Module1_final.vba. Length:', fullNewVba.length);
