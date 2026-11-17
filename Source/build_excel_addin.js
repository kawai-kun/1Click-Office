const fs = require('fs');
const path = require('path');

// Helper to convert UTF-8 string to VBA ChrW concatenation
function toChrW(str) {
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
  return chunks.join(' & ');
}

// Generate Unicode Case Conversion Cases
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
  let lc = lArr[i].charCodeAt(0);
  let uc = uArr[i].charCodeAt(0);
  uCaseCode += `            Case ${lc}: res = res & ChrW(${uc})\n`;
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
  let lc = lArr[i].charCodeAt(0);
  let uc = uArr[i].charCodeAt(0);
  lCaseCode += `            Case ${uc}: res = res & ChrW(${lc})\n`;
}

lCaseCode += `            Case Else
                res = res & LCase(ch)
        End Select
    Next i
    LCaseUnicode = res
End Function
`;

// Vietnamese Capitalize Names Function
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

// DocSo (Number to Words)
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

// UI Event Handlers
const uiHandlers = `
Public Sub Xl_CanGiuaTrang(control As Object)
    On Error Resume Next
    ActiveSheet.PageSetup.CenterHorizontally = True
    MsgBox ${toChrW('Đã căn giữa trang in theo chiều ngang thành công!')}, vbInformation, "1Click Excel"
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
        MsgBox ${toChrW('Vui lòng chỉ chọn 1 cột chứa Họ và Tên cần tách!')}, vbExclamation, "1Click Excel"
        Exit Sub
    End If
    
    Dim ans As VbMsgBoxResult
    ans = MsgBox(${toChrW('Hệ thống sẽ chèn thêm 2 cột mới bên cạnh để chứa [Họ và đệm] và [Tên].\nBạn có muốn tiếp tục không?')}, vbYesNo + vbQuestion, "1Click Excel")
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
    MsgBox ${toChrW('Đã tách Họ và Tên thành công!')}, vbInformation, "1Click Excel"
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
    MsgBox ${toChrW('Đã xóa các dòng hoàn toàn trống trong vùng chọn!')}, vbInformation, "1Click Excel"
End Sub
`;

// Read original decompressed VBA
const origVbaPath = 'G:\\1Click Build\\Source\\Excel_Module1_clean.vba';
let origContent = fs.readFileSync(origVbaPath, 'utf8');

// Find cut-off point before previous appended functions
let cutIdx = origContent.indexOf("' === EXTENDED FEATURES FOR 1CLICK EXCEL ===");
if (cutIdx === -1) {
  cutIdx = origContent.indexOf("Public Function DocSo(");
}
let baseVba = (cutIdx !== -1) ? origContent.substring(0, cutIdx).trim() : origContent.trim();

let fullNewVba = baseVba + '\n\n' + 
  "' === EXTENDED FEATURES FOR 1CLICK EXCEL 3.0 ENTERPRISE ===\n" +
  "' Author: Tran Quoc Duoc - TTDVTH Phuong Tan Khai\n" +
  "' Zero Telemetry - 100% Offline\n\n" +
  uCaseCode + '\n' +
  lCaseCode + '\n' +
  chuanHoaTenRieng + '\n' +
  docSoCode + '\n' +
  uiHandlers;

fs.writeFileSync('G:\\1Click Build\\Source\\Excel_Module1_final.vba', fullNewVba, 'utf8');
console.log('Successfully generated Excel_Module1_final.vba. Length:', fullNewVba.length);
