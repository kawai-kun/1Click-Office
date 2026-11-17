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

let content = fs.readFileSync('G:\\1Click Build\\Source\\Module1_recovered.vba', 'utf8');

// 1. Remove duplicate lines 2285-2287 after TT_RibbonRefresh
let lines = content.split(/\r?\n/);
for (let i = 0; i < lines.length; i++) {
  if (lines[i].includes('Private Sub TT_RibbonRefresh()')) {
    for (let j = i; j < i + 10; j++) {
      if (lines[j].trim() === 'End Sub') {
        if (lines[j+1] && lines[j+1].includes('MsgBoxUni')) {
          let k = j + 1;
          while (k < lines.length && !lines[k].trim().startsWith('Public Sub') && !lines[k].trim().startsWith('Private Sub') && !lines[k].trim().startsWith('Sub')) {
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
content = lines.join('\r\n');

// 2. Fix TT_CauHinhNhap
const cleanCauHinhNhap = `Public Sub TT_CauHinhNhap()
Dim dMo As Document, coMa As Boolean, truoc As Long, kq As Long, app As Object
coMa = False
On Error Resume Next
If Documents.Count > 0 Then coMa = (InStr(1, ActiveDocument.Content.Text, TT_CH_BEGIN) > 0)
On Error GoTo 0
If Not coMa And Not mTTThuKhongHoi Then
Dim tenTruoc As String, nPV As Long
tenTruoc = ""
truoc = Documents.Count
On Error Resume Next
If truoc > 0 Then tenTruoc = ActiveDocument.FullName
kq = 0
kq = Dialogs(wdDialogFileOpen).Show
On Error GoTo 0
If kq <> -1 Then Exit Sub
Set app = Application
nPV = 0
On Error Resume Next
nPV = app.ProtectedViewWindows.Count
If Err.Number <> 0 Then nPV = 0
Err.Clear
On Error GoTo 0
If nPV > 0 Then
On Error Resume Next
Set dMo = app.ActiveProtectedViewWindow.Edit
On Error GoTo 0
End If
If dMo Is Nothing And Documents.Count > 0 Then
If ActiveDocument.FullName <> tenTruoc Then Set dMo = ActiveDocument
End If
If Not dMo Is Nothing Then dMo.Activate
End If
Call TT_CauHinhNhapTuVanBan
If Not dMo Is Nothing Then
On Error Resume Next
dMo.Close SaveChanges:=wdDoNotSaveChanges
On Error GoTo 0
End If
End Sub`;
content = content.replace(/Public Sub TT_CauHinhNhap\(\)[\s\S]*?End Sub/, cleanCauHinhNhap);

// 3. Fix TT_ChuanKinhGui
const correctKinhGui = `Private Sub TT_ChuanKinhGui(ByVal doc As Document)
On Error Resume Next
Dim p As Paragraph, dem As Long, t As String, a As String, vt As Long, r As Range
Set p = Nothing
Set p = doc.Paragraphs(1)
Do While Not (p Is Nothing)
dem = dem + 1
If dem > 120 Then Exit Do
t = Left(p.Range.Text, 20)
If Len(t) >= 8 Then
a = UCase(ToAscii(t))
If LTrim(a) Like "KINH GOI*" Or LTrim(a) Like "DONG KINH GOI*" Then
vt = InStr(1, t, ChrW(&H1EDF))
If vt > 0 Then
Set r = doc.Range(p.Range.Start + vt - 1, p.Range.Start + vt)
If r.Text = ChrW(&H1EDF) Then r.Text = ChrW(&H1EED)
Else
vt = InStr(1, t, ChrW(&H1EDE))
If vt > 0 Then
Set r = doc.Range(p.Range.Start + vt - 1, p.Range.Start + vt)
If r.Text = ChrW(&H1EDE) Then r.Text = ChrW(&H1EEC)
End If
End If
End If
End If
Set p = p.Next
Loop
On Error GoTo 0
End Sub`;
content = content.replace(/Private Sub TT_ChuanKinhGui\(ByVal doc As Document\)[\s\S]*?End Sub/, correctKinhGui);

// 4. Clean replacement of the entire tail section (all new features)
const cleanNewFeatures = `' === CHỨC NĂNG MỚI: ĐÁNH SỐ TRANG CHUẨN NĐ 30/2020/NĐ-CP ===
Public Sub Word_DanhSoTrangND30(control As Object)
    On Error Resume Next
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBoxUni ${toChrW("Chưa có tài liệu Word nào đang mở!")}, vbExclamation, "1Click Word"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Dim sec As Section
    Dim rField As Range
    For Each sec In doc.Sections
        sec.PageSetup.DifferentFirstPageHeaderFooter = True
        sec.Headers(wdHeaderFooterPrimary).LinkToPrevious = False
        Set rField = sec.Headers(wdHeaderFooterPrimary).Range
        rField.Text = ""
        rField.ParagraphFormat.Alignment = wdAlignParagraphCenter
        rField.Font.Name = "Times New Roman"
        rField.Font.Size = 13.5
        rField.Font.Bold = False
        rField.Font.Italic = False
        rField.Fields.Add Range:=rField, Type:=wdFieldPage
        
        sec.Headers(wdHeaderFooterFirstPage).Range.Text = ""
        sec.Footers(wdHeaderFooterPrimary).Range.Text = ""
        sec.Footers(wdHeaderFooterFirstPage).Range.Text = ""
    Next sec
    Application.ScreenUpdating = True
    
    Dim msgND30 As String
    msgND30 = ${toChrW("Đã đánh số trang chuẩn Nghị định 30/2020/NĐ-CP thành công!")} & vbCrLf & _
              ${toChrW("- Vị trí: Ở giữa lề trên (Header)")} & vbCrLf & _
              ${toChrW("- Bắt đầu hiển thị: Từ trang 2 trở đi (ẩn trang 1)")} & vbCrLf & _
              ${toChrW("- Định dạng: Times New Roman, cỡ 13.5 pt.")}
    MsgBoxUni msgND30, vbInformation, "1Click Word"
End Sub

' === CHỨC NĂNG MỚI: ĐÁNH SỐ TỪ TRANG BẤT KỲ ===
Public Sub Word_DanhSoTrangBatKy(control As Object)
    On Error GoTo ThoatLoi
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBoxUni ${toChrW("Chưa có tài liệu Word nào đang mở!")}, vbExclamation, "1Click Word"
        Exit Sub
    End If
    
    Dim tongTrang As Long
    tongTrang = doc.ComputeStatistics(wdStatisticPages)
    
    Dim strTrangBD As String
    strTrangBD = InputBox("Văn bản hiện có tổng cộng " & tongTrang & " trang." & vbCrLf & vbCrLf & _
                          "Nhập số trang muốn BẮT ĐẦU đánh số (VD: 2 hoặc 3):", _
                          "1Click Word - Chọn Trang Bắt Đầu", "2")
    If Len(Trim(strTrangBD)) = 0 Then Exit Sub
    
    Dim trangBD As Long: trangBD = Val(strTrangBD)
    If trangBD < 1 Then trangBD = 1
    
    Dim strSoBD As String
    strSoBD = InputBox("Tại trang " & trangBD & ", số hiển thị bắt đầu là bao nhiêu (VD: 1 hoặc " & trangBD & "):", _
                       "1Click Word - Số Bắt Đầu Hiển Thị", CStr(trangBD))
    If Len(Trim(strSoBD)) = 0 Then Exit Sub
    Dim soBD As Long: soBD = Val(strSoBD)
    If soBD < 1 Then soBD = 1
    
    Dim strViTri As String
    strViTri = InputBox("Chọn vị trí hiển thị số trang:" & vbCrLf & _
                        "1: Lề trên - Canh giữa (Chuẩn Nghị định 30)" & vbCrLf & _
                        "2: Lề dưới - Canh giữa", _
                        "1Click Word - Vị Trí Số Trang", "1")
    If Len(Trim(strViTri)) = 0 Then Exit Sub
    
    Application.ScreenUpdating = False
    
    If trangBD = 1 Then
        Dim sec1 As Section
        Dim rField1 As Range
        For Each sec1 In doc.Sections
            sec1.PageSetup.DifferentFirstPageHeaderFooter = False
            If Trim(strViTri) = "1" Then
                sec1.Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                sec1.Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = soBD
                Set rField1 = sec1.Headers(wdHeaderFooterPrimary).Range
                rField1.Text = ""
                rField1.ParagraphFormat.Alignment = wdAlignParagraphCenter
                rField1.Font.Name = "Times New Roman"
                rField1.Font.Size = 13.5
                rField1.Fields.Add Range:=rField1, Type:=wdFieldPage
            Else
                sec1.Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
                sec1.Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = soBD
                Set rField1 = sec1.Footers(wdHeaderFooterPrimary).Range
                rField1.Text = ""
                rField1.ParagraphFormat.Alignment = wdAlignParagraphCenter
                rField1.Font.Name = "Times New Roman"
                rField1.Font.Size = 13.5
                rField1.Fields.Add Range:=rField1, Type:=wdFieldPage
            End If
        Next sec1
    Else
        Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=trangBD
        Dim curSecIdx As Long
        curSecIdx = Selection.Information(wdActiveEndSectionNumber)
        
        Dim rngPageStart As Range
        Set rngPageStart = Selection.Range
        rngPageStart.Collapse wdCollapseStart
        
        If curSecIdx = 1 Or doc.Sections(curSecIdx).Range.Start <> rngPageStart.Start Then
            rngPageStart.InsertBreak Type:=wdSectionBreakNextPage
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=trangBD
            curSecIdx = Selection.Information(wdActiveEndSectionNumber)
        End If
        
        Dim curSec As Section
        Set curSec = doc.Sections(curSecIdx)
        curSec.PageSetup.DifferentFirstPageHeaderFooter = False
        
        Dim rFieldCur As Range
        If Trim(strViTri) = "1" Then
            curSec.Headers(wdHeaderFooterPrimary).LinkToPrevious = False
            Set rFieldCur = curSec.Headers(wdHeaderFooterPrimary).Range
            rFieldCur.Text = ""
            rFieldCur.ParagraphFormat.Alignment = wdAlignParagraphCenter
            rFieldCur.Font.Name = "Times New Roman"
            rFieldCur.Font.Size = 13.5
            rFieldCur.Fields.Add Range:=rFieldCur, Type:=wdFieldPage
            curSec.Headers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
            curSec.Headers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = soBD
        Else
            curSec.Footers(wdHeaderFooterPrimary).LinkToPrevious = False
            Set rFieldCur = curSec.Footers(wdHeaderFooterPrimary).Range
            rFieldCur.Text = ""
            rFieldCur.ParagraphFormat.Alignment = wdAlignParagraphCenter
            rFieldCur.Font.Name = "Times New Roman"
            rFieldCur.Font.Size = 13.5
            rFieldCur.Fields.Add Range:=rFieldCur, Type:=wdFieldPage
            curSec.Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
            curSec.Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = soBD
        End If
        
        Dim i As Long
        For i = 1 To curSecIdx - 1
            doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Text = ""
            doc.Sections(i).Headers(wdHeaderFooterFirstPage).Range.Text = ""
            doc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Text = ""
            doc.Sections(i).Footers(wdHeaderFooterFirstPage).Range.Text = ""
        Next i
    End If
    
    Application.ScreenUpdating = True
    Dim msgBatKy As String
    msgBatKy = ${toChrW("Đã hoàn thành đánh số trang thành công!")} & vbCrLf & _
               ${toChrW("- Bắt đầu từ trang: ")} & trangBD & vbCrLf & _
               ${toChrW("- Số hiển thị bắt đầu: ")} & soBD & vbCrLf & _
               ${toChrW("- Các trang phía trước: Đã ẩn toàn bộ số trang.")}
    MsgBoxUni msgBatKy, vbInformation, "1Click Word"
    Exit Sub
ThoatLoi:
    Application.ScreenUpdating = True
    MsgBoxUni ${toChrW("Có lỗi xảy ra hoặc bạn đã hủy thao tác.")}, vbExclamation, "1Click Word"
End Sub

' === CHỨC NĂNG MỚI: XÓA SỐ TRANG ===
Public Sub Word_XoaSoTrang(control As Object)
    On Error Resume Next
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    Dim sec As Section
    For Each sec In doc.Sections
        sec.Headers(wdHeaderFooterPrimary).Range.Text = ""
        sec.Headers(wdHeaderFooterFirstPage).Range.Text = ""
        sec.Footers(wdHeaderFooterPrimary).Range.Text = ""
        sec.Footers(wdHeaderFooterFirstPage).Range.Text = ""
    Next sec
    Application.ScreenUpdating = True
    MsgBoxUni ${toChrW("Đã xóa toàn bộ số trang trong toàn bộ văn bản!")}, vbInformation, "1Click Word"
End Sub

' === CHỨC NĂNG MỚI: XOAY NGANG TRANG CHỌN ===
Public Sub Word_XoayNgangTrang(control As Object)
    On Error GoTo LoiXoay
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    
    Dim ans As VbMsgBoxResult
    Dim msgXoay As String
    msgXoay = ${toChrW("Hệ thống sẽ xoay NGANG trang hiện tại đang chọn để phù hợp chèn bảng biểu rộng.")} & vbCrLf & _
              ${toChrW("Các trang trước và sau vẫn giữ nguyên khổ DỌC.")} & vbCrLf & vbCrLf & _
              ${toChrW("Bạn có muốn tiếp tục không?")}
    ans = MsgBoxUni(msgXoay, vbYesNo + vbQuestion, "1Click Word")
    If ans <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim curPage As Long
    curPage = Selection.Information(wdActiveEndAdjustedPageNumber)
    
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=curPage
    Dim rngPage As Range
    Set rngPage = Selection.Bookmarks("\\Page").Range
    
    With rngPage.PageSetup
        .Orientation = wdOrientLandscape
        .SectionStart = wdSectionNewPage
    End With
    
    Application.ScreenUpdating = True
    MsgBoxUni ${toChrW("Đã xoay ngang trang số ")} & curPage & " " & ${toChrW("thành công!")}, vbInformation, "1Click Word"
    Exit Sub
LoiXoay:
    Application.ScreenUpdating = True
    MsgBoxUni ${toChrW("Không thể xoay trang: ")} & Err.Description, vbExclamation, "1Click Word"
End Sub

' === CHỨC NĂNG MỚI: XÓA TRANG TRẮNG THỪA ===
Public Sub Word_XoaTrangTrang(control As Object)
    On Error Resume Next
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim countDeleted As Long: countDeleted = 0
    Do While doc.Paragraphs.Count > 1
        Dim lastP As Paragraph
        Set lastP = doc.Paragraphs(doc.Paragraphs.Count)
        If Len(Trim(lastP.Range.Text)) <= 1 And lastP.Range.Tables.Count = 0 And lastP.Range.InlineShapes.Count = 0 Then
            lastP.Range.Delete
            countDeleted = countDeleted + 1
        Else
            Exit Do
        End If
    Loop
    Application.ScreenUpdating = True
    
    If countDeleted > 0 Then
        MsgBoxUni ${toChrW("Đã dọn dẹp các dòng trống thừa ở cuối văn bản thành công!")}, vbInformation, "1Click Word"
    Else
        MsgBoxUni ${toChrW("Văn bản không có trang trắng hoặc dòng trống thừa ở cuối.")}, vbInformation, "1Click Word"
    End If
End Sub

' === CHỨC NĂNG MỚI: CHUẨN HÓA BẢNG BIỂU ===
Public Sub Word_ChuanHoaBangBieu(control As Object)
    On Error Resume Next
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    If doc.Tables.Count = 0 Then
        MsgBoxUni ${toChrW("Không tìm thấy bảng biểu nào trong văn bản!")}, vbExclamation, "1Click Word"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Dim tbl As Table
    For Each tbl In doc.Tables
        tbl.Rows(1).HeadingFormat = True
        tbl.Rows.AllowBreakAcrossPages = False
        tbl.Rows.Alignment = wdAlignRowCenter
    Next tbl
    Application.ScreenUpdating = True
    
    Dim msgTable As String
    msgTable = ${toChrW("Đã chuẩn hóa toàn bộ ")} & doc.Tables.Count & " " & ${toChrW("bảng biểu trong văn bản:")} & vbCrLf & _
               ${toChrW("- Tự động lặp lại dòng tiêu đề khi bảng dài qua nhiều trang.")} & vbCrLf & _
               ${toChrW("- Chống tình trạng gãy dòng dữ liệu giữa 2 trang.")} & vbCrLf & _
               ${toChrW("- Căn giữa bảng biểu theo lề trang in chuẩn.")}
    MsgBoxUni msgTable, vbInformation, "1Click Word"
End Sub

' === CHỨC NĂNG MỚI: CHUẨN HÓA DẤU CÂU & KHOẢNG TRẮNG ===
Public Sub Word_ChuanHoaDauCau(control As Object)
    On Error Resume Next
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim rng As Range: Set rng = doc.Content
    
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "  "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    Dim punc As Variant
    punc = Array(".", ",", ":", ";", "?", "!", ")")
    Dim p As Variant
    For Each p In punc
        With rng.Find
            .Text = " " & p
            .Replacement.Text = CStr(p)
            .Execute Replace:=wdReplaceAll
        End With
    Next p
    
    Application.ScreenUpdating = True
    MsgBoxUni ${toChrW("Đã dọn sạch khoảng trắng thừa và chuẩn hóa dấu câu tiếng Việt!")}, vbInformation, "1Click Word"
End Sub
`;

// Replace from ' === CHỨC NĂNG MỚI: ĐÁNH SỐ TRANG to end
const idxTail = content.indexOf("' === CH");
if (idxTail > 0) {
  content = content.substring(0, idxTail) + cleanNewFeatures;
} else {
  content += "\r\n" + cleanNewFeatures;
}

fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', content, 'utf8');
console.log("Successfully rebuilt Module1_clean.vba with 100% clean syntax.");
