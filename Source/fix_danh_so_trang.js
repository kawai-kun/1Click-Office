const fs = require('fs');

let content = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8');

const newDanhSoTrangND30 = `Public Sub Word_DanhSoTrangND30(control As Object)
    On Error Resume Next
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBoxUni "Ch" & ChrW(432) & "a c" & ChrW(243) & " t" & ChrW(224) & "i li" & ChrW(7879) & "u Word n" & ChrW(224) & "o " & ChrW(273) & "ang m" & ChrW(7903) & "!", vbExclamation, "1Click Word"
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
    
    MsgBoxUni ChrW(272) & ChrW(227) & " " & ChrW(273) & ChrW(225) & "nh s" & ChrW(7889) & " trang chu" & ChrW(7849) & "n Ngh" & ChrW(7883) & " " & ChrW(273) & ChrW(7883) & "nh 30/2020/N" & ChrW(272) & "-CP th" & ChrW(224) & "nh c" & ChrW(244) & "ng!\n- V" & ChrW(7883) & " tr" & ChrW(237) & ": " & ChrW(7902) & " gi" & ChrW(7919) & "a l" & ChrW(7873) & " tr" & ChrW(234) & "n (Header)\n- B" & ChrW(7855) & "t " & ChrW(273) & ChrW(7847) & "u hi" & ChrW(7875) & "n th" & ChrW(7883) & ": T" & ChrW(7915) & " trang 2 tr" & ChrW(7903) & " " & ChrW(273) & "i (" & ChrW(7848) & "n trang 1)\n- " & ChrW(272) & ChrW(7883) & "nh d" & ChrW(7841) & "ng: Times New Roman, c" & ChrW(7905) & " 13.5 pt.", vbInformation, "1Click Word"
End Sub`;

const newDanhSoTrangBatKy = `Public Sub Word_DanhSoTrangBatKy(control As Object)
    On Error GoTo ThoatLoi
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBoxUni "Ch" & ChrW(432) & "a c" & ChrW(243) & " t" & ChrW(224) & "i li" & ChrW(7879) & "u Word n" & ChrW(224) & "o " & ChrW(273) & "ang m" & ChrW(7903) & "!", vbExclamation, "1Click Word"
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
    MsgBoxUni ChrW(272) & ChrW(227) & " ho" & ChrW(224) & "n th" & ChrW(224) & "nh " & ChrW(273) & ChrW(225) & "nh s" & ChrW(7889) & " trang th" & ChrW(224) & "nh c" & ChrW(244) & "ng!\n- B" & ChrW(7855) & "t " & ChrW(273) & ChrW(7847) & "u t" & ChrW(7915) & " trang: " & trangBD & vbCrLf & _
              "- S" & ChrW(7889) & " hi" & ChrW(7875) & "n th" & ChrW(7883) & " b" & ChrW(7855) & "t " & ChrW(273) & ChrW(7847) & "u: " & soBD & vbCrLf & _
              "- C" & ChrW(225) & "c trang ph" & ChrW(237) & "a tr" & ChrW(432) & ChrW(7899) & "c: " & ChrW(272) & ChrW(227) & " " & ChrW(7849) & "n to" & ChrW(224) & "n b" & ChrW(7897) & " s" & ChrW(7889) & " trang.", vbInformation, "1Click Word"
    Exit Sub
ThoatLoi:
    Application.ScreenUpdating = True
    MsgBoxUni "C" & ChrW(243) & " l" & ChrW(7895) & "i x" & ChrW(7843) & "y ra ho" & ChrW(7863) & "c b" & ChrW(7841) & "n " & ChrW(273) & ChrW(227) & " h" & ChrW(7911) & "y thao t" & ChrW(225) & "c.", vbExclamation, "1Click Word"
End Sub`;

content = content.replace(/Public Sub Word_DanhSoTrangND30\(control As Object\)[\s\S]*?End Sub/, newDanhSoTrangND30);
content = content.replace(/Public Sub Word_DanhSoTrangBatKy\(control As Object\)[\s\S]*?End Sub/, newDanhSoTrangBatKy);

fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', content, 'utf8');
console.log("Updated Word_DanhSoTrangND30 and BatKy successfully.");
