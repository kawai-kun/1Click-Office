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
// 1. UPDATE WORD customUI.xml
// ==========================================
let xml = fs.readFileSync('G:\\1Click Build\\Source\\customUI.xml', 'utf8');

// Remove any existing grpTacGiaWord and grpTienIchWord so we rebuild them cleanly
xml = xml.replace(/<group id="grpTienIchWord"[\s\S]*?<\/group>\s*/g, '');
xml = xml.replace(/<group id="grpTacGiaWord"[\s\S]*?<\/group>\s*/g, '');

const newWordGroups = `
        <group id="grpTienIchWord" label="Trang &amp; Bảng Biểu">
          <splitButton id="sbDanhSoTrang" size="large">
            <button id="btnDanhSoND30"
                    label="Đánh Số Trang"
                    imageMso="PageNumberInsert"
                    screentip="Đánh số trang chuẩn Nghị định 30/2020/NĐ-CP"
                    supertip="Tự động đánh số trang ở giữa lề trên từ trang thứ 2, cỡ chữ 13.5 Times New Roman, ẩn số trang ở trang 1."
                    onAction="Word_DanhSoTrangND30" />
            <menu id="mnuDanhSoTrang">
              <button id="btnSoTrangND30Menu"
                      label="Đánh số chuẩn NĐ 30 (Từ trang 2)"
                      imageMso="PageNumberInsert"
                      screentip="Đánh số trang lề trên từ trang 2 theo NĐ 30/2020/NĐ-CP"
                      onAction="Word_DanhSoTrangND30" />
              <button id="btnSoTrangBatKyMenu"
                      label="Đánh số từ trang bất kỳ..."
                      imageMso="PageNumberFormatDialog"
                      screentip="Tùy chọn trang bắt đầu và số thứ tự bắt đầu"
                      supertip="Cho phép bạn chọn trang bắt đầu đánh số (VD: trang 3, 4...) và chọn số bắt đầu hiển thị (VD: 1 hoặc 2)."
                      onAction="Word_DanhSoTrangBatKy" />
              <button id="btnXoaSoTrangMenu"
                      label="Xóa toàn bộ số trang"
                      imageMso="PageNumberRemove"
                      screentip="Xóa sạch toàn bộ số trang trong văn bản"
                      onAction="Word_XoaSoTrang" />
            </menu>
          </splitButton>
          <button id="btnXoayNgangTrang"
                  label="Xoay ngang trang"
                  imageMso="PageOrientationLandscape"
                  screentip="Xoay ngang trang đang chọn để chứa bảng biểu rộng"
                  supertip="Chỉ xoay ngang trang hiện tại, các trang trước và sau vẫn giữ nguyên khổ dọc bình thường."
                  onAction="Word_XoayNgangTrang" />
          <button id="btnChuanHoaBang"
                  label="Chuẩn hóa bảng biểu"
                  imageMso="TableProperties"
                  screentip="Lặp tiêu đề bảng qua các trang &amp; căn giữa"
                  supertip="Tự động lặp lại dòng tiêu đề khi bảng dài qua nhiều trang, chống gãy dòng bảng và căn giữa trang."
                  onAction="Word_ChuanHoaBangBieu" />
          <button id="btnXoaTrangTrang"
                  label="Xóa trang trắng thừa"
                  imageMso="Delete"
                  screentip="Xóa các trang trắng rác và đoạn trống ở cuối văn bản"
                  onAction="Word_XoaTrangTrang" />
          <button id="btnChuanHoaDauCau"
                  label="Chuẩn hóa dấu câu"
                  imageMso="ProofingTools"
                  screentip="Xóa khoảng trắng thừa và chuẩn hóa dấu câu tiếng Việt"
                  onAction="Word_ChuanHoaDauCau" />
        </group>
        <group id="grpTacGiaWord" label="Tác Giả &amp; Hỗ Trợ">
          <button id="btnThongTinTacGiaWord"
                  label="Thông Tin Tác Giả"
                  size="large"
                  imageMso="AddressBook"
                  screentip="Thông tin tác giả &amp; Bản quyền"
                  supertip="1Click Office 3.0 Enterprise&#10;Tác giả: Trần Quốc Được&#10;Đơn vị: Trung tâm Dịch vụ Tổng hợp Phường Tân Khai&#10;Bản quyền: Vĩnh viễn (Đã kích hoạt)."
                  onAction="RibbonThongTinTacGia" />
          <button id="btnKiemTraCapNhatWord"
                  label="Kiểm tra cập nhật"
                  imageMso="SyncStatusNormal"
                  screentip="Kiểm tra phiên bản mới từ máy chủ GitHub"
                  supertip="Tự động kiểm tra phiên bản mới từ GitHub và cập nhật nếu có."
                  onAction="RibbonKiemTraCapNhat" />
          <button id="btnHuongDanWord"
                  label="Hướng dẫn sử dụng"
                  imageMso="Help"
                  screentip="Xem tài liệu hướng dẫn sử dụng 1Click Office"
                  onAction="RibbonHuongDanSuDung" />
        </group>
      </tab>
    </tabs>`;

xml = xml.replace(/<\/tab>\s*<\/tabs>/, newWordGroups);
fs.writeFileSync('G:\\1Click Build\\Source\\customUI.xml', xml, 'utf8');
console.log('Saved updated G:\\1Click Build\\Source\\customUI.xml with grpTienIchWord and AddressBook icon');

// ==========================================
// 2. UPDATE WORD VBA (Module1_clean.vba)
// ==========================================
const authorInfoText = `1CLICK OFFICE 3.0 ENTERPRISE
Bộ Tiện Ích Chuẩn Hóa Văn Bản & Bảng Tính Hành Chính

- Tác giả: Trần Quốc Được
- Đơn vị: Trung tâm Dịch vụ Tổng hợp Phường Tân Khai
- Trạng thái bản quyền: VĨNH VIỄN (Đã kích hoạt)

Hỗ trợ kỹ thuật & Đóng góp ý kiến:
Vui lòng liên hệ UBND Phường Tân Khai / TTDVTH Phường Tân Khai.`;

const helpTextWord = `HƯỚNG DẪN NHANH 1CLICK WORD 3.0:

1. Chuẩn Hóa Thể Thức: Căn chỉnh tự động văn bản chuẩn Nghị định 30/2020/NĐ-CP và Đảng, Đoàn thể.
2. Tạo Văn Bản Mẫu: Tờ trình, Kế hoạch, Quyết định, Báo cáo, Giấy mời, Biên bản...
3. Đánh Số Trang NĐ 30: Bắt đầu từ trang 2, lề trên giữa, ẩn trang 1.
4. Đánh Số Trang Bất Kỳ: Cho phép chọn trang bắt đầu và số thứ tự hiển thị.
5. Xoay Ngang Trang: Xoay ngang riêng trang bảng biểu, các trang khác giữ dọc.
6. Chuẩn Hóa Bảng Biểu: Tự động lặp dòng tiêu đề qua các trang, căn giữa, chống gãy dòng.
7. Xóa Trang Trắng & Dấu Câu: Xóa sạch trang trắng thừa, chuẩn hóa khoảng cách.`;

const wordAdvancedSubs = `
Public Function MsgBoxUni(ByVal prompt As String, _
                         Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                         Optional ByVal title As String = "1Click Office") As VbMsgBoxResult
    On Error Resume Next
    MsgBoxUni = MsgBoxW(0, StrPtr(prompt), StrPtr(title), buttons)
End Function

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
${toVbaLines(helpTextWord, "s")}
        MsgBoxUni s, vbInformation, ${toChrW('1Click Word - Hướng Dẫn')}
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

' === CHỨC NĂNG MỚI: ĐÁNH SỐ TRANG CHUẨN NĐ 30/2020/NĐ-CP ===
Public Sub Word_DanhSoTrangND30(control As Object)
    On Error Resume Next
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBoxUni ${toChrW('Chưa có tài liệu Word nào đang mở!')}, vbExclamation, "1Click Word"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Dim sec As Section
    For Each sec In doc.Sections
        sec.PageSetup.DifferentFirstPageHeaderFooter = True
        With sec.Headers(wdHeaderFooterPrimary)
            .LinkToPrevious = False
            With .Range
                .Text = ""
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Font.Name = "Times New Roman"
                .Font.Size = 13.5
                .Font.Bold = False
                .Font.Italic = False
                .Fields.Add Range:=.Range, Type:=wdFieldPage
            End With
        End With
        With sec.Headers(wdHeaderFooterFirstPage)
            .Range.Text = ""
        End With
        sec.Footers(wdHeaderFooterPrimary).Range.Text = ""
        sec.Footers(wdHeaderFooterFirstPage).Range.Text = ""
    Next sec
    Application.ScreenUpdating = True
    
    MsgBoxUni ${toChrW('Đã đánh số trang chuẩn Nghị định 30/2020/NĐ-CP thành công!\n- Vị trí: Ở giữa lề trên (Header)\n- Bắt đầu hiển thị: Từ trang 2 trở đi (Ẩn trang 1)\n- Định dạng: Times New Roman, cỡ 13.5 pt.')}, vbInformation, "1Click Word"
End Sub

' === CHỨC NĂNG MỚI: ĐÁNH SỐ TỪ TRANG BẤT KỲ ===
Public Sub Word_DanhSoTrangBatKy(control As Object)
    On Error GoTo ThoatLoi
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBoxUni ${toChrW('Chưa có tài liệu Word nào đang mở!')}, vbExclamation, "1Click Word"
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
        For Each sec1 In doc.Sections
            sec1.PageSetup.DifferentFirstPageHeaderFooter = False
            If Trim(strViTri) = "1" Then
                With sec1.Headers(wdHeaderFooterPrimary)
                    .PageNumbers.RestartNumberingAtSection = True
                    .PageNumbers.StartingNumber = soBD
                    With .Range
                        .Text = ""
                        .ParagraphFormat.Alignment = wdAlignParagraphCenter
                        .Font.Name = "Times New Roman"
                        .Font.Size = 13.5
                        .Fields.Add Range:=.Range, Type:=wdFieldPage
                    End With
                End With
            Else
                With sec1.Footers(wdHeaderFooterPrimary)
                    .PageNumbers.RestartNumberingAtSection = True
                    .PageNumbers.StartingNumber = soBD
                    With .Range
                        .Text = ""
                        .ParagraphFormat.Alignment = wdAlignParagraphCenter
                        .Font.Name = "Times New Roman"
                        .Font.Size = 13.5
                        .Fields.Add Range:=.Range, Type:=wdFieldPage
                    End With
                End With
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
        
        If Trim(strViTri) = "1" Then
            With curSec.Headers(wdHeaderFooterPrimary)
                .LinkToPrevious = False
                With .Range
                    .Text = ""
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .Font.Name = "Times New Roman"
                    .Font.Size = 13.5
                    .Fields.Add Range:=.Range, Type:=wdFieldPage
                End With
                .PageNumbers.RestartNumberingAtSection = True
                .PageNumbers.StartingNumber = soBD
            End With
        Else
            With curSec.Footers(wdHeaderFooterPrimary)
                .LinkToPrevious = False
                With .Range
                    .Text = ""
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .Font.Name = "Times New Roman"
                    .Font.Size = 13.5
                    .Fields.Add Range:=.Range, Type:=wdFieldPage
                End With
                .PageNumbers.RestartNumberingAtSection = True
                .PageNumbers.StartingNumber = soBD
            End With
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
    MsgBoxUni ${toChrW('Đã hoàn thành đánh số trang thành công!\n- Bắt đầu từ trang: ')} & trangBD & vbCrLf & _
              ${toChrW('- Số hiển thị bắt đầu: ')} & soBD & vbCrLf & _
              ${toChrW('- Các trang phía trước: Đã ẩn toàn bộ số trang.')}, vbInformation, "1Click Word"
    Exit Sub
ThoatLoi:
    Application.ScreenUpdating = True
    MsgBoxUni ${toChrW('Có lỗi xảy ra hoặc bạn đã hủy thao tác.')}, vbExclamation, "1Click Word"
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
    MsgBoxUni ${toChrW('Đã xóa toàn bộ số trang trong toàn bộ văn bản!')}, vbInformation, "1Click Word"
End Sub

' === CHỨC NĂNG MỚI: XOAY NGANG TRANG CHỌN ===
Public Sub Word_XoayNgangTrang(control As Object)
    On Error GoTo LoiXoay
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    
    Dim ans As VbMsgBoxResult
    ans = MsgBoxUni(${toChrW('Hệ thống sẽ xoay NGANG trang hiện tại đang chọn để phù hợp chèn bảng biểu rộng.\nCác trang trước và sau vẫn giữ nguyên khổ DỌC.\n\nBạn có muốn tiếp tục không?')}, vbYesNo + vbQuestion, "1Click Word")
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
    MsgBoxUni ${toChrW('Đã xoay ngang trang số ')} & curPage & ${toChrW(' thành công!')}, vbInformation, "1Click Word"
    Exit Sub
LoiXoay:
    Application.ScreenUpdating = True
    MsgBoxUni ${toChrW('Không thể xoay trang: ')} & Err.Description, vbExclamation, "1Click Word"
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
        MsgBoxUni ${toChrW('Đã dọn dẹp các dòng trống thừa ở cuối văn bản thành công!')}, vbInformation, "1Click Word"
    Else
        MsgBoxUni ${toChrW('Văn bản không có trang trắng hoặc dòng trống thừa ở cuối.')}, vbInformation, "1Click Word"
    End If
End Sub

' === CHỨC NĂNG MỚI: CHUẨN HÓA BẢNG BIỂU ===
Public Sub Word_ChuanHoaBangBieu(control As Object)
    On Error Resume Next
    Dim doc As Document: Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    If doc.Tables.Count = 0 Then
        MsgBoxUni ${toChrW('Không tìm thấy bảng biểu nào trong văn bản!')}, vbExclamation, "1Click Word"
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
    
    MsgBoxUni ${toChrW('Đã chuẩn hóa toàn bộ ')} & doc.Tables.Count & ${toChrW(' bảng biểu trong văn bản:\n- Tự động lặp lại dòng tiêu đề khi bảng dài qua nhiều trang.\n- Chống tình trạng gãy dòng dữ liệu giữa 2 trang.\n- Căn giữa bảng biểu theo lề trang in chuẩn.')}, vbInformation, "1Click Word"
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
    MsgBoxUni ${toChrW('Đã dọn sạch khoảng trắng thừa và chuẩn hóa dấu câu tiếng Việt!')}, vbInformation, "1Click Word"
End Sub
`;

// Read original base Module1_original.vba
let orig = fs.readFileSync('G:\\1Click Build\\Source\\Module1_original.vba', 'utf8');

// Neutralize license and telemetry
orig = orig.replace(/Private Function CheckLicenseActivation\(\) As Boolean[\s\S]*?End Function/, 
`Private Function CheckLicenseActivation() As Boolean
    CheckLicenseActivation = True
End Function`);

orig = orig.replace(/Private Function TT_LicenseStatus\(\) As Boolean[\s\S]*?End Function/,
`Private Function TT_LicenseStatus() As Boolean
    TT_LicenseStatus = True
End Function`);

orig = orig.replace(/Private Function CN_ManifestUrl\(\) As String[\s\S]*?End Function/,
`Private Function CN_ManifestUrl() As String
    CN_ManifestUrl = "https://raw.githubusercontent.com/duoc-tankhai/1Click-Office/main/version.json"
End Function`);

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

const profCountPos = orig.indexOf('Public Function TT_ProfileCount() As Integer');
if (profCountPos !== -1) {
  orig = orig.substring(0, profCountPos) + topDeclares + '\n' + orig.substring(profCountPos);
}

orig = orig + '\n\n' + wordAdvancedSubs;

fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', orig, 'utf8');
console.log('Saved pristine G:\\1Click Build\\Source\\Module1_clean.vba with all advanced features. Length:', orig.length);
