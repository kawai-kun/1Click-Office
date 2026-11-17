const fs = require('fs');

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

let vba = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8');

// Check if already patched
if (!vba.includes('Public Sub RibbonThongTinTacGia')) {
  const wordAdditions = `
' === UNICODE MSGBOX & AUTHOR SUPPORT FOR WORD ===
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

Public Function MsgBoxUni(ByVal prompt As String, _
                         Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                         Optional ByVal title As String = "1Click Office") As VbMsgBoxResult
    On Error Resume Next
    MsgBoxUni = MsgBoxW(0, StrPtr(prompt), StrPtr(title), buttons)
End Function

Public Sub RibbonThongTinTacGia(control As Object)
    Dim s As String
    s = ${toChrW('1CLICK OFFICE 3.0 ENTERPRISE\nBộ Tiện Ích Chuẩn Hóa Văn Bản & Bảng Tính Hành Chính\n\n- Tác giả: Trần Quốc Được\n- Đơn vị: Trung tâm Dịch vụ Tổng hợp Phường Tân Khai\n- Trạng thái bản quyền: VĨNH VIỄN (Đã kích hoạt)\n\nHỗ trợ kỹ thuật & Đóng góp ý kiến:\nVui lòng liên hệ UBND Phường Tân Khai / TTDVTH Phường Tân Khai.')}
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
        s = ${toChrW('HƯỚNG DẪN NHANH 1CLICK WORD 3.0:\n\n1. Chuẩn Hóa Thể Thức: Tự động căn chỉnh văn bản chuẩn Nghị định 30/2020/NĐ-CP và quy định của Đảng, Đoàn thể.\n2. Tạo Văn Bản Mẫu: Đầy đủ các mẫu Tờ trình, Kế hoạch, Quyết định, Báo cáo, Giấy mời...\n3. Đơn Vị Ban Hành: Quản lý và chọn cơ quan ban hành mặc định (UBND Phường Tân Khai, TTDVTH...).\n4. Tùy Chọn Thể Thức: Linh hoạt cỡ chữ, lề trang, giãn dòng, giãn đoạn theo ý muốn.')}
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
`;

  vba = vba + '\n\n' + wordAdditions;
  fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', vba, 'utf8');
  console.log('Successfully appended Word additions to Module1_clean.vba');
}
