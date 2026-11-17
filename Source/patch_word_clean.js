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
3. Đơn Vị Ban Hành: Quản lý thông tin cơ quan ban hành (UBND Phường Tân Khai, TTDVTH...).
4. Tùy Chọn Thể Thức: Linh hoạt cỡ chữ, lề trang, khoảng cách dòng đoạn.`;

const wordSubs = `
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
`;

// Read original base Module1_original.vba to produce pristine Module1_clean.vba
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

// Add top declares right before Public Function TT_ProfileCount
const profCountPos = orig.indexOf('Public Function TT_ProfileCount() As Integer');
if (profCountPos !== -1) {
  orig = orig.substring(0, profCountPos) + topDeclares + '\n' + orig.substring(profCountPos);
}

// Append wordSubs
orig = orig + '\n\n' + wordSubs;

fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', orig, 'utf8');
console.log('Saved pristine G:\\1Click Build\\Source\\Module1_clean.vba. Length:', orig.length);
