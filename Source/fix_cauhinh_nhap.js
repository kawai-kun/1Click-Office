const fs = require('fs');

let content = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8');

// The clean TT_CauHinhNhap function
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

fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', content, 'utf8');
console.log("Successfully restored TT_CauHinhNhap.");
