const fs = require('fs');

let content = fs.readFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', 'utf8');

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

fs.writeFileSync('G:\\1Click Build\\Source\\Module1_clean.vba', content, 'utf8');
console.log("Successfully fixed TT_ChuanKinhGui.");
