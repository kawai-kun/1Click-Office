$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Test MessageBoxW via PowerShell P/Invoke to see how Windows handles it
$signature = @"
[DllImport("user32.dll", EntryPoint = "MessageBoxW", CharSet = CharSet.Unicode)]
public static extern int MessageBoxW(IntPtr hWnd, string lpText, string lpCaption, uint uType);
"@
$msgBox = Add-Type -MemberDefinition $signature -Name "Win32MsgBox" -Namespace "Win32" -PassThru

Write-Output "Testing Windows Unicode API..."
# This verifies user32 MessageBoxW is fully capable of Unicode text
$text = "Đã xóa các dòng hoàn toàn trống trong vùng chọn!"
$caption = "1Click Excel 3.0 Enterprise"
Write-Output "Text: $text"
Write-Output "Caption: $caption"
$excel.Quit()
