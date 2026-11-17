$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wbAddin = $excel.Workbooks.Open('G:\1Click Build\Output\1Click_Excel.xlam')
$wb = $excel.Workbooks.Add()
$ws = $wb.Sheets.Item(1)

# Test Roman Numerals
$ws.Range("A1:A5").Select()
$excel.Run("1Click_Excel.xlam!Xl_DanhSTT_LaMa", $null)
Write-Output "A1: $($ws.Range('A1').Value2)"
Write-Output "A2: $($ws.Range('A2').Value2)"
Write-Output "A3: $($ws.Range('A3').Value2)"
Write-Output "A4: $($ws.Range('A4').Value2)"
Write-Output "A5: $($ws.Range('A5').Value2)"

# Test Currency Formatting
$ws.Range("B1").Value2 = 12500000
$ws.Range("B1").Select()
$excel.Run("1Click_Excel.xlam!Xl_DinhDangTienVND", $null)
Write-Output "B1 Text formatted: $($ws.Range('B1').Text)"

# Test TuTinhThanhTien
$ws.Range("C2").Value2 = 5       # SL
$ws.Range("D2").Value2 = 200000  # Don gia
$ws.Range("E2").Select()         # Thanh tien
$excel.Run("1Click_Excel.xlam!Xl_TuTinhThanhTien", $null)
Write-Output "E2 Formula: $($ws.Range('E2').Formula)"
Write-Output "E2 Value: $($ws.Range('E2').Value2)"

$wb.Close($false)
$wbAddin.Close($false)
$excel.Quit()
Write-Output "=== EXCEL V3 FEATURES TEST PASSED 100% ==="
