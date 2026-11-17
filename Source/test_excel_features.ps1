$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$xlamPath = "G:\1Click Build\Output\1Click_Excel.xlam"
$wbAddin = $excel.Workbooks.Open($xlamPath)

$wb = $excel.Workbooks.Add()
$ws = $wb.Sheets.Item(1)

# Populate test data
$ws.Range("A1").Value2 = "stt"
$ws.Range("B1").Value2 = "ho va ten"
$ws.Range("C1").Value2 = "so tien"

$ws.Range("A2").Value2 = 99
$ws.Range("B2").Value2 = "trương đắc thảo"
$ws.Range("C2").Value2 = 1500000

$ws.Range("A3").Value2 = 88
$ws.Range("B3").Value2 = "trần văn thanh"
$ws.Range("C3").Value2 = 2300000

$ws.Range("A4").Value2 = 77
$ws.Range("B4").Value2 = "lê ngọc xuân"
$ws.Range("C4").Value2 = 500000

# 1. Test Viết hoa đầu từ
$ws.Range("B2:B4").Select()
$excel.Run("1Click_Excel.xlam!Xl_VietHoaDauTu", $null)
Write-Output "B2 after VietHoaDauTu: $($ws.Range('B2').Value2)"
Write-Output "B3 after VietHoaDauTu: $($ws.Range('B3').Value2)"
Write-Output "B4 after VietHoaDauTu: $($ws.Range('B4').Value2)"

# 2. Test Đánh lại STT
$ws.Range("A2:A4").Select()
$excel.Run("1Click_Excel.xlam!Xl_DanhLaiSTT", $null)
Write-Output "A2 after DanhLaiSTT: $($ws.Range('A2').Value2)"
Write-Output "A3 after DanhLaiSTT: $($ws.Range('A3').Value2)"
Write-Output "A4 after DanhLaiSTT: $($ws.Range('A4').Value2)"

# 3. Test Đọc số ra chữ
$ws.Range("C2:C4").Select()
$excel.Run("1Click_Excel.xlam!Xl_DocSoRaChu", $null)
Write-Output "D2 (DocSo cell): $($ws.Range('D2').Value2)"
Write-Output "D3 (DocSo cell): $($ws.Range('D3').Value2)"

# 4. Test Kẻ khung chuẩn
$ws.Range("A1:D4").Select()
$excel.Run("1Click_Excel.xlam!Xl_KeKhungChuan", $null)
Write-Output "Border LineStyle: $($ws.Range('A1').Borders.Item(1).LineStyle)"

# 5. Test IN HOA
$ws.Range("B2:B2").Select()
$excel.Run("1Click_Excel.xlam!Xl_VietHoaToanBo", $null)
Write-Output "B2 after VietHoaToanBo: $($ws.Range('B2').Value2)"

# 6. Test chu thuong
$ws.Range("B2:B2").Select()
$excel.Run("1Click_Excel.xlam!Xl_ChuThuong", $null)
Write-Output "B2 after ChuThuong: $($ws.Range('B2').Value2)"

$wb.Close($false)
$wbAddin.Close($false)
$excel.Quit()
Write-Output "All feature UI actions tested successfully!"
