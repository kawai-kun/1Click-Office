Add-Type -AssemblyName System.IO.Compression.FileSystem

$dotm = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm'
$xlam = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\1Click_Excel.xlam'

Write-Output "=== WORD STARTUP CHECK ==="
$zipWord = [System.IO.Compression.ZipFile]::OpenRead($dotm)
$eWord = $zipWord.GetEntry('customUI/customUI.xml')
$reader = New-Object System.IO.StreamReader($eWord.Open())
$xmlWord = $reader.ReadToEnd()
$reader.Close()
$zipWord.Dispose()

Write-Output "Word XML Length: $($xmlWord.Length)"
Write-Output "Has btnDanhSoND30: $($xmlWord.Contains('btnDanhSoND30'))"
Write-Output "Has btnSoTrangBatKyMenu: $($xmlWord.Contains('btnSoTrangBatKyMenu'))"
Write-Output "Has btnXoayNgangTrang: $($xmlWord.Contains('btnXoayNgangTrang'))"
Write-Output "Has btnChuanHoaBang: $($xmlWord.Contains('btnChuanHoaBang'))"
Write-Output "Has btnXoaTrangTrang: $($xmlWord.Contains('btnXoaTrangTrang'))"
Write-Output "Has btnChuanHoaDauCau: $($xmlWord.Contains('btnChuanHoaDauCau'))"
Write-Output "Author Button imageMso: AddressBook in XML: $($xmlWord.Contains('imageMso=""AddressBook""'))"

Write-Output "`n=== EXCEL XLSTART CHECK ==="
$zipExcel = [System.IO.Compression.ZipFile]::OpenRead($xlam)
$eExcel = $zipExcel.GetEntry('customUI/customUI.xml')
$reader2 = New-Object System.IO.StreamReader($eExcel.Open())
$xmlExcel = $reader2.ReadToEnd()
$reader2.Close()
$zipExcel.Dispose()

Write-Output "Excel XML Length: $($xmlExcel.Length)"
Write-Output "Author Button imageMso: AddressBook in XML: $($xmlExcel.Contains('imageMso=""AddressBook""'))"
