$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $xlPath = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\1Click_Excel.xlam'
    $wb = $excel.Workbooks.Open($xlPath)
    Write-Output "Excel loaded: $($wb.Name)"
    $lic = $excel.Run("1Click_Excel.xlam!CheckLicenseActivation")
    $docSo = $excel.Run("1Click_Excel.xlam!DocSo", 24500000)
    $roman = $excel.Run("1Click_Excel.xlam!ToRoman", 4)
    $chuanHoa = $excel.Run("1Click_Excel.xlam!ChuanHoa_TenRieng", "trương đắc thảo")
    
    Write-Output "Excel: CheckLicenseActivation = $lic"
    Write-Output "Excel: DocSo(24500000) = $docSo"
    Write-Output "Excel: ToRoman(4) = $roman"
    Write-Output "Excel: ChuanHoa_TenRieng('trương đắc thảo') = $chuanHoa"
    $wb.Close($false)
}
catch {
    Write-Error $_
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
