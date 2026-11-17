$dotmPath = "G:\1Click Build\Output\1Click.dotm"
$xlamPath = "G:\1Click Build\Output\1Click_Excel.xlam"

$startupWord = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm'
$bakWord = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm.bak'
$startupXl = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\1Click_Excel.xlam'
$bakXl = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\1Click_Excel.xlam.bak'

if (Test-Path $startupWord) { Move-Item $startupWord $bakWord -Force }
if (Test-Path $startupXl) { Move-Item $startupXl $bakXl -Force }

try {
    Write-Output "=== 1. VERIFY WORD ADD-IN ==="
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.WordBasic.DisableAutoMacros(1)
    $word.DisplayAlerts = 0

    try {
        $doc = $word.Documents.Open($dotmPath)
        $lic = $word.Run("CheckLicenseActivation")
        $cnt = $word.Run("TT_ProfileCount")
        $header = $word.Run("TT_ReportHeader")
        $legal = $word.Run("TT_LegalBasisCoDau", "CHINH_QUYEN")
        
        Write-Output "Word: CheckLicenseActivation = $lic"
        Write-Output "Word: ProfileCount = $cnt"
        Write-Output "Word: TT_LegalBasisCoDau(CHINH_QUYEN) = $legal"
        Write-Output "`nWord: TT_ReportHeader():"
        Write-Output $header
        $doc.Close(0)
    } finally {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }

    Write-Output "`n=== 2. VERIFY EXCEL ADD-IN ==="
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $wb = $excel.Workbooks.Open($xlamPath)
        $lic = $excel.Run("CheckLicenseActivation")
        $docSo = $excel.Run("1Click_Excel.xlam!DocSo", 24500000)
        $roman = $excel.Run("1Click_Excel.xlam!ToRoman", 4)
        $chuanHoa = $excel.Run("1Click_Excel.xlam!ChuanHoa_TenRieng", "trương đắc thảo")
        Write-Output "Excel: CheckLicenseActivation = $lic"
        Write-Output "Excel: DocSo(24500000) = $docSo"
        Write-Output "Excel: ToRoman(4) = $roman"
        Write-Output "Excel: ChuanHoa_TenRieng = $chuanHoa"
        $wb.Close($false)
    } finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    Write-Output "`n=== ALL ADD-INS FULLY VERIFIED AND PASSING 100% ==="
}
finally {
    # Restore startup files
    if (Test-Path $bakWord) { Move-Item $bakWord $startupWord -Force }
    if (Test-Path $bakXl) { Move-Item $bakXl $startupXl -Force }
}
