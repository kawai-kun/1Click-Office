$xlamPath = "G:\1Click Build\Output\1Click_Excel.xlam"

Write-Output "Starting Excel COM verification for 1Click_Excel.xlam..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1 # msoAutomationSecurityLow

try {
    $wb = $excel.Workbooks.Open($xlamPath)
    Write-Output "Add-in workbook loaded: $($wb.Name)"
    
    # Test DocSo
    $docSo1 = $excel.Run("1Click_Excel.xlam!DocSo", 1250000)
    Write-Output "TEST 1: DocSo(1250000) = $docSo1"
    
    $docSo0 = $excel.Run("1Click_Excel.xlam!DocSo", 0)
    Write-Output "TEST 2: DocSo(0) = $docSo0"
    
    $docSoAm = $excel.Run("1Click_Excel.xlam!DocSo", -3500000)
    Write-Output "TEST 3: DocSo(-3500000) = $docSoAm"
    
    # Test ChuanHoa_TenRieng
    $name1 = $excel.Run("1Click_Excel.xlam!ChuanHoa_TenRieng", "trương đắc thảo")
    Write-Output "TEST 4: ChuanHoa_TenRieng('trương đắc thảo') = $name1"
    
    $name2 = $excel.Run("1Click_Excel.xlam!ChuanHoa_TenRieng", "   LÊ   thị   phương  ")
    Write-Output "TEST 5: ChuanHoa_TenRieng('   LÊ   thị   phương  ') = $name2"
    
    # Test UCaseUnicode
    $upper = $excel.Run("1Click_Excel.xlam!UCaseUnicode", "quỳnh lưu - nghệ an")
    Write-Output "TEST 6: UCaseUnicode('quỳnh lưu - nghệ an') = $upper"
    
    # Test LCaseUnicode
    $lower = $excel.Run("1Click_Excel.xlam!LCaseUnicode", "UBND PHƯỜNG TÂN KHAI")
    Write-Output "TEST 7: LCaseUnicode('UBND PHƯỜNG TÂN KHAI') = $lower"

    # Test Worksheet formula integration
    $wbTest = $excel.Workbooks.Add()
    $ws = $wbTest.Sheets.Item(1)
    $ws.Range("A1").Value2 = 5000000
    $ws.Range("B1").Formula = "=DocSo(A1)"
    $formulaResult = $ws.Range("B1").Text
    Write-Output "TEST 8: Worksheet formula =DocSo(A1) = $formulaResult"
    $wbTest.Close($false)
    
    Write-Output "=== ALL TESTS COMPLETED SUCCESSFULLY ==="
}
catch {
    Write-Error $_
}
finally {
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
