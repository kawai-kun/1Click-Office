$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.WordBasic.DisableAutoMacros(1)
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open('G:\1Click Build\Output\1Click.dotm')
    Write-Output "Word Add-in loaded: $($doc.Name)"
    
    # Test creating a test doc
    $testDoc = $word.Documents.Add()
    $testDoc.Content.Text = "UBND PHUONG TAN KHAI`tCONG HOA XA HOI CHU NGHIA VIET NAM`r`nTRUNG TAM DICH VU TONG HOP`tDoc lap - Tu do - Hanh phuc`r`nSo:   /TT-UBND`tTan Khai, ngay   thang 10 nam 2026`r`n`r`nTO TRINH`r`nVe viec to chuc hoat dong... [Ho va ten]`r`n"
    
    $word.Run("1Click.dotm!Module1.ChuanHoa_VanBanDang_Silent")
    
    # Read doc variable
    $report = $testDoc.Variables.Item("TheThuc_BaoCao").Value
    Write-Output "`n=== LOGIC REPORT IN WORD ==="
    [System.IO.File]::WriteAllText("G:\1Click Build\Source\report_out.txt", $report, [System.Text.Encoding]::UTF8)
    Write-Output $report
    Write-Output "============================="
    
    $testDoc.Close(0)
    $doc.Close(0)
}
catch {
    Write-Error $_
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
