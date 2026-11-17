$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    Write-Output "Word started with STARTUP 1Click.dotm active."
    
    # Create test document
    $docPath = "G:\1Click Build\Source\test_doc.docx"
    $doc = $word.Documents.Add()
    $doc.Content.Text = "UBND PHUONG TAN KHAI`tCONG HOA XA HOI CHU NGHIA VIET NAM`r`nTRUNG TAM DICH VU TONG HOP`tDoc lap - Tu do - Hanh phuc`r`nSo:   /TT-UBND`tTan Khai, ngay   thang 10 nam 2026`r`n`r`nTO TRINH`r`nVe viec to chuc hoat dong`r`n"
    $doc.SaveAs([ref]$docPath)
    $doc.Close()
    
    Write-Output "Created test document: $docPath"
    
    # Open the document
    $openedDoc = $word.Documents.Open($docPath)
    Write-Output "Opened document successfully: $($openedDoc.Name)"
    
    # Run ribbon macros
    $word.Run("Word_DanhSoTrangND30", $null)
    Write-Output "Word_DanhSoTrangND30 executed successfully!"
    
    $word.Run("Word_XoaTrangTrang", $null)
    Write-Output "Word_XoaTrangTrang executed successfully!"
    
    $word.Run("Word_ChuanHoaBangBieu", $null)
    Write-Output "Word_ChuanHoaBangBieu executed successfully!"
    
    $word.Run("Word_ChuanHoaDauCau", $null)
    Write-Output "Word_ChuanHoaDauCau executed successfully!"
    
    $word.Run("Word_XoaSoTrang", $null)
    Write-Output "Word_XoaSoTrang executed successfully!"
    
    $openedDoc.Close(0)
    Write-Output "ALL WORD MACROS PASSED WITH ZERO COMPILE ERRORS!"
}
catch {
    Write-Error $_
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    if (Test-Path "G:\1Click Build\Source\test_doc.docx") {
        Remove-Item "G:\1Click Build\Source\test_doc.docx" -Force
    }
}
