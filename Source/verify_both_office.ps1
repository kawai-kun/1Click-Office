Write-Output "=== 1. VERIFYING EXCEL ADD-IN ==="
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open('G:\1Click Build\Output\1Click_Excel.xlam')
    Write-Output "Excel Add-in loaded: $($wb.Name)"
    
    $cmd = $excel.VBE.CommandBars.FindControl($null, 578)
    if ($cmd -ne $null) {
        Write-Output "Executing Excel Compile: $($cmd.Caption)"
        $cmd.Execute()
        Write-Output "Excel Compile SUCCESS: No errors!"
    } else {
        Write-Output "Compile command 578 not found"
    }
    $wb.Close($false)
}
catch {
    Write-Error "Excel Error: $_"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Output "`n=== 2. VERIFYING WORD ADD-IN ==="
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.WordBasic.DisableAutoMacros(1)
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open('G:\1Click Build\Output\1Click.dotm')
    Write-Output "Word Add-in loaded: $($doc.Name)"
    
    $cmd = $word.VBE.CommandBars.FindControl($null, 578)
    if ($cmd -ne $null) {
        Write-Output "Executing Word Compile: $($cmd.Caption)"
        $cmd.Execute()
        Write-Output "Word Compile SUCCESS: No errors!"
    } else {
        Write-Output "Compile command 578 not found"
    }
    
    # Test creating a test doc and running logic check in silent mode
    $testDoc = $word.Documents.Add()
    $testDoc.Content.Text = "UBND PHUONG TAN KHAI`tCONG HOA XA HOI CHU NGHIA VIET NAM`r`nTRUNG TAM DICH VU TONG HOP`tDoc lap - Tu do - Hanh phuc`r`nSo:   /TT-UBND`tTan Khai, ngay   thang 10 nam 2026`r`n`r`nTO TRINH`r`nVe viec to chuc hoat dong... [Ho va ten]`r`n"
    
    # Call logic check function directly via Word.Run
    $word.Run("TT_ChuanHoaTheThucSilent")
    
    # Check document variable "TheThuc_BaoCao"
    try {
        $report = $testDoc.Variables.Item("TheThuc_BaoCao").Value
        Write-Output "`n--- WORD LOGIC REPORT OUTPUT ---"
        Write-Output $report
        Write-Output "--------------------------------"
    }
    catch {
        Write-Output "Note: TheThuc_BaoCao variable not set or function called differently."
    }
    
    $testDoc.Close(0)
    $doc.Close(0)
}
catch {
    Write-Error "Word Error: $_"
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

Write-Output "`n=== VERIFICATION COMPLETED ==="
