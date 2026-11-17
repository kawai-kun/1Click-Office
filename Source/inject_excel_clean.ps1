$xlstartFile = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\1Click_Excel.xlam'
$bakFile = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\1Click_Excel.xlam.bak'

# 1. Temporarily move out of XLSTART
if (Test-Path $xlstartFile) {
    Move-Item $xlstartFile $bakFile -Force
    Write-Output "Moved XLSTART file to bak"
}

$vbaPath = "G:\1Click Build\Source\Excel_Module1_final.vba"
$xlamPath = "G:\1Click Build\Output\1Click_Excel.xlam"

Write-Output "Opening Excel COM to inject VBA..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($xlamPath)
    Write-Output "Workbook opened: $($wb.Name)"
    $proj = $wb.VBProject
    Write-Output "VBProject name: $($proj.Name)"
    
    $mod = $proj.VBComponents.Item("Module1")
    $cm = $mod.CodeModule
    Write-Output "Existing lines: $($cm.CountOfLines)"
    
    $cm.DeleteLines(1, $cm.CountOfLines)
    $cm.AddFromFile($vbaPath)
    Write-Output "New lines count: $($cm.CountOfLines)"
    
    $wb.Save()
    Write-Output "Saved successfully."
    $wb.Close($true)
}
catch {
    Write-Error $_
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    # Restore and copy to XLSTART
    Copy-Item $xlamPath $xlstartFile -Force
    if (Test-Path $bakFile) { Remove-Item $bakFile -Force }
    Unblock-File $xlstartFile
    Write-Output "Restored and updated 1Click_Excel.xlam in XLSTART"
}
