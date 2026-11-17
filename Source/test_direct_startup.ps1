$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    Write-Output "Word loaded."
    $header = $word.Run("TT_ReportHeader")
    Write-Output "`n=== TT_ReportHeader() OUTPUT ==="
    Write-Output $header
    Write-Output "================================="
    [System.IO.File]::WriteAllText("G:\1Click Build\Source\header_result.txt", $header, [System.Text.Encoding]::UTF8)
}
catch {
    Write-Error $_
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
