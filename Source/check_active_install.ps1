$wordStartup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$excelStartup = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'

Write-Output "=== WORD STARTUP ==="
Get-ChildItem $wordStartup | Select-Object Name, Length, LastWriteTime

Write-Output "=== EXCEL XLSTART ==="
Get-ChildItem $excelStartup | Select-Object Name, Length, LastWriteTime
