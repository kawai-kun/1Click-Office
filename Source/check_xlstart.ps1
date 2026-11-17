$xlstart = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'
Get-ChildItem $xlstart | Select-Object Name, Length, LastWriteTime
