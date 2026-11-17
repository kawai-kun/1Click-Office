$paths = @(
    (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'),
    (Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'),
    'C:\Program Files\Microsoft Office\root\Office16\STARTUP',
    'C:\Program Files\Microsoft Office\root\Office16\XLSTART',
    (Join-Path $env:APPDATA 'Microsoft\AddIns')
)

foreach ($p in $paths) {
    if (Test-Path $p) {
        Write-Output "Folder: $p"
        Get-ChildItem $p -Recurse -ErrorAction SilentlyContinue | Select-Object Name, Length, LastWriteTime
    }
}
