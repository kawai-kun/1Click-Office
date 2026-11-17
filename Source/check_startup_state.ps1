$sW = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm'
$bW = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm.bak'
Write-Output ("sW exists: " + (Test-Path $sW))
Write-Output ("bW exists: " + (Test-Path $bW))
if (Test-Path $bW) { 
    Move-Item $bW $sW -Force
    Write-Output "Restored sW from bW" 
}
# Also copy latest Output to STARTUP to be 100% sure
Copy-Item "G:\1Click Build\Output\1Click.dotm" $sW -Force
Write-Output "Ensured latest 1Click.dotm is in STARTUP."

$sX = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\1Click_Excel.xlam'
$bX = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\1Click_Excel.xlam.bak'
Write-Output ("sX exists: " + (Test-Path $sX))
Write-Output ("bX exists: " + (Test-Path $bX))
if (Test-Path $bX) {
    Move-Item $bX $sX -Force
    Write-Output "Restored sX from bX"
}
Copy-Item "G:\1Click Build\Output\1Click_Excel.xlam" $sX -Force
Write-Output "Ensured latest 1Click_Excel.xlam is in XLSTART."
