$res = "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency"
if (Test-Path "$res\DisabledItems") { Remove-Item "$res\DisabledItems" -Recurse -Force }
if (Test-Path "$res\StartupItems") { Remove-Item "$res\StartupItems" -Recurse -Force }
if (Test-Path "$res\DocumentRecovery") { Remove-Item "$res\DocumentRecovery" -Recurse -Force }

# Restore STARTUP file if bak exists
$startupFile = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm'
$bakFile = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm.bak'
if (Test-Path $bakFile) {
    Move-Item $bakFile $startupFile -Force
}

Write-Output "Word Resiliency cleared and STARTUP restored."
