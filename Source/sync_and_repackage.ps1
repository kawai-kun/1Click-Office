$excelStartup = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\1Click_Excel.xlam'
Copy-Item "G:\1Click Build\Output\1Click_Excel.xlam" $excelStartup -Force
Copy-Item "G:\1Click Build\Output\1Click_Excel.xlam" "G:\1Click Build\BoCaiDat_1Click\1Click_Excel.xlam" -Force
Copy-Item "G:\1Click Build\Output\1Click.dotm" "G:\1Click Build\BoCaiDat_1Click\1Click.dotm" -Force

Unblock-File $excelStartup
Unblock-File (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm')

Write-Output "Files synced to BoCaiDat_1Click and XLSTART."

# Recompile Inno Setup
& "C:\Program Files\Inno Setup 7\ISCC.exe" "G:\1Click Build\BoCaiDat_1Click\setup_script.iss"

# Sign
$signtool = "C:\Program Files (x86)\Microsoft SDKs\ClickOnce\SignTool\signtool.exe"
& $signtool sign /sha1 5857F124829744F7344EFADB8A7BC09582F05666 /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "G:\1Click Build\BoCaiDat_1Click\1Click_Setup.exe"

# Zip
Add-Type -AssemblyName System.IO.Compression.FileSystem
$sourceDir = "G:\1Click Build\BoCaiDat_1Click"
$zipPath = "G:\1Click Build\BoCaiDat_1Click_v3.0.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory($sourceDir, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
Write-Output "Complete deployment finished successfully."
