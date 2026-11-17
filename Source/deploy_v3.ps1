# 1. Copy to BoCaiDat_1Click
Copy-Item "G:\1Click Build\Output\1Click.dotm" "G:\1Click Build\BoCaiDat_1Click\1Click.dotm" -Force
Copy-Item "G:\1Click Build\Output\1Click_Excel.xlam" "G:\1Click Build\BoCaiDat_1Click\1Click_Excel.xlam" -Force
Write-Output "Copied 1Click.dotm and 1Click_Excel.xlam to BoCaiDat_1Click."

# 2. Copy to Active Startup Locations
$wordStartup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$excelStartup = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'

Copy-Item "G:\1Click Build\Output\1Click.dotm" (Join-Path $wordStartup "1Click.dotm") -Force
Copy-Item "G:\1Click Build\Output\1Click_Excel.xlam" (Join-Path $excelStartup "1Click_Excel.xlam") -Force
Unblock-File (Join-Path $wordStartup "1Click.dotm")
Unblock-File (Join-Path $excelStartup "1Click_Excel.xlam")
Write-Output "Updated active installations in Word STARTUP and Excel XLSTART."

# 3. Recompile Inno Setup
& "C:\Program Files\Inno Setup 7\ISCC.exe" "G:\1Click Build\BoCaiDat_1Click\setup_script.iss"
Write-Output "Inno Setup compiled successfully."

# 4. Sign Setup Executable
$signtool = "C:\Program Files (x86)\Microsoft SDKs\ClickOnce\SignTool\signtool.exe"
& $signtool sign /sha1 5857F124829744F7344EFADB8A7BC09582F05666 /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "G:\1Click Build\BoCaiDat_1Click\1Click_Setup.exe"
Write-Output "Signed 1Click_Setup.exe."

# 5. Create Zip Package
Add-Type -AssemblyName System.IO.Compression.FileSystem
$sourceDir = "G:\1Click Build\BoCaiDat_1Click"
$zipPath = "G:\1Click Build\BoCaiDat_1Click_v3.0.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory($sourceDir, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
Write-Output "Created G:\1Click Build\BoCaiDat_1Click_v3.0.zip successfully."
