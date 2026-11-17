& "C:\Program Files\Inno Setup 7\ISCC.exe" "G:\1Click Build\BoCaiDat_1Click\setup_script.iss"
$signtool = "C:\Program Files (x86)\Microsoft SDKs\ClickOnce\SignTool\signtool.exe"
& $signtool sign /sha1 5857F124829744F7344EFADB8A7BC09582F05666 /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "G:\1Click Build\BoCaiDat_1Click\1Click_Setup.exe"

Add-Type -AssemblyName System.IO.Compression.FileSystem
$sourceDir = "G:\1Click Build\BoCaiDat_1Click"
$zipPath = "G:\1Click Build\BoCaiDat_1Click_v3.0.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory($sourceDir, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
Write-Output "Installer recompiled, signed, and zip updated successfully."
