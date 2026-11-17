$xlstart = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'
$oldFile = Join-Path $xlstart 'ChuanHoa_Excel.xlam'
$newFile = Join-Path $xlstart '1Click_Excel.xlam'
$srcFile = 'G:\1Click Build\Output\1Click_Excel.xlam'

# Copy to BoCaiDat_1Click
Copy-Item $srcFile 'G:\1Click Build\BoCaiDat_1Click\1Click_Excel.xlam' -Force
Write-Output "Copied to G:\1Click Build\BoCaiDat_1Click\1Click_Excel.xlam"

# Copy to XLSTART
if (Test-Path $oldFile) {
    Remove-Item $oldFile -Force -ErrorAction SilentlyContinue
    Write-Output "Removed old ChuanHoa_Excel.xlam"
}
Copy-Item $srcFile $newFile -Force
Write-Output "Installed 1Click_Excel.xlam to XLSTART"

# Verify XLSTART
Get-ChildItem $xlstart | Select-Object Name, Length, LastWriteTime
