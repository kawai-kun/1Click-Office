$st = (Get-ChildItem 'C:\Program Files (x86)\Windows Kits' -Recurse -Filter 'signtool.exe' -ErrorAction SilentlyContinue | Select-Object -First 1).FullName
Write-Output "Signtool: $st"
if ($st) {
    & $st sign /sha1 5857F124829744F7344EFADB8A7BC09582F05666 /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 'G:\1Click Build\BoCaiDat_1Click\1Click_Setup.exe'
    Write-Output "Signature exit code: $LASTEXITCODE"
}
