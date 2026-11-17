Add-Type -AssemblyName System.IO.Compression.FileSystem

$sourceDir = "G:\1Click Build\BoCaiDat_1Click"
$zipPath = "G:\1Click Build\BoCaiDat_1Click_v3.0.zip"

if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

[System.IO.Compression.ZipFile]::CreateFromDirectory($sourceDir, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)

$fi = Get-Item $zipPath
Write-Output "Created $zipPath (Size: $($fi.Length) bytes)"
