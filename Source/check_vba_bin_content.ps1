Add-Type -AssemblyName System.IO.Compression.FileSystem

$dotm = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm'
$zip = [System.IO.Compression.ZipFile]::OpenRead($dotm)
$entry = $zip.GetEntry('word/vbaProject.bin')

$ms = New-Object System.IO.MemoryStream
$stream = $entry.Open()
$stream.CopyTo($ms)
$stream.Close()
$zip.Dispose()

$bytes = $ms.ToArray()
$str = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetString($bytes)

Write-Output "Decompressed vbaProject.bin length: $($bytes.Length)"
Write-Output "Has old 'Thanh Phu': $($str.Contains('Thanh Phu'))"
Write-Output "Has old 'VPDU xa': $($str.Contains('VPDU xa'))"
Write-Output "Has new 'Tran Quoc Duoc': $($str.Contains('Tran Quoc Duoc'))"
Write-Output "Has new 'Tan Khai': $($str.Contains('Tan Khai'))"
