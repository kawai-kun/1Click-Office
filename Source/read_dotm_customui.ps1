Add-Type -AssemblyName System.IO.Compression.FileSystem

$dotm = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm'
$zip = [System.IO.Compression.ZipFile]::OpenRead($dotm)

foreach ($e in $zip.Entries) {
    if ($e.FullName -like "*customUI*") {
        Write-Output "Entry: $($e.FullName) ($($e.Length) bytes)"
        $reader = New-Object System.IO.StreamReader($e.Open())
        $content = $reader.ReadToEnd()
        $reader.Close()
        Write-Output "Has grpTacGia: $($content.Contains('grpTacGia'))"
        Write-Output "Has ThongTinTacGia: $($content.Contains('ThongTinTacGia'))"
        Write-Output "Tail 500 chars:"
        $len = $content.Length
        $start = [Math]::Max(0, $len - 500)
        Write-Output $content.Substring($start)
    }
}

$zip.Dispose()
