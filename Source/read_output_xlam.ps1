Add-Type -AssemblyName System.IO.Compression.FileSystem

$xlam = 'G:\1Click Build\Output\1Click_Excel.xlam'
$zip = [System.IO.Compression.ZipFile]::OpenRead($xlam)

foreach ($e in $zip.Entries) {
    if ($e.FullName -like "*customUI*") {
        Write-Output "Entry: $($e.FullName) ($($e.Length) bytes)"
        $reader = New-Object System.IO.StreamReader($e.Open())
        $content = $reader.ReadToEnd()
        $reader.Close()
        Write-Output "Has grpTacGia: $($content.Contains('grpTacGia'))"
        Write-Output "Has ThongTinTacGia: $($content.Contains('ThongTinTacGia'))"
    }
}

$zip.Dispose()
