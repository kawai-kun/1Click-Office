Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$customUIPath = "G:\1Click Build\Source\customUI.xml"
$customUIBytes = [System.IO.File]::ReadAllBytes($customUIPath)

$dotmTargets = @(
    "G:\1Click Build\Output\1Click.dotm",
    "G:\1Click Build\BoCaiDat_1Click\1Click.dotm",
    (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm')
)

foreach ($dotm in $dotmTargets) {
    if (Test-Path $dotm) {
        Write-Output "Updating OpenXML in: $dotm"
        $zip = [System.IO.Compression.ZipFile]::Open($dotm, [System.IO.Compression.ZipArchiveMode]::Update)
        
        # 1. Update customUI/customUI.xml
        $e1 = $zip.GetEntry('customUI/customUI.xml')
        if ($e1) { $e1.Delete() }
        $n1 = $zip.CreateEntry('customUI/customUI.xml')
        $s1 = $n1.Open()
        $s1.Write($customUIBytes, 0, $customUIBytes.Length)
        $s1.Close()

        # 2. Update customUI/customUI14.xml
        $e2 = $zip.GetEntry('customUI/customUI14.xml')
        if ($e2) { $e2.Delete() }
        $n2 = $zip.CreateEntry('customUI/customUI14.xml')
        $s2 = $n2.Open()
        $s2.Write($customUIBytes, 0, $customUIBytes.Length)
        $s2.Close()

        $zip.Dispose()
        Unblock-File $dotm
        Write-Output "Updated successfully."
    }
}
