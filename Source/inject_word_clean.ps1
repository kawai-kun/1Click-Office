$startupFile = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm'
$bakFile = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm.bak'

# 1. Temporarily move out of STARTUP so Word starts completely clean
if (Test-Path $startupFile) {
    Move-Item $startupFile $bakFile -Force
    Write-Output "Moved STARTUP file to bak"
}

$vbaPath = "G:\1Click Build\Source\Module1_clean.vba"
$dotmPath = "G:\1Click Build\Output\1Click.dotm"
$customUIPath = "G:\1Click Build\Source\customUI.xml"

# 2. Update customUI.xml & customUI14.xml via Zip
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$customUIBytes = [System.IO.File]::ReadAllBytes($customUIPath)

$zip = [System.IO.Compression.ZipFile]::Open($dotmPath, [System.IO.Compression.ZipArchiveMode]::Update)
$e1 = $zip.GetEntry('customUI/customUI.xml')
if ($e1) { $e1.Delete() }
$n1 = $zip.CreateEntry('customUI/customUI.xml')
$s1 = $n1.Open()
$s1.Write($customUIBytes, 0, $customUIBytes.Length)
$s1.Close()

$e2 = $zip.GetEntry('customUI/customUI14.xml')
if ($e2) { $e2.Delete() }
$n2 = $zip.CreateEntry('customUI/customUI14.xml')
$s2 = $n2.Open()
$s2.Write($customUIBytes, 0, $customUIBytes.Length)
$s2.Close()

$zip.Dispose()
Write-Output "Updated customUI XML in 1Click.dotm"

# 3. Inject VBA via Word COM
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.WordBasic.DisableAutoMacros(1)
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($dotmPath)
    Write-Output "Word Document opened: $($doc.Name)"
    $mod = $doc.VBProject.VBComponents.Item("Module1")
    $cm = $mod.CodeModule
    Write-Output "Existing lines: $($cm.CountOfLines)"
    $cm.DeleteLines(1, $cm.CountOfLines)
    $cm.AddFromFile($vbaPath)
    Write-Output "New lines count: $($cm.CountOfLines)"
    
    $doc.Save()
    Write-Output "Saved 1Click.dotm successfully."
    $doc.Close(0)
}
catch {
    Write-Error $_
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    # Restore and update STARTUP file
    Copy-Item $dotmPath $startupFile -Force
    if (Test-Path $bakFile) { Remove-Item $bakFile -Force }
    Unblock-File $startupFile
    Write-Output "Restored and updated 1Click.dotm in STARTUP"
}
