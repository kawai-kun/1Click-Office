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

# 2. Update customUI.xml in dotm via Zip
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$zip = [System.IO.Compression.ZipFile]::Open($dotmPath, [System.IO.Compression.ZipArchiveMode]::Update)
$entry = $zip.GetEntry('customUI/customUI.xml')
if ($entry) { $entry.Delete() }
$newEntry = $zip.CreateEntry('customUI/customUI.xml')
$stream = $newEntry.Open()
$bytes = [System.IO.File]::ReadAllBytes($customUIPath)
$stream.Write($bytes, 0, $bytes.Length)
$stream.Close()
$zip.Dispose()
Write-Output "Updated customUI.xml in 1Click.dotm"

# 3. Inject VBA via Word COM
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.AutomationSecurity = 1
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
    
    # Restore STARTUP file
    Copy-Item $dotmPath $startupFile -Force
    if (Test-Path $bakFile) { Remove-Item $bakFile -Force }
    Write-Output "Restored updated 1Click.dotm to STARTUP"
}
