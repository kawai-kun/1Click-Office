$dotmPath = "G:\1Click Build\Output\1Click.dotm"
$vbaPath = "G:\1Click Build\Source\Module1_clean.vba"
$customUIPath = "G:\1Click Build\Source\customUI.xml"

# Step 1: Update customUI.xml in dotm via Zip
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

# Step 2: Inject VBA via Word COM
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
    
    # Compile
    $compileBtn = $word.VBE.CommandBars.FindControl($null, 578)
    if ($compileBtn -ne $null -and $compileBtn.Enabled) {
        $compileBtn.Execute()
        Write-Output "Compiled Word VBA successfully."
    }
    
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
}
