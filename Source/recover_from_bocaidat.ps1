$startupWord = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm'
$bakWord = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm.bak'
if (Test-Path $startupWord) { Move-Item $startupWord $bakWord -Force }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.WordBasic.DisableAutoMacros(1)
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open('G:\1Click Build\BoCaiDat_1Click\1Click.dotm')
    $mod = $doc.VBProject.VBComponents.Item("Module1")
    $cm = $mod.CodeModule
    Write-Output "Lines in BoCaiDat: $($cm.CountOfLines)"
    
    $code = $cm.Lines(1, $cm.CountOfLines)
    [System.IO.File]::WriteAllText("G:\1Click Build\Source\Module1_recovered.vba", $code, [System.Text.Encoding]::UTF8)
    Write-Output "Recovered successfully to Module1_recovered.vba"
    $doc.Close(0)
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    if (Test-Path $bakWord) { Move-Item $bakWord $startupWord -Force }
}
