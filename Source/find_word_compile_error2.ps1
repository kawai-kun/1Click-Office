$startupWord = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm'
$bakWord = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\1Click.dotm.bak'
if (Test-Path $startupWord) { Move-Item $startupWord $bakWord -Force }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.WordBasic.DisableAutoMacros(1)
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open('G:\1Click Build\Output\1Click.dotm')
    Write-Output "Document opened: $($doc.Name)"
    
    $vbe = $word.VBE
    $proj = $doc.VBProject
    Write-Output "Doc VBProject: $($proj.Name)"
    $vbe.ActiveVBProject = $proj
    
    # Show code pane of Module1
    $mod = $proj.VBComponents.Item("Module1")
    $mod.CodeModule.CodePane.Show()
    
    $menuBar = $vbe.CommandBars.Item("Menu Bar")
    $debugMenu = $menuBar.Controls.Item("Debug")
    $compileBtn = $debugMenu.Controls.Item(1)
    Write-Output "Compile button caption: $($compileBtn.Caption)"
    
    $compileBtn.Execute()
    
    # Check selection in CodePane
    $pane = $vbe.ActiveCodePane
    if ($pane -ne $null) {
        $startLine = 0; $startCol = 0; $endLine = 0; $endCol = 0
        $pane.GetSelection([ref]$startLine, [ref]$startCol, [ref]$endLine, [ref]$endCol)
        Write-Output "Active Pane: $($pane.CodeModule.Name) at Line $startLine, Col $startCol to Line $endLine, Col $endCol"
        Write-Output "Code at Line: $($pane.CodeModule.Lines($startLine, 1))"
        if ($startLine -gt 1) {
            Write-Output "Context -1: $($pane.CodeModule.Lines($startLine - 1, 1))"
            Write-Output "Context +1: $($pane.CodeModule.Lines($startLine + 1, 1))"
        }
    }
    
    $doc.Close(0)
}
catch {
    Write-Error $_
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    if (Test-Path $bakWord) { Move-Item $bakWord $startupWord -Force }
}
