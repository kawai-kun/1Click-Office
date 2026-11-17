$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open('G:\1Click Build\Output\1Click_Excel.xlam')

$menuBar = $excel.VBE.CommandBars.Item("Menu Bar")
$debugMenu = $menuBar.Controls.Item("Debug")
$compileBtn = $debugMenu.Controls.Item(1)
$compileBtn.Execute()

# Check ActiveCodePane
try {
    $pane = $excel.VBE.ActiveCodePane
    if ($pane -ne $null) {
        $startLine = 0; $startCol = 0; $endLine = 0; $endCol = 0
        $pane.GetSelection([ref]$startLine, [ref]$startCol, [ref]$endLine, [ref]$endCol)
        Write-Output "Selection at Line $startLine, Col $startCol to Line $endLine, Col $endCol"
        Write-Output "Code at Line: $($pane.CodeModule.Lines($startLine, 1))"
    }
} catch {
    Write-Output "Error reading pane: $_"
}

$wb.Close($false)
$excel.Quit()
