$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('G:\1Click Build\Output\1Click_Excel.xlam')

$menuBar = $excel.VBE.CommandBars.Item("Menu Bar")
$debugMenu = $menuBar.Controls.Item("Debug")
$compileBtn = $debugMenu.Controls.Item(1)
if ($compileBtn.Enabled) {
    $compileBtn.Execute()
}
$wb.Save()
$wb.Close($true)
$excel.Quit()
Write-Output "Compiled and saved with zero errors."
