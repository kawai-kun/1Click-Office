$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open('G:\1Click Build\Output\1Click_Excel.xlam')

$menuBar = $excel.VBE.CommandBars.Item("Menu Bar")
$debugMenu = $menuBar.Controls.Item("Debug")
$compileBtn = $debugMenu.Controls.Item(1)
Write-Output "Compile Button: $($compileBtn.Caption)"
$compileBtn.Execute()
Write-Output "After compile, button enabled: $($compileBtn.Enabled)"

$wb.Close($false)
$excel.Quit()
