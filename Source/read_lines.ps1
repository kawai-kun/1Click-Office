$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open('G:\1Click Build\Output\1Click_Excel.xlam')
$cm = $wb.VBProject.VBComponents.Item('Module1').CodeModule
Write-Output "Total lines: $($cm.CountOfLines)"
Write-Output "Lines 1 to 20:"
Write-Output $cm.Lines(1, 20)
Write-Output "Last 30 lines:"
Write-Output $cm.Lines($cm.CountOfLines - 30, 30)
$wb.Close($false)
$excel.Quit()
