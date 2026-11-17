$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open('G:\1Click Build\Output\1Click_Excel.xlam')
$cm = $wb.VBProject.VBComponents.Item('Module1').CodeModule
Write-Output "Total lines: $($cm.CountOfLines)"
$sl = 1; $sc = 1; $el = -1; $ec = -1
$found = $cm.Find("Xl_DanhSTT_LaMa", [ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
Write-Output "Found Xl_DanhSTT_LaMa: $found at line $sl"

$sl2 = 1; $sc2 = 1; $el2 = -1; $ec2 = -1
$found2 = $cm.Find("MsgBoxUni", [ref]$sl2, [ref]$sc2, [ref]$el2, [ref]$ec2)
Write-Output "Found MsgBoxUni: $found2 at line $sl2"

$wb.Close($false)
$excel.Quit()
