$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open('G:\1Click Build\Output\1Click_Excel.xlam')
foreach ($comp in $wb.VBProject.VBComponents) {
    Write-Output ($comp.Name + ' (Type: ' + $comp.Type + ', Lines: ' + $comp.CodeModule.CountOfLines + ')')
}
$wb.Close($false)
$excel.Quit()
