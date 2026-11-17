$dotmPath = "G:\1Click Build\Output\1Click.dotm"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.AutomationSecurity = 1
$word.DisplayAlerts = 0

Write-Output "Opening document..."
$doc = $word.Documents.Open($dotmPath, $false, $false, $false)
Write-Output "Document opened: $($doc.Name)"
Write-Output "Lines: $($doc.VBProject.VBComponents.Item('Module1').CodeModule.CountOfLines)"
$doc.Close(0)
$word.Quit()
Write-Output "Done."
