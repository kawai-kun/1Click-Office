$dotmPath = "G:\1Click Build\Output\1Click.dotm"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.WordBasic.DisableAutoMacros(1)
$word.DisplayAlerts = 0

Write-Output "Opening document with auto macros disabled..."
$doc = $word.Documents.Open($dotmPath)
Write-Output "Document opened: $($doc.Name)"
Write-Output "Lines: $($doc.VBProject.VBComponents.Item('Module1').CodeModule.CountOfLines)"
$doc.Close(0)
$word.Quit()
Write-Output "Word closed cleanly."
