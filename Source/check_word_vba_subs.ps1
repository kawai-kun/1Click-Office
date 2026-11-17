$dotm = 'G:\1Click Build\Output\1Click.dotm'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.WordBasic.DisableAutoMacros(1)
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($dotm)
    $cm = $doc.VBProject.VBComponents.Item('Module1').CodeModule
    Write-Output "Module1 line count: $($cm.CountOfLines)"
    
    $sl = 1; $sc = 1; $el = -1; $ec = -1
    $f1 = $cm.Find("RibbonThongTinTacGia", [ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
    Write-Output "Found RibbonThongTinTacGia: $f1 at line $sl"
    
    $sl = 1; $sc = 1; $el = -1; $ec = -1
    $f2 = $cm.Find("MsgBoxUni", [ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
    Write-Output "Found MsgBoxUni: $f2 at line $sl"
    
    $doc.Close(0)
} finally {
    $word.Quit()
}
