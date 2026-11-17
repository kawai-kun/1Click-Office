$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Open('G:\1Click Build\Output\1Click.dotm')
$lic = $word.Run('CheckLicenseActivation')
$ver = $word.Run('GetVersion')
$cnt = $word.Run('Profile_GetCount')
Write-Output "Word 1Click: Lic=$lic, Ver=$ver, Profiles=$cnt"
$doc.Close(0)
$word.Quit()
