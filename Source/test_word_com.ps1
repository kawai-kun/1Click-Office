$word = New-Object -ComObject Word.Application
$word.Visible = $false
Write-Output "Word Version: $($word.Version)"
$word.Quit()
