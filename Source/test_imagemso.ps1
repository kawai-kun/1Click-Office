$word = New-Object -ComObject Word.Application
$word.Visible = $false

$candidates = @(
    "PageNumberInsert",
    "HeaderFooterPageNumberInsert",
    "HeaderFooterPageNumber",
    "PageNumberFormatDialog",
    "PageNumberRemove",
    "PageNumberMarginsGallery",
    "PageNumberCurrentPositionGallery",
    "PageNumberBottomGallery",
    "PageNumberTopGallery",
    "PageSetupPageDialog",
    "NumberingGallery",
    "NumberingInsert",
    "FilePrintQuick"
)

foreach ($c in $candidates) {
    try {
        $img = $word.CommandBars.GetImageMso($c, 32, 32)
        if ($img -ne $null) {
            Write-Output "VALID: $c ($($img.Type))"
        } else {
            Write-Output "NULL: $c"
        }
    }
    catch {
        Write-Output "INVALID: $c"
    }
}

$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
