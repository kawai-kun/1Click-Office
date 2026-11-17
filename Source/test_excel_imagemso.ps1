$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

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
    "FilePrintQuick",
    "Numbering",
    "PageNumber",
    "PageNumberTop",
    "HeaderFooterCurrentPageNumber",
    "HeaderFooterTotalPages",
    "TableOfContentsGallery",
    "PageBreaksInsertMenu",
    "PageBreakInsert"
)

foreach ($c in $candidates) {
    try {
        $img = $excel.CommandBars.GetImageMso($c, 32, 32)
        if ($img -ne $null) {
            Write-Output "VALID: $c"
        }
    }
    catch {
        # invalid
    }
}

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
