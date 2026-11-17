Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$srcZip = 'G:\1Click Build\Source\ChuanHoa_Excel_original.zip'
$outXlam = 'G:\1Click Build\Output\1Click_Excel.xlam'
Copy-Item $srcZip $outXlam -Force

# Update customUI.xml, customUI14.xml, core.xml
$customUIContent = [System.IO.File]::ReadAllBytes('G:\1Click Build\Source\Excel_customUI.xml')
$zip = [System.IO.Compression.ZipFile]::Open($outXlam, [System.IO.Compression.ZipArchiveMode]::Update)

$entry1 = $zip.GetEntry('customUI/customUI.xml')
if ($entry1) { $entry1.Delete() }
$newEntry1 = $zip.CreateEntry('customUI/customUI.xml')
$stream1 = $newEntry1.Open()
$stream1.Write($customUIContent, 0, $customUIContent.Length)
$stream1.Close()

$entry2 = $zip.GetEntry('customUI/customUI14.xml')
if ($entry2) { $entry2.Delete() }
$newEntry2 = $zip.CreateEntry('customUI/customUI14.xml')
$stream2 = $newEntry2.Open()
$stream2.Write($customUIContent, 0, $customUIContent.Length)
$stream2.Close()

# Update docProps/core.xml
$coreXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
'<dc:title>1Click Excel 3.0 Enterprise</dc:title>' +
'<dc:subject>1Click Excel - Chuan Hoa In va Xu Ly Van Ban Hanh Chinh</dc:subject>' +
'<dc:creator>Trần Quốc Được - TTDVTH Phường Tân Khai</dc:creator>' +
'<cp:lastModifiedBy>Trần Quốc Được - TTDVTH Phường Tân Khai</cp:lastModifiedBy>' +
'<dcterms:created xsi:type="dcterms:W3CDTF">2026-11-15T00:00:00Z</dcterms:created>' +
'<dcterms:modified xsi:type="dcterms:W3CDTF">2026-11-15T14:00:00Z</dcterms:modified>' +
'</cp:coreProperties>'

$coreBytes = [System.Text.Encoding]::UTF8.GetBytes($coreXml)
$entryCore = $zip.GetEntry('docProps/core.xml')
if ($entryCore) { $entryCore.Delete() }
$newCore = $zip.CreateEntry('docProps/core.xml')
$streamCore = $newCore.Open()
$streamCore.Write($coreBytes, 0, $coreBytes.Length)
$streamCore.Close()

$zip.Dispose()
Write-Output 'Zip OpenXML updated successfully.'
