$regSec = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
Set-ItemProperty -Path $regSec -Name "VBAWarnings" -Value 1 -Type DWord
Set-ItemProperty -Path $regSec -Name "AccessVBOM" -Value 1 -Type DWord

$regTrusted = "$regSec\Trusted Locations"
$locBuild = "$regTrusted\1ClickBuild"
if (-not (Test-Path $locBuild)) {
    New-Item -Path $locBuild -Force | Out-Null
}
Set-ItemProperty -Path $locBuild -Name "Path" -Value "G:\1Click Build"
Set-ItemProperty -Path $locBuild -Name "AllowSubfolders" -Value 1 -Type DWord
Set-ItemProperty -Path $locBuild -Name "Description" -Value "1Click Build Directory"

# Fix VPDU_ExcelStartup
$badKey = "$regTrusted\VPDU_ExcelStartup"
if (Test-Path $badKey) {
    Remove-Item -Path $badKey -Force -Recurse
}

Write-Output "Excel Security & Trusted Locations updated successfully."
