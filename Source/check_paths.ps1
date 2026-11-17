Get-ChildItem "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations" | ForEach-Object {
    Get-ItemProperty $_.PSPath | Select-Object PSChildName, Path, Description
}
