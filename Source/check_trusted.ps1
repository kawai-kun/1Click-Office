Get-ChildItem "HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations" -Recurse | Select-Object Name, Property
