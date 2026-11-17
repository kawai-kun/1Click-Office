$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open('G:\1Click Build\Output\1Click_Excel.xlam')

# Iterate VBE menu items to find Compile
$menuBar = $excel.VBE.CommandBars.Item("Menu Bar")
$debugMenu = $menuBar.Controls.Item("Debug")
Write-Output "Debug Menu Controls:"
foreach ($ctrl in $debugMenu.Controls) {
    Write-Output "Control: $($ctrl.Caption) (Id: $($ctrl.Id), Enabled: $($ctrl.Enabled))"
    if ($ctrl.Caption -like "*Compile*") {
        Write-Output "Triggering: $($ctrl.Caption)"
        $ctrl.Execute()
        Write-Output "Compile executed."
    }
}

$wb.Close($false)
$excel.Quit()
