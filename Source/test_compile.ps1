$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open('G:\1Click Build\Output\1Click_Excel.xlam')

try {
    # Find Compile command in VBE
    $cmd = $excel.VBE.CommandBars.FindControl($null, 578)
    if ($cmd -ne $null) {
        Write-Output "Executing Compile: $($cmd.Caption)"
        $cmd.Execute()
        Write-Output "Compile executed without throwing error."
    } else {
        Write-Output "Compile command not found by ID 578"
    }
}
catch {
    Write-Error "Compile Error: $_"
}
finally {
    $wb.Close($false)
    $excel.Quit()
}
