$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open('G:\1Click Build\Output\1Click_Excel.xlam')

try {
    Write-Output "Testing ChuanHoa_TrangIn_Excel_Silent..."
    $excel.Run("1Click_Excel.xlam!ChuanHoa_TrangIn_Excel_Silent")
    Write-Output "Ran ChuanHoa_TrangIn_Excel_Silent successfully."
}
catch {
    Write-Error $_
}
finally {
    $wb.Close($false)
    $excel.Quit()
}
