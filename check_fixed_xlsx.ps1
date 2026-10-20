$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $fullPath = (Resolve-Path "test_fixed.xlsx").Path
    $wb = $excel.Workbooks.Open($fullPath)
    Write-Host "Excel opened test_fixed.xlsx successfully! Sheets count: " $wb.Sheets.Count
    $wb.Close($false)
} catch {
    Write-Host "Excel open error: " $_.Exception.Message
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
