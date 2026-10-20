$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $fullPath = (Resolve-Path "test_processed.xlsx").Path
    $wb = $excel.Workbooks.Open($fullPath)
    Write-Host "Excel opened test_processed.xlsx successfully! Sheets: " $wb.Sheets.Count
    $ws = $wb.Sheets.Item(1)
    Write-Host "Row 5 SKU: " $ws.Cells.Item(5, 1).Text
    Write-Host "Row 5 Model: " $ws.Cells.Item(5, 3).Text
    Write-Host "Row 5 Keywords: " $ws.Cells.Item(5, 6).Text
    Write-Host "Row 5 Brand: " $ws.Cells.Item(5, 2).Text
    Write-Host "Row 5 Price: " $ws.Cells.Item(5, 5).Text
    $wb.Close($false)
} catch {
    Write-Host "Excel open error: " $_.Exception.Message
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
