$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $fullPath = (Resolve-Path "test_xlutils_out.xls").Path
    $wb = $excel.Workbooks.Open($fullPath)
    Write-Host "Excel opened test_xlutils_out.xls successfully! Sheets count: " $wb.Sheets.Count
    for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
        $sName = $wb.Sheets.Item($i).Name
        Write-Host "  Sheet: " $sName
    }
    $ws = $wb.Sheets.Item("chopper")
    $cellVal = $ws.Cells.Item(6, 47).Text
    Write-Host "Row 6 (index 5) Model Name in Excel: " $cellVal
    $wb.Close($false)
} catch {
    Write-Host "Excel open error: " $_.Exception.Message
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
