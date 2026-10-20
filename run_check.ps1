
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open('C:\Users\HP\.gemini\antigravity-ide\scratch\flipkart-bulk-lister\user_downloaded_exact.xls')
    Write-Host 'EXCEL SUCCESS! Sheets count:' $wb.Sheets.Count
    $wb.Close($false)
} catch {
    Write-Host 'EXCEL FAILURE:' $_.Exception.Message
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
