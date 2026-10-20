$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $true
try {
    $fullPath = (Resolve-Path "user_downloaded_exact.xls").Path
    $wb = $excel.Workbooks.Open($fullPath)
    Write-Host "Opened! Sheets:" $wb.Sheets.Count "HasVB:" $wb.HasVBProject
    $wb.Close($false)
} catch {
    Write-Host "CATCH: " $_.Exception.Message
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
