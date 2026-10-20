import requests
import os
import subprocess
import xlrd

url = 'http://127.0.0.1:8000/api/download/633e42cf-0c9f-484a-a8c7-43ee67b96ec1'
r = requests.get(url)
print('Status code:', r.status_code)
print('Headers:', r.headers)

out_path = os.path.abspath('user_downloaded_exact.xls')
with open(out_path, 'wb') as f:
    f.write(r.content)
print('Saved file to:', out_path, 'Length:', len(r.content))

# Inspect with xlrd
try:
    rb = xlrd.open_workbook(out_path, formatting_info=True)
    print('xlrd open OK! Sheets:', rb.sheet_names())
except Exception as e:
    print('xlrd error:', e)

# Test in Excel via powershell script
ps_script = f"""
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
    $wb = $excel.Workbooks.Open('{out_path}')
    Write-Host 'EXCEL SUCCESS! Sheets count:' $wb.Sheets.Count
    $wb.Close($false)
}} catch {{
    Write-Host 'EXCEL FAILURE:' $_.Exception.Message
}} finally {{
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
"""
with open('run_check.ps1', 'w') as f:
    f.write(ps_script)

proc = subprocess.run(['powershell', '-ExecutionPolicy', 'Bypass', '-File', 'run_check.ps1'], capture_output=True, text=True)
print('Excel Output:', proc.stdout.strip())
print('Excel Err:', proc.stderr.strip())
