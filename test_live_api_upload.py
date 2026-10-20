"""
Test uploading the real Flipkart file to the running FastAPI server.
Verifies the HTTP 200 response, binary download, and Microsoft Excel opening.
"""

import requests
import io
import xlrd
import subprocess

url_process = "http://127.0.0.1:8000/api/process"
real_file = r"C:\Users\HP\Downloads\C_chopper_4939e43e7e634736_1710-1350FK_REQQJLV9VJO6Q.xls"

with open(real_file, "rb") as f:
    files = {"file": ("flipkart_chopper.xls", f, "application/vnd.ms-excel")}
    data = {"preserve_existing": "true"}
    resp = requests.post(url_process, files=files, data=data)

assert resp.status_code == 200, f"Process failed: {resp.status_code} - {resp.text}"
res_json = resp.json()
print("Process response:", res_json["success"])
print("Audit:", res_json["audit"]["fast_validate_status"], "Total rows:", res_json["audit"]["total_processed_rows"])

file_id = res_json["file_id"]
out_filename = res_json["filename"]
print("Downloading file:", out_filename)

url_download = f"http://127.0.0.1:8000/api/download/{file_id}"
dl_resp = requests.get(url_download)
assert dl_resp.status_code == 200, f"Download failed: {dl_resp.status_code}"

with open("downloaded_test.xls", "wb") as f:
    f.write(dl_resp.content)

print("Downloaded file size:", len(dl_resp.content))

# Verify with xlrd
rb = xlrd.open_workbook("downloaded_test.xls", formatting_info=True)
print("Sheets in downloaded file:", rb.sheet_names())
assert len(rb.sheet_names()) == 10, "Missing sheets!"

# Verify in Microsoft Excel COM
cmd = 'powershell -ExecutionPolicy Bypass -Command "$excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excel.DisplayAlerts = $false; try { $wb = $excel.Workbooks.Open((Resolve-Path \'downloaded_test.xls\').Path); Write-Host \'Excel OK! Sheets:\' $wb.Sheets.Count; $wb.Close($false) } catch { Write-Host \'Excel ERR:\' $_.Exception.Message } finally { $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }"'
proc = subprocess.run(cmd, shell=True, capture_output=True, text=True)
print("Excel COM verification output:", proc.stdout.strip())
assert "Excel OK!" in proc.stdout, f"Excel failed to open: {proc.stdout}"

print("\n[SUCCESS] Live API upload, binary in-place processing, and download verified with 0 errors!")
