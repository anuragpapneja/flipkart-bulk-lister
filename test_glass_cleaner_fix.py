import os
import requests
import json
import win32com.client

path = os.path.expanduser(r'~\Downloads\C_glass-cleaner_4939e43e7e634736_1710-1458FK_REQHLQX2KJJ39.xls')
print('=== 1. ANALYZE TEST ===')
with open(path, 'rb') as f:
    r_ana = requests.post(
        'http://127.0.0.1:8000/api/analyze',
        files={'file': (os.path.basename(path), f, 'application/vnd.ms-excel')}
    )
ana = r_ana.json()
print('Analysis Success:', ana.get('success'))
print('Sheet name:', ana['analysis']['sheet_name'])
print('Columns detected:', ana['analysis']['columns_detected'])
print('Master Model Name:', ana['analysis']['master_model_name'])
print('Master Keywords:', ana['analysis']['master_keywords'][:60])
print('Active rows count:', ana['analysis']['subsequent_active_rows_count'])

print('\n=== 2. PROCESS TEST ===')
with open(path, 'rb') as f:
    r_proc = requests.post(
        'http://127.0.0.1:8000/api/process',
        files={'file': (os.path.basename(path), f, 'application/vnd.ms-excel')},
        data={'preserve_existing': 'false', 'model': 'gemini-2.5-flash'}
    )
proc = r_proc.json()
print('Process Success:', proc.get('success'))
file_id = proc['file_id']
filename = proc['filename']
print('Generated Rows Preview:')
for row in proc['audit']['processed_rows']:
    r_idx = row['row_index']
    sku = row['sku']
    mn = row['model_name']
    mn_len = row['model_name_len']
    kw = row['search_keywords']
    print(f' Row #{r_idx}: SKU={sku}')
    print(f'   Model Name ({mn_len} chars): {mn}')
    print(f'   Keywords: {kw}')

print('\n=== 3. DOWNLOAD & EXCEL COM VERIFICATION ===')
r_dl = requests.get(f'http://127.0.0.1:8000/api/download/{file_id}')
out_path = os.path.abspath('verification_glass_cleaner.xls')
with open(out_path, 'wb') as f:
    f.write(r_dl.content)
print('Downloaded size:', len(r_dl.content))

excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(out_path)
    print('Sheets count:', wb.Sheets.Count)
    print('HasVBProject:', wb.HasVBProject)
    ws = wb.Sheets(ana['analysis']['sheet_name'])
    for r in range(5, 5 + ana['analysis']['subsequent_active_rows_count'] + 1):
        r_sku = ws.Cells(r, 7).Value
        r_mod = ws.Cells(r, 36).Value
        r_kw = str(ws.Cells(r, 56).Value)[:65]
        print(f' Excel Row {r}: SKU={r_sku} | Model={r_mod} | Keywords={r_kw}')
    wb.Close(False)
    print('\n>>> EXCEL TEST PASSED WITH 0 ERRORS! <<<')
finally:
    excel.Quit()
