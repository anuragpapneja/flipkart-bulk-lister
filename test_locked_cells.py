import win32com.client
import os

excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False

src = os.path.abspath(r'C:\Users\HP\Downloads\C_chopper_4939e43e7e634736_1710-1409FK_REQX62FYHKI6J.xls')
wb = excel.Workbooks.Open(src)
ws = wb.Sheets('chopper')

print("ProtectContents:", ws.ProtectContents)

# Check cell Locked property for first 15 columns
for c in range(1, 16):
    h = ws.Cells(1, c).Value
    locked = ws.Cells(5, c).Locked
    print(f"Col {c:2d} ({h}): Locked = {locked}")

# Test Unprotect
try:
    ws.Unprotect("")
    print("Unprotect with empty string succeeded! ProtectContents:", ws.ProtectContents)
except Exception as e:
    print("Unprotect empty failed:", e)

try:
    ws.Unprotect()
    print("Unprotect without args succeeded! ProtectContents:", ws.ProtectContents)
except Exception as e:
    print("Unprotect no args failed:", e)

wb.Close(False)
excel.Quit()
