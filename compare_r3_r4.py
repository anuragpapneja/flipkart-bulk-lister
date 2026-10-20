import xlrd

path = r'C:\Users\HP\Downloads\C_chopper_4939e43e7e634736_1710-1350FK_REQQJLV9VJO6Q.xls'
rb = xlrd.open_workbook(path, formatting_info=True)
sheet = rb.sheet_by_name('chopper')

print("Row 3 values:")
for c in range(sheet.ncols):
    val = sheet.cell_value(3, c)
    if str(val).strip():
        print(f"  Col {c} ({sheet.cell_value(0, c)}): {repr(val)[:50]}")

print("\nRow 4 values:")
for c in range(sheet.ncols):
    val = sheet.cell_value(4, c)
    if str(val).strip():
        print(f"  Col {c} ({sheet.cell_value(0, c)}): {repr(val)[:50]}")
