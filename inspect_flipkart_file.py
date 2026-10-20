import xlrd

path = r'C:\Users\HP\Downloads\C_chopper_4939e43e7e634736_1710-1350FK_REQQJLV9VJO6Q.xls'
rb = xlrd.open_workbook(path, formatting_info=True)
print("Sheet names in Flipkart template:", rb.sheet_names())

for sname in rb.sheet_names():
    sheet = rb.sheet_by_name(sname)
    print(f"\n--- Sheet: '{sname}' ({sheet.nrows} rows, {sheet.ncols} cols) ---")
    if sheet.nrows > 0:
        for r in range(min(sheet.nrows, 6)):
            row_vals = [sheet.cell_value(r, c) for c in range(min(sheet.ncols, 10))]
            print(f"  Row {r}: {row_vals}")
