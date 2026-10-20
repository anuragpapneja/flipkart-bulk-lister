import xlrd

path = r'C:\Users\HP\Downloads\C_chopper_4939e43e7e634736_1710-1350FK_REQQJLV9VJO6Q.xls'
rb = xlrd.open_workbook(path, formatting_info=True)
sheet = rb.sheet_by_name('chopper')

for r in range(8):
    populated = [(c, sheet.cell_value(0, c), sheet.cell_value(r, c)) for c in range(sheet.ncols) if str(sheet.cell_value(r, c)).strip() != '']
    print(f"\n=== ROW {r} (populated count: {len(populated)}) ===")
    for c, h, val in populated[:15]:
        val_str = str(val).replace('\n', ' ')
        if len(val_str) > 40: val_str = val_str[:37] + '...'
        print(f"  Col {c:2d} ({h}): {val_str}")
