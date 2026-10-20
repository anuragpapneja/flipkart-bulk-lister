import xlrd

path = r'C:\Users\HP\Downloads\C_chopper_4939e43e7e634736_1710-1350FK_REQQJLV9VJO6Q.xls'
rb = xlrd.open_workbook(path, formatting_info=True)
sheet = rb.sheet_by_name('chopper')

print("Total rows:", sheet.nrows, "Total cols:", sheet.ncols)

# Let's find columns
headers_row0 = [sheet.cell_value(0, c) for c in range(sheet.ncols)]
for c, h in enumerate(headers_row0):
    if h and str(h).strip():
        # print non-empty header
        print(f"Col {c:3d}: '{h}'")
