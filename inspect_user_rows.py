import xlrd

path = r'C:\Users\HP\Downloads\C_chopper_4939e43e7e634736_1710-1350FK_REQQJLV9VJO6Q.xls'
rb = xlrd.open_workbook(path, formatting_info=True)
sheet = rb.sheet_by_name('chopper')

for r in range(3, 10):
    sku = sheet.cell_value(r, 6)
    model = sheet.cell_value(r, 46)
    kw = sheet.cell_value(r, 66)
    mrp = sheet.cell_value(r, 10)
    brand = sheet.cell_value(r, 32)
    img = sheet.cell_value(r, 41)
    print(f"Row {r}: SKU='{sku}', Brand='{brand}', MRP='{mrp}', Model='{model}', Keywords='{kw}', Img='{str(img)[:30]}'")
