import os
import win32com.client
from backend.gemini_engine import generate_heuristic_batch

def test_full_mirror():
    src_path = os.path.abspath(r'C:\Users\HP\Downloads\C_chopper_4939e43e7e634736_1710-1409FK_REQX62FYHKI6J.xls')
    out_path = os.path.abspath('test_full_mirror.xls')
    if os.path.exists(out_path): os.remove(out_path)

    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(src_path)
        ws = wb.Sheets('chopper')
        
        # In Excel COM, rows and cols are 1-indexed.
        # Find headers in row 1
        max_col = ws.UsedRange.Columns.Count
        max_row = ws.UsedRange.Rows.Count
        print(f"UsedRange: {max_row} rows, {max_col} cols")

        headers = {}
        sku_col = None
        img_col = None
        other_img_cols = []
        model_name_col = None
        keywords_col = None

        for c in range(1, max_col + 1):
            val = ws.Cells(1, c).Value
            if val:
                h = str(val).strip()
                headers[c] = h
                hl = h.lower()
                if 'seller sku' in hl: sku_col = c
                elif 'main image' in hl: img_col = c
                elif 'other image' in hl: other_img_cols.append(c)
                elif hl == 'model name': model_name_col = c
                elif 'search keyword' in hl: keywords_col = c

        print(f"Mapped: SKU={sku_col}, Img={img_col}, Model={model_name_col}, Keywords={keywords_col}")

        # Check Seed row: Row 5 (0-indexed 4) vs Row 4 (0-indexed 3)
        # If Row 4 has 'Seller SKU ID is the identification', then seed row is Row 5
        seed_row = 4
        r4_sku = str(ws.Cells(4, sku_col).Value or '')
        r5_sku = str(ws.Cells(5, sku_col).Value or '')
        if 'identification number' in r4_sku.lower() or 'filled by flipkart' in r4_sku.lower():
            seed_row = 5
        print(f"Detected seed row: Excel Row {seed_row} (SKU: '{ws.Cells(seed_row, sku_col).Value}')")

        # Read all populated values from seed row
        seed_values = {}
        for c in range(1, max_col + 1):
            val = ws.Cells(seed_row, c).Value
            if val is not None and str(val).strip() != '':
                seed_values[c] = val

        print(f"Seed row has {len(seed_values)} populated attributes.")

        # Find active subsequent rows
        active_rows = []
        for r in range(seed_row + 1, max_row + 1):
            sku_val = ws.Cells(r, sku_col).Value
            img_val = ws.Cells(r, img_col).Value
            has_sku = sku_val and str(sku_val).strip() != '' and 'identification number' not in str(sku_val).lower()
            has_img = img_val and str(img_val).strip() != '' and 'image details' not in str(img_val).lower()
            if has_sku or has_img:
                active_rows.append(r)

        print(f"Active subsequent rows to process: {active_rows}")

        # Generate AI listings
        gen_data = generate_heuristic_batch("Manual Chopper", "650ml Chopper C00", len(active_rows))

        # Mirror seed values and write AI fields
        for idx, r in enumerate(active_rows):
            ai_item = gen_data[idx]
            # 1. Mirror standard attributes
            for c, s_val in seed_values.items():
                if c in [sku_col, img_col, model_name_col, keywords_col] or c in other_img_cols:
                    continue
                ws.Cells(r, c).Value = s_val

            # 2. Write unique Model Name (<80 chars)
            ws.Cells(r, model_name_col).Value = ai_item["model_name"]
            # 3. Write Search Keywords (::)
            ws.Cells(r, keywords_col).Value = ai_item["search_keywords"]

        wb.SaveCopyAs(out_path)
        wb.Close(False)
        print("SaveCopyAs completed successfully!")

    finally:
        excel.Quit()

    # Re-open in Excel to verify
    excel2 = win32com.client.DispatchEx('Excel.Application')
    excel2.Visible = False
    excel2.DisplayAlerts = False
    try:
        wb2 = excel2.Workbooks.Open(out_path)
        print(f"Re-opened! Sheets: {wb2.Sheets.Count}, HasVBProject: {wb2.HasVBProject}")
        ws2 = wb2.Sheets('chopper')
        for r in active_rows:
            print(f"Row {r}: SKU='{ws2.Cells(r, sku_col).Value}', Brand='{ws2.Cells(r, 33).Value}', Price='{ws2.Cells(r, 12).Value}', Model='{ws2.Cells(r, model_name_col).Value}', Keywords='{ws2.Cells(r, keywords_col).Value}'")
        wb2.Close(False)
    finally:
        excel2.Quit()

if __name__ == '__main__':
    test_full_mirror()
