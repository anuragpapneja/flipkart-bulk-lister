import win32com.client
import os
import subprocess
from backend.gemini_engine import generate_heuristic_batch

def run_native_mirror():
    src_path = os.path.abspath(r'C:\Users\HP\Downloads\C_chopper_4939e43e7e634736_1710-1409FK_REQX62FYHKI6J.xls')
    out_path = os.path.abspath('flipkart_native_perfect.xls')
    if os.path.exists(out_path): os.remove(out_path)

    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.AutomationSecurity = 3 # msoAutomationSecurityForceDisable

    try:
        wb = excel.Workbooks.Open(src_path)
        ws = wb.Sheets('chopper')
        
        max_col = ws.UsedRange.Columns.Count
        max_row = ws.UsedRange.Rows.Count
        print(f"Workbook opened: {max_row} rows, {max_col} cols, HasVBProject: {wb.HasVBProject}")

        # Map columns in row 1
        sku_col = None
        img_col = None
        other_img_cols = []
        model_name_col = None
        keywords_col = None

        for c in range(1, max_col + 1):
            val = ws.Cells(1, c).Value
            if val:
                hl = str(val).strip().lower()
                if 'seller sku' in hl: sku_col = c
                elif 'main image' in hl: img_col = c
                elif 'other image' in hl: other_img_cols.append(c)
                elif hl == 'model name': model_name_col = c
                elif 'search keyword' in hl: keywords_col = c

        print(f"Mapped: SKU={sku_col}, Img={img_col}, Model={model_name_col}, Keywords={keywords_col}")

        # Determine seed row (Row 5 in Excel 1-based if Row 4 has description text)
        seed_row = 4
        r4_sku = str(ws.Cells(4, sku_col).Value or '').lower()
        if 'identification number' in r4_sku or 'filled by flipkart' in r4_sku:
            seed_row = 5
        print(f"Using Master Seed Row: {seed_row} (SKU: '{ws.Cells(seed_row, sku_col).Value}')")

        # Read seed row values for all UNLOCKED columns
        seed_values = {}
        for c in range(1, max_col + 1):
            # Check if cell is locked
            if ws.Cells(seed_row, c).Locked:
                continue # Skip protected system columns
            val = ws.Cells(seed_row, c).Value
            if val is not None and str(val).strip() != '':
                seed_values[c] = val

        print(f"Extracted {len(seed_values)} unlocked attributes from seed row.")

        # Find active subsequent rows (Row 6 onwards)
        active_rows = []
        for r in range(seed_row + 1, max_row + 1):
            sku_val = ws.Cells(r, sku_col).Value
            img_val = ws.Cells(r, img_col).Value
            has_sku = sku_val and str(sku_val).strip() != '' and 'identification number' not in str(sku_val).lower()
            has_img = img_val and str(img_val).strip() != '' and 'image details' not in str(img_val).lower()
            if has_sku or has_img:
                active_rows.append(r)

        print(f"Active rows to mirror: {active_rows}")
        gen_data = generate_heuristic_batch("Manual Push Chopper", "650ml Chopper C00", len(active_rows))

        # Mirror attributes into active rows
        for idx, r in enumerate(active_rows):
            ai_item = gen_data[idx]
            # 1. Mirror standard unlocked attributes
            for c, s_val in seed_values.items():
                if c in [sku_col, img_col, model_name_col, keywords_col] or c in other_img_cols:
                    continue
                # Write standard attribute
                ws.Cells(r, c).Value = s_val

            # 2. Write unique Model Name
            ws.Cells(r, model_name_col).Value = ai_item["model_name"]
            # 3. Write Search Keywords
            ws.Cells(r, keywords_col).Value = ai_item["search_keywords"]

        wb.SaveCopyAs(out_path)
        wb.Close(False)
        print("Successfully saved:", out_path, "Size:", os.path.getsize(out_path))

    finally:
        excel.Quit()

    # Re-verify with Excel
    excel2 = win32com.client.DispatchEx('Excel.Application')
    excel2.Visible = False
    excel2.DisplayAlerts = False
    try:
        wb2 = excel2.Workbooks.Open(out_path)
        print("Reopened saved file successfully! Sheets count:", wb2.Sheets.Count, "HasVBProject:", wb2.HasVBProject)
        ws2 = wb2.Sheets('chopper')
        for r in active_rows:
            print(f"Row {r}: SKU='{ws2.Cells(r, sku_col).Value}', Brand='{ws2.Cells(r, 33).Value}', Price='{ws2.Cells(r, 12).Value}', Model='{ws2.Cells(r, model_name_col).Value}'")
        wb2.Close(False)
    finally:
        excel2.Quit()

if __name__ == '__main__':
    run_native_mirror()
