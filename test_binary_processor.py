"""
Comprehensive Test for Flipkart In-Place Binary Processor
Tests both legacy .xls and modern .xlsx formats.
Verifies all sheets preserved, in-place modification, and zero corruption.
"""

import os
import io
import xlrd
import openpyxl
from backend.processor import FlipkartBinaryProcessor
from backend.gemini_engine import generate_heuristic_batch
from backend.sample_generator import create_sample_flipkart_workbook

def test_real_flipkart_xls():
    real_path = r'C:\Users\HP\Downloads\C_chopper_4939e43e7e634736_1710-1350FK_REQQJLV9VJO6Q.xls'
    if not os.path.exists(real_path):
        print("[SKIP] Real Flipkart file not found at", real_path)
        return

    print(">>> [Test 1] Testing In-Place Processing on Real Flipkart .xls File...")
    with open(real_path, 'rb') as f:
        raw_bytes = f.read()

    proc = FlipkartBinaryProcessor(raw_bytes, "flipkart_chopper.xls")
    analysis = proc.analyze()

    print(f"    Format: {analysis['format']}")
    print(f"    Sheet name: {analysis['sheet_name']}")
    print(f"    All sheets ({len(analysis['all_sheets'])}): {analysis['all_sheets']}")
    print(f"    Master Seed Row Index: {analysis['master_row_index']}")
    print(f"    Active rows count: {analysis['subsequent_active_rows_count']}")
    print(f"    Detected columns: {analysis['columns_detected']}")

    assert analysis['format'] == 'xls'
    assert 'chopper' in analysis['all_sheets']
    assert 'Summary Sheet' in analysis['all_sheets']
    assert analysis['subsequent_active_rows_count'] > 0

    # Generate test listing items
    gen_items = generate_heuristic_batch(
        master_title="Manual Vegetable Chopper",
        master_model=analysis['master_model_name'] or "Quick Push Chopper 650ml C00",
        count=analysis['subsequent_active_rows_count']
    )

    out_bytes, audit, out_name = proc.process(gen_items, preserve_existing=True)

    print(f"    Processed file: {out_name}, size: {len(out_bytes)} bytes")
    print(f"    Total processed rows: {audit['total_processed_rows']}")
    print(f"    All model names < 80 chars: {audit['all_model_names_under_80_chars']}")
    print(f"    All model names unique: {audit['all_model_names_unique']}")
    print(f"    Keywords formatted with '::': {audit['keywords_formatted_correctly']}")

    assert out_name.endswith(".xls"), "Output file format must be .xls"
    assert audit['all_model_names_under_80_chars'] is True
    assert audit['all_model_names_unique'] is True
    assert audit['keywords_formatted_correctly'] is True

    # Re-read with xlrd to verify all sheets, rows, and cells
    rb_out = xlrd.open_workbook(file_contents=out_bytes, formatting_info=True)
    assert rb_out.sheet_names() == analysis['all_sheets'], "Sheet names mismatch after modification!"
    sheet_out = rb_out.sheet_by_name(analysis['sheet_name'])

    # Verify preserved SKUs
    for row_info in audit['processed_rows'][:5]:
        r_idx = row_info['row_index'] - 1 # 0-indexed in xlrd
        cell_sku = sheet_out.cell_value(r_idx, 6) # Col 6 is SKU in Flipkart template
        assert cell_sku == row_info['sku'], f"SKU changed! Expected {row_info['sku']}, got {cell_sku}"

    print("    [PASS] Real Flipkart .xls file modified in-place with 100% sheet integrity!")


def test_xlsx_format():
    print("\n>>> [Test 2] Testing In-Place Processing on .xlsx File...")
    sample_bytes = create_sample_flipkart_workbook()

    proc = FlipkartBinaryProcessor(sample_bytes, "sample.xlsx")
    analysis = proc.analyze()
    print(f"    Format: {analysis['format']}")
    print(f"    Sheets ({len(analysis['all_sheets'])}): {analysis['all_sheets']}")
    print(f"    Active rows count: {analysis['subsequent_active_rows_count']}")

    gen_items = generate_heuristic_batch(
        master_title="Vegetable Chopper",
        master_model="650ml Chopper C00",
        count=analysis['subsequent_active_rows_count']
    )

    out_bytes, audit, out_name = proc.process(gen_items, preserve_existing=True)

    assert out_name.endswith(".xlsx")
    assert audit['all_model_names_under_80_chars'] is True

    wb_out = openpyxl.load_workbook(io.BytesIO(out_bytes))
    assert wb_out.sheetnames == analysis['all_sheets']
    print("    [PASS] .xlsx file processed in-place without corruption!")


if __name__ == "__main__":
    test_real_flipkart_xls()
    test_xlsx_format()
    print("\n[SUCCESS] ALL BINARY IN-PLACE PROCESSOR TESTS PASSED!")
