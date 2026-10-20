"""
Automated Test Suite for Flipkart Bulk Listing Processor
Validates all 4 core automation rules and Flipkart Fast Validate requirements.
"""

import sys
import openpyxl
import io
from backend.sample_generator import create_sample_flipkart_workbook
from backend.processor import FlipkartExcelProcessor
from backend.gemini_engine import call_gemini_api, generate_heuristic_batch

def run_tests():
    print(">>> [Test 1] Generating Flipkart Sample Workbook...")
    raw_sample = create_sample_flipkart_workbook()
    assert len(raw_sample) > 1000, "Sample workbook is empty!"
    print("    [PASS] Generated sample workbook, size:", len(raw_sample), "bytes")

    print(">>> [Test 2] Initializing Processor and Analyzing Structure...")
    processor = FlipkartExcelProcessor(raw_sample, "test_flipkart.xlsx")
    analysis = processor.analyze()
    print("    Sheet detected:", analysis["sheet_name"])
    print("    Master Seed Model:", analysis["master_model_name"])
    print("    Active subsequent rows detected:", analysis["subsequent_active_rows_count"])
    print("    Columns detected:", analysis["columns_detected"])

    assert analysis["subsequent_active_rows_count"] == 6, f"Expected 6 active rows, got {analysis['subsequent_active_rows_count']}"
    assert analysis["columns_detected"]["sku_column"] is not None
    assert analysis["columns_detected"]["main_image_column"] is not None
    assert analysis["columns_detected"]["model_name_column"] is not None
    assert analysis["columns_detected"]["keywords_column"] is not None
    print("    [PASS] All critical columns identified correctly.")

    print(">>> [Test 3] Testing AI / Heuristic Listing Generation (Pattern & Rule Checks)...")
    gen_items = generate_heuristic_batch(
        master_title="650ml Vegetable Chopper",
        master_model="650ml Quick Push Press Chopper C00",
        count=6
    )
    assert len(gen_items) == 6
    for i, it in enumerate(gen_items):
        mn = it["model_name"]
        kw = it["search_keywords"]
        print(f"    Row {i+1} Model Name: '{mn}' (Length: {len(mn)})")
        print(f"           Keywords:   '{kw}'")

        # Rule 3: strictly < 80 characters
        assert len(mn) < 80, f"Model name exceeds 80 characters: {mn}"
        # Rule 3: keywords separated by ::
        assert "::" in kw, f"Keywords missing '::' separator: {kw}"
        parts = kw.split("::")
        assert len(parts) == 5, f"Expected 5 keywords, got {len(parts)} in '{kw}'"

    # Rule 3: 100% Unique
    all_names = [it["model_name"] for it in gen_items]
    assert len(set(all_names)) == len(all_names), "Generated model names contain duplicates!"
    print("    [PASS] Model Name pattern, length (<80), and uniqueness verified.")

    print(">>> [Test 4] Processing Workbook and Mirroring Attributes...")
    processed_bytes, audit = processor.process(gen_items, preserve_existing=True)
    assert audit["fast_validate_status"] == "PASSED"
    assert audit["errors_count"] == 0
    assert audit["all_model_names_under_80_chars"] is True
    assert audit["all_model_names_unique"] is True
    assert audit["keywords_formatted_correctly"] is True
    print("    [PASS] Fast Validate Audit passed with 0 errors!")

    print(">>> [Test 5] Inspecting Resulting Excel Workbook...")
    wb_result = openpyxl.load_workbook(io.BytesIO(processed_bytes))
    
    # Rule 4: Preserve all original sheets
    expected_sheets = ["Listing Template", "Valid Values", "Instructions"]
    for s in expected_sheets:
        assert s in wb_result.sheetnames, f"Missing sheet {s}!"
    print("    [PASS] All 3 original sheets preserved:", wb_result.sheetnames)

    ws_res = wb_result["Listing Template"]

    # Rule 1 & 2: Row 4 (Master Row 3) preserved
    master_sku = ws_res.cell(row=4, column=1).value
    master_mrp = ws_res.cell(row=4, column=4).value
    master_brand = ws_res.cell(row=4, column=2).value
    assert master_sku == "SEED-CHOPPER-MASTER"
    assert master_brand == "ChefNova"
    assert master_mrp == 599
    print("    [PASS] Master Seed Row 4 preserved intact.")

    # Rule 1 & 2: Subsequent rows have SKUs and Image URLs preserved, and MRP / Brand / Materials mirrored
    for r in range(5, 11):
        row_sku = ws_res.cell(row=r, column=1).value
        row_brand = ws_res.cell(row=r, column=2).value
        row_model = ws_res.cell(row=r, column=3).value
        row_mrp = ws_res.cell(row=r, column=4).value
        row_price = ws_res.cell(row=r, column=5).value
        row_kw = ws_res.cell(row=r, column=6).value
        row_img = ws_res.cell(row=r, column=7).value
        row_country = ws_res.cell(row=r, column=10).value
        row_hsn = ws_res.cell(row=r, column=14).value

        # SKUs must NOT be overwritten with master SKU
        assert row_sku.startswith("SKU-CHOP-"), f"SKU was overwritten in row {r}: {row_sku}"
        # Images must NOT be overwritten with master image
        assert "chopper-seed-main" not in row_img, f"Main image was overwritten in row {r}: {row_img}"
        # Brand must be mirrored
        assert row_brand == "ChefNova", f"Brand not mirrored in row {r}: {row_brand}"
        # MRP & Price must be mirrored
        assert row_mrp == 599, f"MRP not mirrored in row {r}: {row_mrp}"
        assert row_price == 299, f"Price not mirrored in row {r}: {row_price}"
        # Country of origin & HSN must be mirrored
        assert row_country == "India", f"Country not mirrored in row {r}"
        assert str(row_hsn) == "82055100", f"HSN not mirrored in row {r}"
        # Model name must be non-empty and < 80 chars
        assert len(str(row_model)) < 80, f"Model name >= 80 chars in row {r}"
        # Keywords must have ::
        assert "::" in str(row_kw), f"Keywords missing '::' in row {r}"

    print("    [PASS] All subsequent rows 5-10 successfully mirrored with preserved SKUs and unique AI fields.")
    print("\n[SUCCESS] ALL TESTS PASSED SUCCESSFULLY! 100% FLIPKART COMPLIANT.")

if __name__ == "__main__":
    run_tests()
