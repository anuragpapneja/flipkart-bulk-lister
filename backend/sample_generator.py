"""
Flipkart Sample Listing Template Generator.
Creates a valid Flipkart Bulk Listing workbook (.xlsx) with:
- Rows 1-3: Flipkart Category and Technical Headers
- Row 4 (Index 3): Fully populated Master Seed Row
- Rows 5-10: Subsequent active rows with SKUs and Image URLs, awaiting automation
- 'Valid Values' and 'Instructions' sheets to mimic genuine Flipkart template structure.
"""

import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_sample_flipkart_workbook() -> bytes:
    wb = openpyxl.Workbook()
    
    # 1. Main Sheet: "Listing Template"
    ws = wb.active
    ws.title = "Listing Template"
    ws.views.sheetView[0].showGridLines = True

    # Styling
    font_cat = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_cat = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")

    font_header = Font(name="Calibri", size=10, bold=True, color="000000")
    fill_mandatory = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_system = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")

    font_master = Font(name="Calibri", size=10, bold=True, color="004D40")
    fill_master = PatternFill(start_color="E8F8F5", end_color="E8F8F5", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    columns = [
        # (Display Label, Technical Key, Is Mandatory)
        ("Seller SKU ID", "seller_sku_id", True),
        ("Brand", "brand", True),
        ("Model Name", "model_name", True),
        ("MRP", "mrp", True),
        ("Your Selling Price", "selling_price", True),
        ("Search Keywords", "search_keywords", True),
        ("Main Image URL", "main_image_url", True),
        ("Other Image URL 1", "other_image_url_1", False),
        ("Other Image URL 2", "other_image_url_2", False),
        ("Country of Origin", "country_of_origin", True),
        ("Pack of", "pack_of", True),
        ("Body Material", "body_material", True),
        ("Blade Material", "blade_material", True),
        ("HSN Code", "hsn_code", True),
        ("Tax Code", "tax_code", True),
        ("Warranty Summary", "warranty_summary", False),
        ("Product Weight (g)", "weight_g", False),
        ("Package Dimensions (L*B*H cm)", "package_dimensions", False)
    ]

    # Row 1: Category Info Header
    ws.row_dimensions[1].height = 24
    for col_idx in range(1, len(columns) + 1):
        cell = ws.cell(row=1, column=col_idx, value="Category: Home & Kitchen > Kitchen Tools > Vegetable Choppers")
        cell.font = font_cat
        cell.fill = fill_cat
        cell.alignment = Alignment(vertical="center")

    # Row 2: User Field Labels
    ws.row_dimensions[2].height = 28
    for col_idx, (label, _, is_mand) in enumerate(columns, start=1):
        cell = ws.cell(row=2, column=col_idx, value=f"{label}*" if is_mand else label)
        cell.font = font_header
        cell.fill = fill_mandatory if is_mand else fill_system
        cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
        cell.border = thin_border

    # Row 3: Technical Attribute Names
    ws.row_dimensions[3].height = 20
    for col_idx, (_, tech, _) in enumerate(columns, start=1):
        cell = ws.cell(row=3, column=col_idx, value=tech)
        cell.font = Font(name="Calibri", size=9, italic=True, color="595959")
        cell.fill = fill_system
        cell.alignment = Alignment(vertical="center", horizontal="center")
        cell.border = thin_border

    # Row 4 (Excel Row 4 / Index 3): MASTER SEED ROW
    ws.row_dimensions[4].height = 22
    master_values = [
        "SEED-CHOPPER-MASTER",  # Seller SKU ID
        "ChefNova",             # Brand
        "650ml Quick Push Press Chopper C00",  # Model Name
        599,                    # MRP
        299,                    # Selling Price
        "vegetable chopper::hand cutter::onion chopper::kitchen dicer::manual push slicer", # Keywords
        "https://rukminim1.flixcart.com/image/832/832/chopper-seed-main.jpg", # Main Image URL
        "https://rukminim1.flixcart.com/image/832/832/chopper-seed-side.jpg", # Other Image 1
        "https://rukminim1.flixcart.com/image/832/832/chopper-seed-box.jpg",  # Other Image 2
        "India",                # Country of Origin
        1,                      # Pack of
        "Virgin ABS Plastic",   # Body Material
        "Stainless Steel",      # Blade Material
        "82055100",             # HSN Code
        "GST_18",               # Tax Code
        "6 Months Domestic Manufacturer Warranty", # Warranty
        280,                    # Weight
        "15 x 12 x 10"          # Dimensions
    ]

    for col_idx, val in enumerate(master_values, start=1):
        cell = ws.cell(row=4, column=col_idx, value=val)
        cell.font = font_master
        cell.fill = fill_master
        cell.alignment = Alignment(vertical="center")
        cell.border = thin_border

    # Rows 5 - 10: Subsequent rows with only SKU and Image URLs entered
    subsequent_samples = [
        ("SKU-CHOP-RED-01", "https://rukminim1.flixcart.com/image/832/832/chopper-red-01.jpg", "https://rukminim1.flixcart.com/image/832/832/chopper-red-side.jpg"),
        ("SKU-CHOP-BLU-02", "https://rukminim1.flixcart.com/image/832/832/chopper-blue-02.jpg", "https://rukminim1.flixcart.com/image/832/832/chopper-blue-side.jpg"),
        ("SKU-CHOP-GRN-03", "https://rukminim1.flixcart.com/image/832/832/chopper-green-03.jpg", ""),
        ("SKU-CHOP-BLK-04", "https://rukminim1.flixcart.com/image/832/832/chopper-black-04.jpg", "https://rukminim1.flixcart.com/image/832/832/chopper-black-side.jpg"),
        ("SKU-CHOP-TEAL-05", "https://rukminim1.flixcart.com/image/832/832/chopper-teal-05.jpg", ""),
        ("SKU-CHOP-GREY-06", "https://rukminim1.flixcart.com/image/832/832/chopper-grey-06.jpg", "https://rukminim1.flixcart.com/image/832/832/chopper-grey-side.jpg")
    ]

    for r_idx, (sku, img, other_img) in enumerate(subsequent_samples, start=5):
        ws.row_dimensions[r_idx].height = 20
        # Seller SKU ID (Col 1)
        c_sku = ws.cell(row=r_idx, column=1, value=sku)
        c_sku.border = thin_border
        # Main Image URL (Col 7)
        c_img = ws.cell(row=r_idx, column=7, value=img)
        c_img.border = thin_border
        # Other Image URL (Col 8)
        if other_img:
            c_other = ws.cell(row=r_idx, column=8, value=other_img)
            c_other.border = thin_border
        # Fill rest with empty bordered cells
        for c_idx in range(1, len(columns) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 15), 35)

    # 2. Sheet 2: "Valid Values" (Replicating Flipkart drop-downs metadata)
    ws_valid = wb.create_sheet(title="Valid Values")
    ws_valid.cell(row=1, column=1, value="Attribute Name").font = Font(bold=True)
    ws_valid.cell(row=1, column=2, value="Allowed Values").font = Font(bold=True)
    valid_entries = [
        ("country_of_origin", "India, China, Vietnam, Thailand, Malaysia"),
        ("tax_code", "GST_0, GST_5, GST_12, GST_18, GST_28"),
        ("body_material", "Virgin ABS Plastic, Plastic, Stainless Steel, Glass"),
        ("blade_material", "Stainless Steel, High Carbon Steel, Titanium")
    ]
    for idx, (attr, vals) in enumerate(valid_entries, start=2):
        ws_valid.cell(row=idx, column=1, value=attr)
        ws_valid.cell(row=idx, column=2, value=vals)

    # 3. Sheet 3: "Instructions"
    ws_inst = wb.create_sheet(title="Instructions")
    ws_inst.cell(row=1, column=1, value="Flipkart Fast Validate Instructions:").font = Font(bold=True, size=12)
    ws_inst.cell(row=2, column=1, value="1. Press CTRL+SHIFT+S in Microsoft Excel to run Flipkart Fast Validate macro.")
    ws_inst.cell(row=3, column=1, value="2. Model Name must be unique and strictly under 80 characters.")
    ws_inst.cell(row=4, column=1, value="3. Search Keywords must contain 5 keywords separated by '::'.")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
