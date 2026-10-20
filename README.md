# Flipkart Bulk Listing Automation Tool

An intelligent full-stack automation tool designed for e-commerce sellers and catalog teams. It takes a Flipkart bulk listing sheet (`.xlsx` / `.xlsm`) where **Row 3 (1st product / Excel Row 4)** is manually seeded as the master template, and automates listing creation for all subsequent rows.

---

## ⚡ Core Rules & Automation Engine

1. **Master Row Pattern Mirroring**:
   - Reads **Row 3** (`index 3`, 4th physical Excel row) as the Master Template.
   - For every subsequent row (Row 4 onwards) containing a valid `Seller SKU ID` or `Main Image URL`, all standard attributes (MRP, Selling Price, Brand, Body Material, Blade Material, Country of Origin, Pack of, HSN, Dimensions, Tax Code, Warranty, etc.) are mirrored directly from Row 3.
   - Preserves cell types, numeric formatting, formulas, and validations.

2. **Flipkart Pre-Filled Data Preservation**:
   - `Seller SKU ID`, `Main Image URL`, and all `Other Image URLs` are **strictly preserved** and never overwritten with Row 3 values.
   - Supports a toggle to preserve custom non-blank fields entered by the user in subsequent rows.

3. **AI Generation (Gemini API) for Unique Fields ONLY**:
   - **Model Name**:
     - Auto-generated based on the master product type/feature in Row 3.
     - **Strictly < 80 characters** (enforced by a character budget validation rule).
     - **100% Unique** across all rows in the file.
     - **Zero trademarked brand names** (no Tupperware, Philips, Pigeon, etc.).
     - Follows pattern: `[Master Product Type/Feature] [Unique SEO Modifier] [Random 3-Digit Code]` (e.g., `650ml Quick Push Press Chopper C01`, `Manual Easy Press Hand Dicer C02`).
   - **Search Keywords**:
     - Exactly 5 varied, high-volume organic search keyword phrases separated by `::` (e.g., `vegetable chopper::hand cutter::onion chopper::kitchen dicer::manual push slicer`).

4. **Output Verification & Fast Validate (`CTRL+SHIFT+S`)**:
   - Retains all original sheets (`Listing Template`, `Valid Values`, `Instructions`), category headers (Row 0), user display labels (Row 1), and technical keys (Row 2).
   - In-place openpyxl modification guarantees **0 Fast Validate errors** in Microsoft Excel.

---

## 🏗️ Architecture & Project Structure

```
flipkart-bulk-lister/
├── backend/
│   ├── __init__.py
│   ├── main.py              # FastAPI server (upload, analyze, process, download)
│   ├── processor.py         # Openpyxl in-place engine, mirroring & validation audit
│   ├── gemini_engine.py     # Gemini API integration + offline intelligent generator
│   └── sample_generator.py  # Realistic seeded Flipkart Excel generator (.xlsx)
├── frontend/
│   ├── index.html           # Modern semantic HTML5 interface
│   ├── style.css            # Dark mode design system (glassmorphism & micro-animations)
│   └── app.js               # Reactive client logic, drag-and-drop, diff preview
├── run.py                   # One-click startup script (starts uvicorn & opens browser)
├── start.bat                # Windows quick launcher
├── test_engine.py           # Automated test suite
└── requirements.txt         # Dependencies
```

---

## 🚀 Quick Start Guide

### 1. Requirements
- Python 3.10+
- Modern Web Browser (Chrome, Edge, Firefox, Safari)

### 2. Installation
```powershell
# Navigate to project folder
cd C:\Users\HP\.gemini\antigravity-ide\scratch\flipkart-bulk-lister

# Activate virtual environment
.\venv\Scripts\Activate.ps1

# Install dependencies (already pre-installed)
pip install -r requirements.txt
```

### 3. Launch Application
```powershell
python run.py
```
Or double-click `start.bat` on Windows.
The application will automatically launch at **`http://127.0.0.1:8000`** in your default web browser.

---

## 🧪 Testing & Verification
To run the automated verification suite:
```powershell
python test_engine.py
```
This tests:
- Sample sheet generation
- Header and column mapping
- Rule 1: Master Row 3 attribute mirroring
- Rule 2: SKU and image URL preservation
- Rule 3: Model Name length (< 80 chars), uniqueness, and `::` keyword formatting
- Rule 4: Multi-sheet integrity and Fast Validate 0-error compliance
