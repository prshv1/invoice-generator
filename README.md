<div align="center">

# 📄 Invoice Generator

**Affordable GST-compliant invoicing for India's small businesses.**

[![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-22c55e?style=for-the-badge)](LICENSE)
[![Tests](https://img.shields.io/badge/Tests-62%20passed-22c55e?style=for-the-badge)]()
[![Version](https://img.shields.io/badge/Version-0.2.0-6366f1?style=for-the-badge)]()

---

*Small businesses in India can't afford big SaaS invoicing software.*
*Upload your raw delivery data → get a digitized, GST-compliant tax invoice instantly.*

**🇮🇳 Promoting Digital Bharat — one invoice at a time.**

</div>

---

## 🎯 The Problem

Millions of small businesses across India — from building material suppliers to kirana stores — still rely on handwritten invoices. They can't afford ₹5,000–₹20,000/year SaaS subscriptions. Without proper GST invoices, they risk non-compliance with tax laws.

**Invoice Generator** bridges this gap — a free, open-source tool that turns raw delivery records into professional, GST-compliant tax invoices in seconds.

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| 📊 **Instant Generation** | Upload Excel data → download formatted invoice |
| 📄 **PDF Export** | Generate print-ready PDF invoices alongside Excel |
| 📦 **Batch Processing** | Process entire folders with auto-incrementing invoice numbers |
| ⚙️ **Config File** | Customize company, buyer, bank, GST via `config.yaml` — no code editing |
| 🧾 **GST Compliant** | Auto-calculates CGST + SGST with HSN codes |
| 🧱 **Any Material** | Dynamic — works with any material types, no hardcoding |
| 📍 **Multi-Site** | Handles deliveries across multiple sites per invoice |
| 🔢 **Per-Material Rates** | Supports different rates for each material at each site |
| 📅 **Smart Dates** | Auto-detects fiscal year, handles multiple date formats |
| 📋 **Two-Sheet Output** | Professional Bill sheet + detailed List sheet |
| 🛡️ **Robust** | 62 tests covering edge cases and failure scenarios |
| 💻 **Cross-Platform** | Works on Windows, macOS, and Linux |

---

## 🚀 Quick Start

### 1. Clone & Install

```bash
git clone https://github.com/prshv1/invoice-generator.git
cd invoice-generator
pip install -r requirements.txt
```

### 2. Prepare Your Data

Create an Excel file (`.xlsx` or `.xls`) with these columns:

| Column | Description | Example |
|--------|-------------|---------|
| `Date` | Delivery date | `16/02/2025` |
| `Challan No.` | Receipt number | `1234` |
| `Vehicle No.` | Vehicle number | `5678` |
| `Site` | Delivery site | `Malad` |
| `Material` | Material type | `10 mm` |
| `Quantity` | Qty (tonnes) | `35.61` |
| `Rate` | ₹ per tonne | `380` |
| `Per` | Unit | `Tonne` |

> 💡 Sample files included in [`Example Data/`](Example%20Data/)

### 3. Generate

```bash
# Interactive mode
python3 src/invoice_generator/generator.py

# Specify invoice number
python3 src/invoice_generator/generator.py -i 178

# Generate Excel + PDF
python3 src/invoice_generator/generator.py -i 178 --pdf

# Batch process all files in a folder
python3 src/invoice_generator/generator.py --batch ./data/ --start 178

# Batch with PDF output
python3 src/invoice_generator/generator.py --batch ./data/ --start 178 --pdf --output-dir ./invoices/
```

---

## ⚙️ Configuration

Customize `config.yaml` for your business — **no code editing required**:

```yaml
company:
  name: "YOUR COMPANY NAME"
  subtitle: "(BUSINESS TYPE)"
  address: "Your Address"
  contact: "9876543210"
  gstn: "00XXXXX0000X0XX"
  pan: "XXXXX0000X"

buyer:
  name: "BUYER NAME"
  address: "Buyer Address"
  gstn: "00XXXXX0000X0XX"

bank:
  account_name: "YOUR COMPANY"
  bank_name: "BANK NAME"
  account_no: "000000000000"
  branch: "BRANCH"
  ifsc: "XXXX0000000"

gst:
  cgst_rate: 0.09
  sgst_rate: 0.09
  hsn_code: 996511

unit: "Tonne"
```

Use a custom config file:
```bash
python3 src/invoice_generator/generator.py -i 178 --config my_config.yaml
```

---

## 🔧 Python Module API

```python
from generate_invoices import generate_invoice, generate_pdf, batch_process

# Generate Excel workbook
wb = generate_invoice("data.xlsx", inv_num=178)
wb.save("Invoice_178.xlsx")

# Generate PDF
generate_pdf("data.xlsx", inv_num=178, output_path="Invoice_178.pdf")

# Get PDF as bytes (for web apps, email attachments, etc.)
pdf_bytes = generate_pdf("data.xlsx", inv_num=178)

# Batch process with custom config
config = {"company": {"name": "My Co"}, ...}
results = batch_process("./data/", start_num=100, pdf=True, config=config)
```

---

## 🧪 Testing

Run the full test suite:

```bash
pytest tests/ -v
```

```
62 passed in 1.49s
```

Tests cover:
- ✅ Helper functions (date parsing, formatting, rounding, sanitization)
- ✅ Config loading (YAML, defaults, missing files)
- ✅ Data processing (single/multi site, rates, totals, GST)
- ✅ Failure scenarios (missing files, bad columns, non-numeric data, empty files)
- ✅ Excel generation (sheets, amounts, formulas, edge cases)
- ✅ PDF generation (output, file saving, format validation)
- ✅ Batch processing (auto-numbering, error handling, empty dirs)

---

## 📁 Project Structure

```
invoice-generator/
├── src/invoice_generator/generator.py    # Core engine (CLI + module API)
├── config.yaml             # Business configuration (customizable)
├── requirements.txt        # Dependencies
├── Example Data/
│   ├── Raw_Data.xlsx       # Full sample (3 sites, 3 materials)
│   ├── Single_Site.xlsx    # Edge case: 1 site, 1 material
│   ├── Multi_Rate.xlsx     # Edge case: different rates per material
│   ├── Minimal.xlsx        # Edge case: 1 row of data
│   └── Edge_Cases.xlsx     # Edge case: string dates, missing values
├── tests/
│   └── test_invoices.py    # 62 test cases
├── .gitignore
└── README.md
```

---

## 🛡️ Edge Cases Handled

- ✅ Multiple date formats (`dd/mm/yyyy`, `yyyy-mm-dd`, `dd-mm-yyyy`, etc.)
- ✅ Non-numeric or missing values in Quantity/Rate (clear error with row numbers)
- ✅ Different rates per material at the same site
- ✅ Materials present at some sites but not others (blank, not zero)
- ✅ Single-day deliveries (shows date, not redundant range)
- ✅ Negative quantities for returns/credits (warns, doesn't block)
- ✅ Special characters in site/material names
- ✅ Both `.xlsx` and `.xls` file formats
- ✅ Proper conventional rounding (0.5 rounds up, not banker's rounding)
- ✅ Float precision in GST calculations (rounded to 2 decimal places)
- ✅ Sheet names >31 characters (auto-truncated)
- ✅ Missing config file (falls back to defaults)

---

## 📋 CLI Reference

```
usage: src/invoice_generator/generator.py [-h] [-i INVOICE] [--input INPUT] [--output OUTPUT]
                            [--pdf] [--batch DIR] [--start START]
                            [--output-dir DIR] [--config CONFIG] [--version]

options:
  -i, --invoice    Invoice number
  --input          Input Excel file (default: Example Data/Raw_Data.xlsx)
  --output         Output Excel file (default: Final_Output.xlsx)
  --pdf            Also generate PDF output
  --batch DIR      Batch process all Excel files in directory
  --start START    Starting invoice number for batch mode (default: 1)
  --output-dir DIR Output directory for batch mode
  --config CONFIG  Path to config YAML file
  --version        Show version number
```

---

## 🛣️ Roadmap

- [x] ~~Config file for business details~~
- [x] ~~PDF export~~
- [x] ~~Batch processing with auto-numbering~~
- [x] ~~Test suite~~
- [ ] 🌐 Web interface — browser-based upload & download
- [ ] 📸 OCR support — upload photos of handwritten bills
- [ ] 📧 Email integration — auto-send invoices to buyers
- [ ] 🔄 IGST support — inter-state invoice detection

---

## 📝 Changelog

### v0.2.0 — 2025-03-02

**New Features:**
- 📄 **PDF Export** — generate print-ready PDF alongside Excel (`--pdf` flag)
- ⚙️ **Config File** — `config.yaml` for company/buyer/bank/GST customization
- 📦 **Batch Processing** — process entire folders with `--batch` and `--start`
- 🔢 **Auto Invoice Numbering** — sequential numbering in batch mode
- 🖥️ **CLI with argparse** — proper flags: `-i`, `--pdf`, `--batch`, `--config`
- 🧪 **Test Suite** — 62 pytest cases covering all features and edge cases
- 📂 **Multiple Test Datasets** — 5 Excel files for various scenarios

**Improvements:**
- Extracted data processing into `process_input_data()` for reuse
- Config-driven templates (company, buyer, bank, GST rates from YAML)
- All features available as both CLI and importable Python functions

### v0.1.0 — 2025-03-02

**Initial Release:**
- Core invoice generation engine
- Dynamic material support (no hardcoded lists)
- CGST/SGST auto-calculation with HSN codes
- Multi-site support with fiscal year detection
- Robust edge case handling (13 scenarios)
- Cross-platform date formatting

---

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Run tests (`pytest tests/ -v`)
4. Commit your changes (`git commit -m 'Add amazing feature'`)
5. Push and open a Pull Request

---

## 📝 License

MIT — free to use, modify, and distribute.

---

<div align="center">

**Made with ❤️ for India's small businesses**

*If this tool helps your business, consider ⭐ starring the repo!*

</div>

### v0.3.0 (March 2026)
* **Image Processor (AI/OCR):** Added `src/invoice_generator/image_processor.py` to extract structured data directly from photos of handwritten or printed delivery challans.
* **Architecture:** 3-tier fallback execution: OpenRouter Free LLM Vision models -> OCR (EasyOCR) + LLM text structuring -> pure OCR with python heuristics.
* **Workarounds implemented:** LLM model cascading, exponential backoff, JSON repairs, Exif auto-rotation, automatic deduplication, and totals row filtering.
* **Batch processing for images:** Create independent invoices recursively through photos using `--batch`.
* **Refactoring:** Reorganized project into modular `src/invoice_generator` python package for better developer experience and code standard compliance. Added 39 image processing tests.
