# OMW Risk Register - Local Excel-to-JSON Extractor

A lightweight, company-grade, local Python command-line application that ingests an Excel workbook (`.xlsx`, `.xlsm`), extracts all worksheet data, maps headers to the canonical 24 Master Fields across 5 sections, and exports structured JSON files.

No web server, no database, no frontend, and zero external runtime dependencies beyond `openpyxl`.

---

## 1. Directory Structure

```text
omw-extractor/
├── main.py              # CLI entry point (prompts for path or accepts argument)
├── excel_reader.py      # openpyxl loader, unrolls merged cells, extracts raw data
├── header_mapper.py     # Maps headers to the 24 master fields across 5 sections
├── risk_extractor.py    # Extracts risk records preserving values, cell refs, and row numbers
├── test_pipeline.py     # Unit and end-to-end integration tests
├── requirements.txt     # Python dependencies (openpyxl)
└── output/              # Auto-created output directory
    ├── workbook_raw.json       # Complete raw extracted workbook structure
    ├── risk_records.json       # Clean mapped risk records adhering to the master schema
    └── extraction_report.json  # Worksheet names, detected headers, and record counts
```

---

## 2. Installation & Quickstart (Windows PowerShell)

### Step 1: Install Requirements
```powershell
python -m pip install -r omw-extractor/requirements.txt
```

### Step 2: Run the Pipeline

You can either pass the Excel file path directly:
```powershell
python omw-extractor/main.py "C:\path\to\your\OMW_Risk_Register.xlsx"
```

Or run `python omw-extractor/main.py` and paste the file path at the interactive prompt:
```powershell
python omw-extractor/main.py
```

### Step 3: Run Automated Tests
```powershell
python -m pytest omw-extractor/test_pipeline.py -v
```

---

## 3. Output Files Generated in `output/`

1. **`workbook_raw.json`**: Complete raw workbook extraction preserving all worksheets, dimensions, cells, row numbers, and merged cell relationships.
2. **`risk_records.json`**: Risk records formatted with `sourceSheet` and the 5 Master Sections (`hazardIdentification`, `initialRiskEvaluation`, `riskTreatment`, `postResponseRiskEvaluation`, `commentsAndAudit`).
3. **`extraction_report.json`**: Structural audit report with worksheet names, mapped headers, unmapped columns, and record tallies.
