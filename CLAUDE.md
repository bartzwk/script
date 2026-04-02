# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

CP1 Material Data Aggregation System for GEMTECH — processes bill of materials (BOM) data from multiple Excel XLS files across Sections B and C, consolidating them into master Excel files for engineering/construction reporting.

## Running the Scripts

```bash
# Generate master consolidated XLSX from all XLS source files
python create_master_excel.py

# Generate delivered/assembled BOM in native XLS format
python create_delivered_structure_excel.py
```

Both scripts read from the `B/` and `C/` directory trees and write output files to the root `script/` directory. No arguments needed.

## Dependencies

Install with pip:
```bash
pip install pandas xlrd==2.0.1 xlwt openpyxl
```

`xlrd` must be version 2.0.1 — newer versions dropped XLS support.

## Architecture

### Data Flow

1. **Discovery**: Both scripts recursively scan `B/4. List of steel elements and connectors/` and `C/4. List of steel elements and connectors/` for `.xls` files.
2. **Context Extraction**: Metadata is derived from directory path components — section (B/C), company (GT/KG), height variant (H=3,50m / H=6,10m), revision.
3. **Document Classification**: File names are pattern-matched to determine type (profiles, sheets, C-channels, clamps, shipping items, connectors).
4. **Header Detection**: Each XLS sheet uses a scoring algorithm to locate the actual data header row — necessary because source files have inconsistent layouts.
5. **Data Extraction & Normalization**: Rows are parsed, Polish/European number formats converted (`1.234,56` → `1234.56`), Polish diacritics normalized.
6. **Aggregation/Deduplication**: `create_delivered_structure_excel.py` deduplicates by normalized part numbers, prioritizing shipping items over other document types.
7. **Export**: `create_master_excel.py` outputs `.xlsx` via openpyxl/pandas; `create_delivered_structure_excel.py` outputs native `.xls` via xlwt.

### File Responsibilities

| File | Output | Purpose |
|------|--------|---------|
| `create_master_excel.py` | `_MASTER_CP1_B+C.xlsx` | Full consolidated BOM with metadata, BOQ categorization, weld allowance calculations |
| `create_delivered_structure_excel.py` | `master CP1 assembled.xls` | Deduplicated assembled BOM with reconstructed component codes |

### Key Design Patterns

- **Path-derived metadata**: Section, company, height, and revision are extracted purely from the directory structure — not from file content.
- **Scoring-based header detection**: Each sheet scans candidate rows and scores them against known Polish/English column name patterns to find the real header.
- **Bilingual columns**: Output files use parallel Polish/English column names throughout (e.g., `Ilość`/quantity, `Waga`/weight).
- **BOQ categorization**: Items are mapped to construction sections (Ceilings/Walls × height variants × company) for quantity reporting.
- **Weld allowance**: `create_master_excel.py` applies a +1.8% weld allowance to net weights, tracked in a dedicated `Naddatek na spoiny` column.

### Hardcoded Paths

Both scripts have `ROOT_DIR` hardcoded at the top:
```python
ROOT_DIR = r"C:\Users\Bartek\Desktop\GEMTECH\CP1\script"
```

Change this constant if the project moves.
