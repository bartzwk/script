# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

CP1 Material Data Aggregation System for GEMTECH — processes bill of materials (BOM) data from multiple Excel XLS files across Sections B and C, consolidating them into master Excel files for engineering/construction reporting.

## Running the Scripts

```bash
# Generate master consolidated XLSX from all XLS source files
python create_master_excel.py

# Generate deduplicated assembled BOM (.xlsx). Multi-project (CP1/FA1).
python create_assembled_excel.py        # interactive project menu
python create_assembled_excel.py CP1    # only CP1
```

`create_master_excel.py` reads from the `B/` and `C/` directory trees; `create_assembled_excel.py` is driven by `PROJECT_CONFIGS` (root dir + search folders per project). All generated files (master/assembled Excel files and `find_text_in_pdfs.py` result TXTs) are written to the gitignored `out/` folder next to the scripts.

## Dependencies

Install with pip:
```bash
pip install pandas xlrd==2.0.1 openpyxl
```

`xlrd` must be version 2.0.1 — newer versions dropped XLS support.

## Architecture

### Data Flow

1. **Discovery**: Both scripts recursively scan the `4. List of steel elements and connectors/` subtrees of the configured section folders for `.xls` files (`create_master_excel.py` uses `B/` and `C/`; `create_assembled_excel.py` uses each project's `search_folders`).
2. **Context Extraction**: Metadata is derived from directory path components — section (B/C), company (GT/KG), height variant (H=3,50m / H=6,10m), revision.
3. **Document Classification**: File names are pattern-matched to determine type (profiles, sheets, C-channels, clamps, shipping items, connectors).
4. **Header Detection**: Each XLS sheet uses a scoring algorithm to locate the actual data header row — necessary because source files have inconsistent layouts.
5. **Data Extraction & Normalization**: Rows are parsed, Polish/European number formats converted (`1.234,56` → `1234.56`), Polish diacritics normalized.
6. **Aggregation/Deduplication**: `create_assembled_excel.py` deduplicates by normalized part numbers (first occurrence wins) and filters to element numbers 0–999.
7. **Export**: both scripts output `.xlsx` via openpyxl/pandas; `create_assembled_excel.py` adds formulas, an Excel table, and a totals row.

### File Responsibilities

| File | Output | Purpose |
|------|--------|---------|
| `create_master_excel.py` | `out/_MASTER_CP1.xlsx` | Full consolidated BOM with metadata, BOQ categorization, weld allowance calculations |
| `create_assembled_excel.py` | `out/master <PROJECT> assembled.xlsx` | Deduplicated assembled BOM with reconstructed component codes, formulas, and totals |

### Key Design Patterns

- **Path-derived metadata**: Section, company, height, and revision are extracted purely from the directory structure — not from file content.
- **Scoring-based header detection**: Each sheet scans candidate rows and scores them against known Polish/English column name patterns to find the real header.
- **Bilingual columns**: Output files use parallel Polish/English column names throughout (e.g., `Ilość`/quantity, `Waga`/weight).
- **BOQ categorization**: Items are mapped to construction sections (Ceilings/Walls × height variants × company) for quantity reporting.
- **Weld allowance**: `create_master_excel.py` applies a +1.8% weld allowance to net weights, tracked in a dedicated `Naddatek na spoiny` column.

### Hardcoded Paths

`create_master_excel.py` has `ROOT_DIR` hardcoded at the top:
```python
ROOT_DIR = r"C:\Users\Bartek\Desktop\GEMTECH\CP1\script"
```

`create_assembled_excel.py` instead keeps per-project roots in `PROJECT_CONFIGS` (`root_dir`, `search_folders`, `output_name`, `prefix_pattern`). Change these if the project moves.
