import os
import re
import math
from pathlib import Path

import pandas as pd


# =========================================================
# CONFIG
# =========================================================
ROOT_DIR = r"C:\Users\Bartek\Desktop\GEMTECH\CP1\script"
OUTPUT_XLSX = os.path.join(ROOT_DIR, "_MASTER_CP1.xlsx")

FIXED_DATA_COLUMNS = [
    "Numer Number",
    "Nazwa Name",
    "Klasa Class",
    "Klasa Grade",
    "Długość Length [mm]",
    "Szerokość Width [mm]",
    "Grubość Thickness [mm]",
    "Ilość Quantity [szt.]",
    "Ilość z naddatkiem 5% Quantity with excess 5%",
    "Waga na metr Weight per meter [kg/m]",
    "Waga elementu Item weight [kg/szt.]",
    "Waga elementu Weight of the part [kg]",
    "Łącznie waga Total weight [kg]",
    "Powierzchnia elementu Surface area of the element [m2/szt.]",
    "Łącznie powierzchnia Total area [m2]",
    "Powłoka Coating",
    "Norma Standard",
    "Wszystkie dołączone pozycje All items included",
    "Uwagi Comments",
    "Naddatek na spoiny / Weld allowance [%]",
    "Naddatek na spoiny / Weld allowance [kg]",
    "Razem/Total [kg]",
]

METADATA_COLUMNS = [
    "zona",
    "kategoria / category",
    "rewizja / revision",
    "wysokość / height_raw",
    "height_m",
    "firma_raw / company_raw",
    "firma / company",
    "typ_wykonania / fabrication_type",
    "typ_dokumentu / document_type",
    "pozycja_pl / position_pl",
    "position_en",
    "podpozycja_pl / subposition_pl",
    "subposition_en",
    "kg_base",
    "kg_total",
    "ruta_relativa",
    "plik / file",
    "wiersz_źródłowy / source_row",
]

FINAL_COLUMNS = METADATA_COLUMNS + FIXED_DATA_COLUMNS

NUMERIC_COLUMNS = [
    "height_m",
    "Długość Length [mm]",
    "Szerokość Width [mm]",
    "Grubość Thickness [mm]",
    "Ilość Quantity [szt.]",
    "Ilość z naddatkiem 5% Quantity with excess 5%",
    "Waga na metr Weight per meter [kg/m]",
    "Waga elementu Item weight [kg/szt.]",
    "Waga elementu Weight of the part [kg]",
    "Łącznie waga Total weight [kg]",
    "Powierzchnia elementu Surface area of the element [m2/szt.]",
    "Łącznie powierzchnia Total area [m2]",
    "Naddatek na spoiny / Weld allowance [%]",
    "Naddatek na spoiny / Weld allowance [kg]",
    "Razem/Total [kg]",
    "kg_base",
    "kg_total",
]


# =========================================================
# HELPERS
# =========================================================
def clean_text(value):
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
    text = str(value)
    text = text.replace("\xa0", " ")
    text = text.replace("\r", "\n")
    return text.strip()


def one_line(value):
    text = clean_text(value)
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def excel_value_to_python(value):
    if value is None:
        return None
    if isinstance(value, float):
        if math.isnan(value):
            return None
        return value
    text = clean_text(value)
    return text if text != "" else None


def parse_euro_number(value):
    """
    Convierte de forma robusta:
    '18,3' -> 18.3
    '1 341' -> 1341.0
    '1 559,3' -> 1559.3
    '1.559,3' -> 1559.3
    '1,559.3' -> 1559.3
    '30,2 [kg]' -> 30.2
    """
    if value is None:
        return None

    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)

    text = one_line(value)
    if not text:
        return None

    text = re.sub(r"\[[^\]]*\]", "", text).strip()
    text = re.sub(r"[^0-9,.\- ]", "", text).strip()

    if not text:
        return None

    text = text.replace(" ", "").replace("\xa0", "")

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "")
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")
    else:
        if "," in text:
            text = text.replace(",", ".")

    if text.count(".") > 1:
        parts = text.split(".")
        text = "".join(parts[:-1]) + "." + parts[-1]

    try:
        return float(text)
    except ValueError:
        return None


def safe_float_from_height(text):
    if not text:
        return None
    match = re.search(r"H\s*=\s*([\d]+(?:[.,]\d+)?)\s*m?", text, re.IGNORECASE)
    if not match:
        return None
    return float(match.group(1).replace(",", "."))


def normalize_company(folder_name):
    name = one_line(folder_name).lower()

    if "kg+gt" in name or "kg i gt" in name:
        return "KG+GT"
    if name == "kg" or "kg section" in name:
        return "KG"
    if name == "gt" or "gt section" in name:
        return "GT"

    return one_line(folder_name)


def normalize_fabrication_type(folder_name):
    name = one_line(folder_name).lower()

    if "gorącowalcowane" in name or "goracowalcowane" in name:
        return "El. Gorącowalcowane"

    if "zimnogięte" in name or "zimnogiete" in name:
        return "El. Zimnogięte"

    return ""


def infer_doc_type(file_name):
    name = one_line(file_name).lower()

    if "sumaryczne" in name or "summary list" in name:
        return "Summary"
    if "łączników" in name or "lacznik" in name or "connectors" in name:
        return "List of connectors"
    if "elementów wysyłkowych" in name or "elementow wysylkowych" in name or "shipping items" in name:
        return "List of shipping items"
    if "stali profilowej" in name or "profiles" in name:
        return "List of profiles"
    if "c-channel" in name:
        return "List of C-channel"
    if "blachy-obejmy" in name or "steel sheets-clamp" in name:
        return "List of steel sheets-clamp"
    if "blach" in name or "sheets" in name:
        return "List of sheets"

    return "Unknown"

    
def derive_boq_category(category, document_type, fabrication_type):
    """
    Devuelve:
    - pozycja_pl / position_pl
    - position_en
    - podpozycja_pl / subposition_pl
    - subposition_en
    """
    position_pl = ""
    position_en = ""
    subposition_pl = ""
    subposition_en = ""

    if category == "Walls":
        if document_type == "List of connectors":
            position_pl = "Łączniki"
            position_en = "Connectors"
        else:
            position_pl = "Konstrukcja stalowa"
            position_en = "Steel structure"

    elif category == "Ceilings":
        if document_type == "List of connectors":
            position_pl = "Łączniki"
            position_en = "Connectors"

        elif document_type == "List of shipping items":
            position_pl = "Zawieszenie sufitu"
            position_en = "Tbars and hangers"

        else:
            position_pl = "Podkonstrukcja dla sufitów"
            position_en = "Substructure for ceilings"

            if fabrication_type == "El. Zimnogięte":
                subposition_pl = "El. Zimnogięte"
                subposition_en = "Cold-formed steel elements"
            elif fabrication_type == "El. Gorącowalcowane":
                subposition_pl = "El. Gorącowalcowane"
                subposition_en = "Hot-rolled steel elements"

    return {
        "pozycja_pl / position_pl": position_pl,
        "position_en": position_en,
        "podpozycja_pl / subposition_pl": subposition_pl,
        "subposition_en": subposition_en,
    }


# =========================================================
# PATH PARSING
# =========================================================
def parse_context_from_path(file_path, root_dir):
    rel_path = os.path.relpath(file_path, root_dir)
    parts = Path(rel_path).parts

    zona = ""
    categoria = ""
    revision = None
    altura_raw = ""
    altura_m = None
    empresa_raw = ""
    empresa_norm = ""
    fabricacion_tipo = ""

    if parts:
        zona = parts[0]

    for part in parts:
        part_text = one_line(part)
        part_low = part_text.lower()

        if part_low == "walls":
            categoria = "Walls"
        elif part_low == "ceilings":
            categoria = "Ceilings"

        rev_match = re.match(r"revision\s+(\d+)$", part_low)
        if rev_match:
            revision = int(rev_match.group(1))

        if re.search(r"^H\s*=\s*[\d]+(?:[.,]\d+)?m?$", part_text, re.IGNORECASE):
            altura_raw = part_text
            altura_m = safe_float_from_height(part_text)

        company_norm = normalize_company(part_text)
        if company_norm in {"KG", "GT", "KG+GT"}:
            empresa_raw = part_text
            empresa_norm = company_norm

        fabrication = normalize_fabrication_type(part_text)
        if fabrication:
            fabricacion_tipo = fabrication

    return {
        "zona": zona,
        "categoria": categoria,
        "revision": revision,
        "altura_raw": altura_raw,
        "altura_m": altura_m,
        "empresa_raw": empresa_raw,
        "empresa_norm": empresa_norm,
        "fabricacion_tipo": fabricacion_tipo,
        "ruta_relativa": rel_path,
        "archivo": os.path.basename(file_path),
        "tipo_documento": infer_doc_type(os.path.basename(file_path)),
    }


def should_include_file(file_path, ctx):
    if not file_path.lower().endswith(".xls"):
        return False

    if ctx["empresa_norm"] == "KG+GT":
        return False

    if ctx["tipo_documento"] == "List of shipping items":
        return False

    return True


def collect_candidate_files(root_dir):
    files = []

    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if not filename.lower().endswith(".xls"):
                continue

            full_path = os.path.join(dirpath, filename)
            ctx = parse_context_from_path(full_path, root_dir)

            if should_include_file(full_path, ctx):
                files.append({
                    "path": full_path,
                    **ctx
                })

    return files



# =========================================================
# XLS READING
# =========================================================
def open_xls_book(path):
    try:
        import xlrd
    except ImportError as exc:
        raise RuntimeError("Falta xlrd. Instala con: py -m pip install xlrd==2.0.1") from exc

    return xlrd.open_workbook(path, formatting_info=False)


def sheet_cell_value(sheet, row_idx, col_idx):
    try:
        return sheet.cell_value(row_idx, col_idx)
    except Exception:
        return None


def normalize_header_text(text):
    text = clean_text(text)
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def canonical_header(header_text):
    header = normalize_header_text(header_text).lower()

    if (("numer" in header and "number" in header) or header in {"numer", "number"}):
        return "Numer Number"

    if (("nazwa" in header and "name" in header) or header in {"nazwa", "name"}):
        return "Nazwa Name"

    if (("klasa" in header and "class" in header) or header in {"klasa", "class"}):
        return "Klasa Class"

    if (("klasa" in header and "grade" in header) or header == "grade" or "grade" in header):
        return "Klasa Grade"

    if (("długość" in header or "dlugosc" in header or "length" in header or "lenght" in header) and "[mm]" in header):
        return "Długość Length [mm]"

    if (("szerokość" in header or "szerokosc" in header or "width" in header) and "[mm]" in header):
        return "Szerokość Width [mm]"

    if (("grubość" in header or "grubosc" in header or "thickness" in header) and "[mm]" in header):
        return "Grubość Thickness [mm]"

    if ("ilość z naddatkiem 5%" in header or "ilosc z naddatkiem 5%" in header or "quantity with excess 5%" in header):
        return "Ilość z naddatkiem 5% Quantity with excess 5%"

    if (("ilość" in header or "ilosc" in header or "quantity" in header) and "excess" not in header):
        return "Ilość Quantity [szt.]"

    if ("waga na metr" in header or "weight per meter" in header or "kg/m" in header):
        return "Waga na metr Weight per meter [kg/m]"

    if (("waga elementu" in header or "item weight" in header) and ("kg/szt" in header or "kg/pcs" in header or "[kg]" in header)):
        return "Waga elementu Item weight [kg/szt.]"

    if ("weight of the part" in header or ("waga elementu" in header and "part" in header)):
        return "Waga elementu Weight of the part [kg]"

    if ("łącznie waga" in header or "lacznie waga" in header or "total weight" in header):
        return "Łącznie waga Total weight [kg]"

    if ("powierzchnia elementu" in header or "surface area of the element" in header):
        return "Powierzchnia elementu Surface area of the element [m2/szt.]"

    if ("łącznie powierzchnia" in header or "lacznie powierzchnia" in header or "total area" in header):
        return "Łącznie powierzchnia Total area [m2]"

    if ("powłoka" in header or "powloka" in header or "coating" in header):
        return "Powłoka Coating"

    if ("norma" in header or "standard" in header):
        return "Norma Standard"

    if ("wszystkie dołączone pozycje" in header or "wszystkie dolaczone pozycje" in header or "all items included" in header):
        return "Wszystkie dołączone pozycje All items included"

    if ("uwagi" in header or "comments" in header):
        return "Uwagi Comments"

    return None


def row_header_score(row_values):
    score = 0
    canonical_hits = 0

    for value in row_values:
        if canonical_header(value):
            canonical_hits += 1

    score += canonical_hits * 20
    score += sum(1 for value in row_values if clean_text(value))

    return score


def find_header_row(sheet):
    best_row = None
    best_score = -1

    scan_rows = min(30, sheet.nrows)

    for row_idx in range(scan_rows):
        row_values = [sheet_cell_value(sheet, row_idx, col_idx) for col_idx in range(sheet.ncols)]
        score = row_header_score(row_values)

        if score > best_score and score >= 40:
            best_score = score
            best_row = row_idx

    return best_row


def find_data_start_row(sheet, header_row):
    for row_idx in range(header_row + 1, min(header_row + 8, sheet.nrows)):
        values = [excel_value_to_python(sheet_cell_value(sheet, row_idx, col_idx)) for col_idx in range(sheet.ncols)]
        non_empty = [value for value in values if value not in (None, "")]
        if len(non_empty) >= 2:
            return row_idx

    return header_row + 1


# =========================================================
# SUMMARY / FOOTER DETECTION
# =========================================================
def lower_join(values):
    return " | ".join(one_line(value).lower() for value in values if one_line(value))


def is_annotation_row(raw_values):
    joined = lower_join(raw_values)

    patterns = [
        "uwaga/annotation",
        "uwaga / annotation",
        "the list includes",
        "w zestawieniu ujęto ilości",
        "w zestawieniu ujêto ilości",
        "quantity of connectors with excess 5%",
    ]

    return any(pattern in joined for pattern in patterns)


def extract_summary_from_raw_values(raw_values):
    joined = lower_join(raw_values)
    result = {}

    numbers = []
    for value in raw_values:
        parsed = parse_euro_number(value)
        if parsed is not None:
            numbers.append(parsed)

    if "naddatek na spoiny" in joined or "weld allowance" in joined:
        pct_match = re.search(r"(\d+(?:[.,]\d+)?)\s*%", joined)
        if pct_match:
            result["Naddatek na spoiny / Weld allowance [%]"] = float(
                pct_match.group(1).replace(",", ".")
            )

        if numbers:
            result["Naddatek na spoiny / Weld allowance [kg]"] = numbers[-1]

        return result

    return None


def is_real_item_row(row_dict):
    numer = row_dict.get("Numer Number")
    nazwa = row_dict.get("Nazwa Name")

    if numer not in (None, ""):
        return True
    if nazwa not in (None, ""):
        return True

    return False


# =========================================================
# EXTRACT SHEET DATA
# =========================================================
def extract_sheet_rows(path):
    book = open_xls_book(path)
    sheet = book.sheet_by_index(0)

    empty_summary = {
        "Naddatek na spoiny / Weld allowance [%]": None,
        "Naddatek na spoiny / Weld allowance [kg]": None,
    }

    header_row = find_header_row(sheet)
    if header_row is None:
        return [], [], empty_summary

    raw_headers = [normalize_header_text(sheet_cell_value(sheet, header_row, col_idx)) for col_idx in range(sheet.ncols)]
    mapped_headers = [canonical_header(header) for header in raw_headers]

    useful_cols = []
    seen = set()

    for idx, mapped in enumerate(mapped_headers):
        if mapped and mapped not in seen:
            useful_cols.append((idx, mapped, raw_headers[idx]))
            seen.add(mapped)

    if not useful_cols:
        return [], [], empty_summary

    data_start_row = find_data_start_row(sheet, header_row)

    file_summary = {
        "Naddatek na spoiny / Weld allowance [%]": None,
        "Naddatek na spoiny / Weld allowance [kg]": None,
    }

    numeric_field_names = {
        "Długość Length [mm]",
        "Szerokość Width [mm]",
        "Grubość Thickness [mm]",
        "Ilość Quantity [szt.]",
        "Ilość z naddatkiem 5% Quantity with excess 5%",
        "Waga na metr Weight per meter [kg/m]",
        "Waga elementu Item weight [kg/szt.]",
        "Waga elementu Weight of the part [kg]",
        "Łącznie waga Total weight [kg]",
        "Powierzchnia elementu Surface area of the element [m2/szt.]",
        "Łącznie powierzchnia Total area [m2]",
    }

    rows = []

    for row_idx in range(data_start_row, sheet.nrows):
        raw_values = [sheet_cell_value(sheet, row_idx, col_idx) for col_idx in range(sheet.ncols)]

        if all(excel_value_to_python(value) in (None, "") for value in raw_values):
            continue

        if is_annotation_row(raw_values):
            continue

        summary_data = extract_summary_from_raw_values(raw_values)
        if summary_data:
            for key, value in summary_data.items():
                if value is not None:
                    file_summary[key] = value
            continue

        row_dict = {col: None for col in FIXED_DATA_COLUMNS}
        non_empty_count = 0

        for col_idx, mapped_name, _raw_header in useful_cols:
            value = excel_value_to_python(sheet_cell_value(sheet, row_idx, col_idx))

            if mapped_name in numeric_field_names:
                parsed = parse_euro_number(value)
                if parsed is not None:
                    value = parsed

            if value not in (None, ""):
                non_empty_count += 1

            row_dict[mapped_name] = value

        if non_empty_count == 0:
            continue

        if not is_real_item_row(row_dict):
            continue

        row_dict["_source_row"] = row_idx + 1
        rows.append(row_dict)

    for row in rows:
        row["Naddatek na spoiny / Weld allowance [%]"] = file_summary.get("Naddatek na spoiny / Weld allowance [%]")
        row["Naddatek na spoiny / Weld allowance [kg]"] = file_summary.get("Naddatek na spoiny / Weld allowance [kg]")

    detected_headers = [raw_header for _, _, raw_header in useful_cols]
    return rows, detected_headers, file_summary


# =========================================================
# RECALC TOTALS
# =========================================================
def recalc_document_totals(df_data):
    if df_data.empty:
        return df_data

    df_data = df_data.copy()

    total_weight = pd.to_numeric(df_data["Łącznie waga Total weight [kg]"], errors="coerce")
    allowance_pct = pd.to_numeric(df_data["Naddatek na spoiny / Weld allowance [%]"], errors="coerce")

    df_data["Razem/Total [kg]"] = total_weight * (1 + allowance_pct.fillna(0) / 100.0)

    return df_data


# =========================================================
# BUILD DATAFRAME
# =========================================================
def build_master_dataframe(root_dir):
    files_info = collect_candidate_files(root_dir)

    all_records = []
    files_summary = []

    for file_info in files_info:
        path = file_info["path"]

        try:
            rows, detected_headers, file_summary = extract_sheet_rows(path)

            files_summary.append({
                "ruta_relativa": file_info["ruta_relativa"],
                "archivo": file_info["archivo"],
                "zona": file_info["zona"],
                "categoria": file_info["categoria"],
                "revision": file_info["revision"],
                "altura_raw": file_info["altura_raw"],
                "altura_m": file_info["altura_m"],
                "empresa_raw": file_info["empresa_raw"],
                "empresa_norm": file_info["empresa_norm"],
                "fabricacion_tipo": file_info["fabricacion_tipo"],
                "tipo_documento": file_info["tipo_documento"],
                "columnas_detectadas": " || ".join(detected_headers),
                "filas_detectadas": len(rows),
                "weld_allowance_pct": file_summary.get("Naddatek na spoiny / Weld allowance [%]"),
                "weld_allowance_kg": file_summary.get("Naddatek na spoiny / Weld allowance [kg]"),
                "estado": "OK",
            })
            
            for row in rows:
                boq_info = derive_boq_category(
                    file_info["categoria"],
                    file_info["tipo_documento"],
                    file_info["fabricacion_tipo"],
                )

                record = {
                    "zona": file_info["zona"],
                    "kategoria / category": file_info["categoria"],
                    "rewizja / revision": file_info["revision"],
                    "wysokość / height_raw": file_info["altura_raw"],
                    "height_m": file_info["altura_m"],
                    "firma_raw / company_raw": file_info["empresa_raw"],
                    "firma / company": file_info["empresa_norm"],
                    "typ_wykonania / fabrication_type": file_info["fabricacion_tipo"],
                    "typ_dokumentu / document_type": file_info["tipo_documento"],
                    "pozycja_pl / position_pl": boq_info["pozycja_pl / position_pl"],
                    "position_en": boq_info["position_en"],
                    "podpozycja_pl / subposition_pl": boq_info["podpozycja_pl / subposition_pl"],
                    "subposition_en": boq_info["subposition_en"],
                    "kg_base": row.get("Łącznie waga Total weight [kg]"),
                    "kg_total": row.get("Razem/Total [kg]"),
                    "ruta_relativa": file_info["ruta_relativa"],
                    "plik / file": file_info["archivo"],
                    "wiersz_źródłowy / source_row": row.pop("_source_row", None),
                }

                for col in FIXED_DATA_COLUMNS:
                    record[col] = row.get(col)

                all_records.append(record)

        except Exception as exc:
            files_summary.append({
                "ruta_relativa": file_info["ruta_relativa"],
                "archivo": file_info["archivo"],
                "zona": file_info["zona"],
                "categoria": file_info["categoria"],
                "revision": file_info["revision"],
                "altura_raw": file_info["altura_raw"],
                "altura_m": file_info["altura_m"],
                "empresa_raw": file_info["empresa_raw"],
                "empresa_norm": file_info["empresa_norm"],
                "fabricacion_tipo": file_info["fabricacion_tipo"],
                "tipo_documento": file_info["tipo_documento"],
                "columnas_detectadas": "",
                "filas_detectadas": 0,
                "weld_allowance_pct": None,
                "weld_allowance_kg": None,
                "estado": f"ERROR: {exc}",
            })

    df_data = pd.DataFrame(all_records)
    if df_data.empty:
        df_data = pd.DataFrame(columns=FINAL_COLUMNS)
    else:
        for col in FINAL_COLUMNS:
            if col not in df_data.columns:
                df_data[col] = None

        df_data = df_data[FINAL_COLUMNS]

        for col in NUMERIC_COLUMNS:
            if col in df_data.columns:
                df_data[col] = pd.to_numeric(df_data[col], errors="coerce")

        df_data = recalc_document_totals(df_data)
        
        df_data["kg_base"] = pd.to_numeric(df_data["Łącznie waga Total weight [kg]"], errors="coerce")
        df_data["kg_total"] = pd.to_numeric(df_data["Razem/Total [kg]"], errors="coerce")

    df_files = pd.DataFrame(files_summary)
    df_columns = pd.DataFrame({"column_name": FINAL_COLUMNS})

    return df_data, df_files, df_columns


# =========================================================
# EXPORT
# =========================================================
def autosize_worksheet(ws, dataframe):
    for idx, col in enumerate(dataframe.columns, start=1):
        max_len = len(str(col))
        sample = dataframe[col].astype(str).fillna("").head(300)

        for value in sample:
            if value == "None":
                value = ""
            max_len = max(max_len, len(value))

        ws.column_dimensions[ws.cell(1, idx).column_letter].width = min(max_len + 2, 60)


def export_to_excel(df_data, df_files, df_columns, output_path):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_data.to_excel(writer, sheet_name="DATA", index=False)
        df_files.to_excel(writer, sheet_name="FILES", index=False)
        df_columns.to_excel(writer, sheet_name="COLUMNS", index=False)

        for sheet_name, dataframe in [
            ("DATA", df_data),
            ("FILES", df_files),
            ("COLUMNS", df_columns),
        ]:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)

            if not dataframe.empty:
                autosize_worksheet(ws, dataframe)


# =========================================================
# MAIN
# =========================================================
def main():
    print(f"Escaneando: {ROOT_DIR}")

    df_data, df_files, df_columns = build_master_dataframe(ROOT_DIR)

    print(f"Filas DATA: {len(df_data)}")
    print(f"Ficheros procesados: {len(df_files)}")

    export_to_excel(df_data, df_files, df_columns, OUTPUT_XLSX)

    print("OK. Excel generado en:")
    print(OUTPUT_XLSX)


if __name__ == "__main__":
    main()