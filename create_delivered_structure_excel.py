import os
import re
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import xlrd
import xlwt

# =========================================================
# CONFIG
# =========================================================
ROOT_DIR = r"C:\Users\Bartek\Desktop\GEMTECH\CP1\script"
OUTPUT_XLS = os.path.join(ROOT_DIR, "master CP1 assembled.xls")
SEARCH_TOP_FOLDERS = ["B", "C"]

OUTPUT_COLUMNS = [
    "Nazwa / Name",
    "Strefa / Zone",
    "Symbol / Symbol",
    "Element / Element",
    "Numer elementu / Element number",
    "Rewizja / REV",
    "Część 1 / Part 1",
    "Część 2 / Part 2",
    "Konstrukcja / Construction",
    "Typ stali / Steel type",
    "Ilość / Quantity",
    "Nazwa opisowa / Name",
    "Długość (mm) / Length (mm)",
    "Waga netto (kg/szt.) / Weight netto (kg/pcs.)",
    "Waga netto (kg) / Weight netto (kg)",
    "Waga brutto (kg) +1,8% / Weight brutto (kg) +1.8%",
    "Data odbioru dostawy / Delivery acceptance date",
    "Powierzchnia elementu (m²/szt.) / Surface area (m²/pcs.)",
    "Wszystkie dołączone pozycje / All items included",
    "Ścieżka pliku / Source file path",
    "Numer wiersza / Source row number",
]

# =========================================================
# TEXT HELPERS
# =========================================================



def safe_int(value: object) -> Optional[int]:
    f = safe_float(value)
    if f is None:
        return None
    if abs(f - round(f)) < 1e-9:
        return int(round(f))
    return None


def excel_number(value: object) -> object:
    f = safe_float(value)
    if f is None:
        return clean_value(value)
    if abs(f - round(f)) < 1e-9:
        return int(round(f))
    return f


def extract_revision_from_text(text: object) -> Optional[int]:
    s = str(text or "")
    m = re.search(r"(?:^|[^A-Z0-9])R\s*(\d+)(?:[^A-Z0-9]|$)", s, flags=re.IGNORECASE)
    if m:
        return int(m.group(1))
    m = re.search(r"Revision\s*(\d+)", s, flags=re.IGNORECASE)
    if m:
        return int(m.group(1))
    return None

def strip_accents(text: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", str(text))
        if not unicodedata.combining(ch)
    )


def normalize_text(text: object) -> str:
    if text is None:
        return ""
    s = str(text).replace("\xa0", " ").replace("\n", " ").replace("\r", " ")
    s = strip_accents(s).lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_key(text: object) -> str:
    s = normalize_text(text)
    s = s.replace(" ", "")
    return s


def clean_value(value: object) -> object:
    if value is None:
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return value
    return str(value).replace("\xa0", " ").strip()


def split_included_items(value: object) -> List[str]:
    s = str(value or "").replace(";", ",")
    parts = [p.strip() for p in s.split(",")]
    return [p for p in parts if p]


def safe_float(value: object) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip().replace(" ", "")
    if not s:
        return None

    # intentar formatos 1 234,56 y 1234.56
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


# =========================================================
# FILE CLASSIFICATION
# =========================================================

def get_top_area(rel_path: str) -> str:
    parts = Path(rel_path).parts
    return parts[0] if parts else ""


def is_summary_company_folder(path_parts_lower: List[str]) -> bool:
    for p in path_parts_lower:
        p_norm = normalize_text(p)
        if p_norm in {"kg+gt", "kg i gt section", "kg i gt", "kg&gt", "gt+kg"}:
            return True
    return False


def detect_company(path_parts: Tuple[str, ...], file_name: str) -> str:
    candidates = [file_name, *path_parts]
    joined = " | ".join(candidates)
    n = normalize_text(joined)
    if re.search(r"(^|[^a-z])gt([^a-z]|$)", n) and not re.search(r"kg i gt|kg\+gt|gt\+kg", n):
        if re.search(r"(^|[^a-z])kg([^a-z]|$)", n):
            return "MULTI"
        return "GT"
    if re.search(r"(^|[^a-z])kg([^a-z]|$)", n):
        return "KG"
    return "UNKNOWN"


def infer_doc_type(file_name: str) -> Optional[str]:
    n = normalize_text(file_name)
    if "zestawienie lacznikow" in n or "list of connectors" in n:
        return None
    if "zestawienie elementow wysylkowych" in n or "shipping items" in n:
        return "shipping"
    if "zestawienie stali profilowej" in n or "list of profiles" in n:
        return "profiles"
    if "zestawienie blachy-obejmy" in n or "steel sheets-clamp" in n:
        return "clamp_sheets"
    if "zestawienie c-channel" in n or "list of c-channel" in n:
        return "c_channel"
    if "zestawienie blach" in n or "list of sheets" in n:
        return "sheets"
    return None


def should_skip_file(full_path: str, root_dir: str) -> Tuple[bool, str]:
    rel_path = os.path.relpath(full_path, root_dir)
    parts = Path(rel_path).parts
    parts_lower = [p.lower() for p in parts]
    file_name = os.path.basename(full_path)
    file_name_norm = normalize_text(file_name)

    if len(parts) < 2:
        return True, "not_inside_B_or_C_subfolder"

    if parts[0].upper() not in SEARCH_TOP_FOLDERS:
        return True, "outside_B_C"

    if file_name.startswith("~$"):
        return True, "temp_file"

    if not file_name.lower().endswith(".xls"):
        return True, "not_xls"

    if is_summary_company_folder(parts_lower):
        return True, "summary_company_folder"

    if file_name_norm in {
        normalize_text("salida.xls"),
        normalize_text("master CP1 assembled.xls"),
    }:
        return True, "generated_output"

    if file_name_norm.startswith(normalize_text("master ")):
        return True, "generated_output"

    doc_type = infer_doc_type(file_name)
    if doc_type is None:
        return True, "unsupported_or_connectors"

    return False, doc_type


# =========================================================
# EXCEL READING
# =========================================================

def read_sheet_matrix(xls_path: str) -> List[List[object]]:
    book = xlrd.open_workbook(xls_path)
    sheet = book.sheet_by_index(0)
    matrix = []
    for r in range(sheet.nrows):
        matrix.append([sheet.cell_value(r, c) for c in range(sheet.ncols)])
    return matrix


def score_header_row(row: List[object]) -> int:
    hits = 0
    for cell in row:
        t = normalize_text(cell)
        if "numer" in t or "number" in t:
            hits += 1
        if "nazwa" in t or "name" in t:
            hits += 1
        if "ilosc" in t or "quantity" in t:
            hits += 1
        if "dlugosc" in t or "lenght" in t or "length" in t:
            hits += 1
        if "waga elementu" in t or "item weight" in t:
            hits += 1
        if "lacznie waga" in t or "total weight" in t:
            hits += 1
        if "wszystkie dolaczone pozycje" in t or "all items included" in t:
            hits += 2
    return hits


def find_header_row(matrix: List[List[object]]) -> Optional[int]:
    best_idx = None
    best_score = -1
    for i, row in enumerate(matrix[:20]):
        score = score_header_row(row)
        if score > best_score:
            best_score = score
            best_idx = i
    if best_score < 2:
        return None
    return best_idx


def build_header_map(header_row: List[object]) -> Dict[str, int]:
    header_map: Dict[str, int] = {}
    for idx, cell in enumerate(header_row):
        t = normalize_text(cell)
        if not t:
            continue

        if "numer" in t or "number" in t:
            header_map.setdefault("number", idx)
        if "nazwa" in t or ("name" in t and "project" not in t):
            header_map.setdefault("name", idx)
        if "ilosc" in t or "quantity" in t:
            header_map.setdefault("quantity", idx)
        if "dlugosc" in t or "lenght" in t or "length" in t:
            header_map.setdefault("length", idx)
        if "szerokosc" in t or "width" in t:
            header_map.setdefault("width", idx)
        if "grubosc" in t or "thickness" in t:
            header_map.setdefault("thickness", idx)
        if "waga elementu" in t or "item weight" in t:
            header_map.setdefault("item_weight", idx)
        if "lacznie waga" in t or "total weight" in t:
            header_map.setdefault("total_weight", idx)
        if "powierzchnia elementu" in t or "surface area" in t:
            header_map.setdefault("surface", idx)
        if "klasa" in t or "class" in t:
            header_map.setdefault("class", idx)
        if "grade" in t:
            header_map.setdefault("grade", idx)
        if "wszystkie dolaczone pozycje" in t or "all items included" in t:
            header_map.setdefault("included", idx)
    return header_map


def find_first_data_row(matrix: List[List[object]], header_row_idx: int, header_map: Dict[str, int]) -> int:
    num_col = header_map.get("number")
    if num_col is None:
        return header_row_idx + 1
    for r in range(header_row_idx + 1, len(matrix)):
        row = matrix[r]
        cell = row[num_col] if num_col < len(row) else ""
        if normalize_text(cell):
            return r
    return len(matrix)


def row_to_record(row: List[object], header_map: Dict[str, int], meta: Dict[str, str], source_row_number: int) -> Optional[Dict[str, object]]:
    def get(key: str) -> object:
        idx = header_map.get(key)
        if idx is None or idx >= len(row):
            return ""
        return clean_value(row[idx])

    number = get("number")
    if not str(number).strip():
        return None

    rec = {
        "zone": meta["zone"],
        "company": meta["company"],
        "source_type": meta["doc_type"],
        "revision": meta.get("revision"),
        "number": number,
        "name": get("name"),
        "quantity": get("quantity"),
        "length": get("length"),
        "width": get("width"),
        "thickness": get("thickness"),
        "item_weight": get("item_weight"),
        "total_weight": get("total_weight"),
        "surface": get("surface"),
        "class_grade": get("class") or get("grade"),
        "included": get("included"),
        "steel_type": meta.get("steel_type", ""),
        "source_file": meta["rel_path"],
        "source_row_number": source_row_number,
    }
    rec["number_key"] = normalize_key(rec["number"])
    return rec


def read_records_from_file(item: Dict[str, str]) -> List[Dict[str, object]]:
    path = item["path"]
    try:
        matrix = read_sheet_matrix(path)
    except Exception as exc:
        print(f"[ERROR:OPEN] {path} -> {exc}")
        return []

    header_row_idx = find_header_row(matrix)
    if header_row_idx is None:
        print(f"[WARN] Cabecera no encontrada en: {path}")
        return []

    header_map = build_header_map(matrix[header_row_idx])
    first_data_row = find_first_data_row(matrix, header_row_idx, header_map)

    print(f"[HEADER] {item['rel_path']}")
    print(f"         header_row={header_row_idx + 1}, first_data_row={first_data_row + 1}, header_map={header_map}")

    records = []
    for r in range(first_data_row, len(matrix)):
        rec = row_to_record(matrix[r], header_map, item, r + 1)
        if rec is None:
            continue
        records.append(rec)

    print(f"[ROWS_READ] {item['rel_path']} -> {len(records)}")
    return records


# =========================================================
# DISCOVERY
# =========================================================

def collect_files(root_dir: str) -> List[Dict[str, str]]:
    found: List[Dict[str, str]] = []

    for top in SEARCH_TOP_FOLDERS:
        start_dir = os.path.join(root_dir, top)
        print(f"[SCAN] Buscando dentro de: {start_dir}")
        if not os.path.isdir(start_dir):
            print(f"[WARN] No existe carpeta: {start_dir}")
            continue

        for dirpath, _, filenames in os.walk(start_dir):
            for filename in filenames:
                full_path = os.path.join(dirpath, filename)
                if not filename.lower().endswith(".xls"):
                    continue

                print(f"[XLS] {full_path}")
                skip, reason = should_skip_file(full_path, root_dir)
                if skip:
                    print(f"[SKIP:{reason}] {full_path}")
                    continue

                rel_path = os.path.relpath(full_path, root_dir)
                parts = Path(rel_path).parts
                company = detect_company(parts, filename)
                zone = get_top_area(rel_path)

                revision = extract_revision_from_text(rel_path)

                steel_type = ""
                for part in parts:
                    p_norm = normalize_text(part)
                    if "goracowalcowane" in p_norm:
                        steel_type = "gorącowalcowane"
                        break
                    elif "zimnogiete" in p_norm:
                        steel_type = "zimnogięte"
                        break

                item = {
                    "path": full_path,
                    "rel_path": rel_path,
                    "file_name": filename,
                    "doc_type": reason,
                    "company": company,
                    "zone": zone,
                    "revision": revision,
                    "steel_type": steel_type,
                }
                print(f"[USE:{reason}] {rel_path} | company={company}")
                found.append(item)

    return found


# =========================================================
# PARSING OUTPUT COLUMNS
# =========================================================

def parse_number_fields(number: object, zone_from_path: str = "", company_from_meta: str = "", revision_from_meta: object = None) -> Dict[str, object]:
    raw = str(number or "").strip()
    if not raw:
        rev_val = safe_int(revision_from_meta)
        return {
            "name_code": "",
            "zone": zone_from_path,
            "symbol": company_from_meta if company_from_meta in {"GT", "KG"} else "",
            "element": "",
            "element_number": "",
            "rev": rev_val if rev_val is not None else "",
            "part1": "",
            "part2": "",
            "construction": "",
        }

    s = raw.strip().replace("/", "_")
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"^CP\d+[-_ ]*", "", s, flags=re.IGNORECASE)

    rev_val = extract_revision_from_text(s)
    if rev_val is None:
        rev_val = safe_int(revision_from_meta)
    s = re.sub(r"^R\s*\d+[-_ ]*", "", s, flags=re.IGNORECASE)

    zone = zone_from_path
    m_zone = re.match(r"^([A-Z]\d*[A-Z]?)[-_ ]+", s, flags=re.IGNORECASE)
    if m_zone:
        zone = m_zone.group(1).upper().replace(" ", "")
        s = s[m_zone.end():]

    symbol = company_from_meta if company_from_meta in {"GT", "KG"} else ""
    m_symbol = re.search(r"(?:^|[-_ ])(GT|KG)(?:[-_ ]|$)", s, flags=re.IGNORECASE)
    if m_symbol:
        symbol = m_symbol.group(1).upper()
        rest = s[m_symbol.end():].strip("_- ")
    else:
        rest = s.strip("_- ")

    m_num = re.search(r"(?:[-_])?(\d+)$", rest)
    element_number: object = ""
    element = ""
    if m_num:
        element_number = int(m_num.group(1))
        element = rest[:m_num.start()].strip("_- ")
    else:
        element = rest.strip("_- ")

    element = re.sub(r"\s+", " ", element).strip()
    if not element:
        element = "_"

    if rev_val not in (None, ""):
        name_prefix = f"R{int(rev_val)}_"
    else:
        name_prefix = ""

    if element and element_number != "":
        name_code = f"{name_prefix}{zone}_{symbol}_{element}-{element_number}".strip("_")
    elif element:
        name_code = f"{name_prefix}{zone}_{symbol}_{element}".strip("_")
    elif element_number != "":
        name_code = f"{name_prefix}{zone}_{symbol}-{element_number}".strip("_")
    else:
        name_code = raw

    construction = ""
    if isinstance(element_number, int):
        n = element_number
        if 0 <= n < 500:
            construction = "wall"
        elif (500 <= n <= 599) or (700 <= n <= 799) or (900 <= n <= 999):
            construction = "ceiling sub"
        elif (600 <= n <= 699) or (800 <= n <= 899):
            construction = "panel sub"

    return {
        "name_code": name_code,
        "zone": zone,
        "symbol": symbol,
        "element": element,
        "element_number": element_number,
        "rev": int(rev_val) if rev_val not in (None, "") else "",
        "part1": "",
        "part2": "",
        "construction": construction,
    }


def build_output_row(rec: Dict[str, object]) -> Dict[str, object]:
    parsed = parse_number_fields(
        rec.get("number", ""),
        str(rec.get("zone", "")),
        str(rec.get("company", "")),
        rec.get("revision"),
    )

    total_weight_val = safe_float(rec.get("total_weight", ""))
    gross_weight = round(total_weight_val * 1.018, 3) if total_weight_val is not None else ""

    return {
        "name_code": rec.get("number", ""),
        "zone": parsed["zone"],
        "symbol": parsed["symbol"],
        "element": parsed["element"],
        "element_number": parsed["element_number"],
        "rev": parsed["rev"],
        "part1": parsed["part1"],
        "part2": parsed["part2"],
        "construction": parsed["construction"],
        "steel_type": rec.get("steel_type", "") if parsed["construction"] == "ceiling sub" else "",
        "quantity": excel_number(rec.get("quantity", "")),
        "name": rec.get("name", ""),
        "length": excel_number(rec.get("length", "")),
        "item_weight": excel_number(rec.get("item_weight", "")),
        "total_weight": excel_number(rec.get("total_weight", "")),
        "gross_weight": excel_number(gross_weight),
        "delivery_acceptance_date": "",
        "surface": excel_number(rec.get("surface", "")),
        "included": rec.get("included", ""),
        "source_file": rec.get("source_file", ""),
        "source_row_number": excel_number(rec.get("source_row_number", "")),
    }


# =========================================================
# OUTPUT
# =========================================================

def write_output_xls(output_path: str, rows: List[Dict[str, object]]) -> None:
    book = xlwt.Workbook()
    ws = book.add_sheet("assembled")

    header_style = xlwt.easyxf("font: bold on; pattern: pattern solid, fore_colour ice_blue;")
    for c, col in enumerate(OUTPUT_COLUMNS):
        ws.write(0, c, col, header_style)

    num_style = xlwt.easyxf(num_format_str="0.###")
    int_style = xlwt.easyxf(num_format_str="0")

    for r, row in enumerate(rows, start=1):
        values = [
            row.get("name_code", ""),
            row.get("zone", ""),
            row.get("symbol", ""),
            row.get("element", ""),
            row.get("element_number", ""),
            row.get("rev", ""),
            row.get("part1", ""),
            row.get("part2", ""),
            row.get("construction", ""),
            row.get("steel_type", ""),
            row.get("quantity", ""),
            row.get("name", ""),
            row.get("length", ""),
            row.get("item_weight", ""),
            row.get("total_weight", ""),
            row.get("gross_weight", ""),
            row.get("delivery_acceptance_date", ""),
            row.get("surface", ""),
            row.get("included", ""),
            row.get("source_file", ""),
            row.get("source_row_number", ""),
        ]
        int_cols = {4, 5, 10, 20}
        for c, value in enumerate(values):
            if isinstance(value, int):
                ws.write(r, c, value, int_style if c in int_cols else num_style)
            elif isinstance(value, float):
                ws.write(r, c, value, num_style)
            else:
                ws.write(r, c, value)

    widths = [
        28, 10, 20, 14, 16, 12, 12, 12, 18, 20,
        10, 32, 14, 18, 18, 20, 20, 22, 34, 60, 14,
    ]
    for idx, width in enumerate(widths):
        ws.col(idx).width = width * 256

    book.save(output_path)


# =========================================================
# MAIN
# =========================================================

def main() -> None:
    files = collect_files(ROOT_DIR)
    if not files:
        print("[ERROR] No se han encontrado archivos válidos en B y C.")
        return

    print(f"Archivos válidos: {len(files)}")

    final_by_number: Dict[str, Dict[str, object]] = {}
    skipped_out_of_range = 0

    for item in files:
        records = read_records_from_file(item)
        for rec in records:
            parsed = parse_number_fields(
                rec.get("number", ""),
                str(rec.get("zone", "")),
                str(rec.get("company", "")),
                rec.get("revision"),
            )
            en = parsed["element_number"]
            if not isinstance(en, int) or en < 0 or en > 999:
                skipped_out_of_range += 1
                continue
            key = str(rec["number_key"])
            if key not in final_by_number:
                final_by_number[key] = rec

    print(f"Elementos fuera de rango (>=1000 o sin número): {skipped_out_of_range}")

    final_rows = [build_output_row(rec) for rec in final_by_number.values()]
    final_rows.sort(key=lambda x: (str(x.get("zone", "")), str(x.get("symbol", "")), str(x.get("name_code", ""))))

    print(f"Elementos finales seleccionados: {len(final_rows)}")
    write_output_xls(OUTPUT_XLS, final_rows)
    print(f"OK. Archivo generado: {OUTPUT_XLS}")


if __name__ == "__main__":
    main()
