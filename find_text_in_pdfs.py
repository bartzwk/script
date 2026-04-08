import os
import re
import sys
import unicodedata
from pathlib import Path
from typing import List, Tuple

try:
    from pypdf import PdfReader
except Exception:
    print("ERROR: falta instalar/usar pypdf en este entorno.")
    raise

# =========================================================
# CONFIG
# =========================================================
ROOT_DIR = r"C:\Users\Bartek\Desktop\GEMTECH\CP1"
SEARCH_TOP_FOLDERS = ["B", "C"]
EXCLUDE_SUMMARY_FOLDERS = True
SHOW_SNIPPET_CHARS = 120


# =========================================================
# HELPERS
# =========================================================
def strip_accents(text: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", str(text))
        if not unicodedata.combining(ch)
    )


def normalize_text(text: str) -> str:
    s = str(text or "")
    s = s.replace("\xa0", " ")
    s = strip_accents(s)
    s = s.lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def is_summary_company_folder(path_parts_lower: List[str]) -> bool:
    for p in path_parts_lower:
        p_norm = normalize_text(p)
        if p_norm in {"kg+gt", "kg i gt section", "kg i gt", "kg&gt", "gt+kg"}:
            return True
    return False


def should_scan_pdf(full_path: str, root_dir: str) -> Tuple[bool, str]:
    rel_path = os.path.relpath(full_path, root_dir)
    parts = Path(rel_path).parts
    parts_lower = [p.lower() for p in parts]
    base_name = os.path.basename(full_path)

    if len(parts) < 2:
        return False, "not_inside_B_or_C_subfolder"

    if parts[0].upper() not in SEARCH_TOP_FOLDERS:
        return False, "outside_B_C"

    if not base_name.lower().endswith(".pdf"):
        return False, "not_pdf"

    if base_name.startswith("~$"):
        return False, "temp_file"

    if EXCLUDE_SUMMARY_FOLDERS and is_summary_company_folder(parts_lower):
        return False, "summary_company_folder"

    return True, "ok"


def collect_pdf_files(root_dir: str) -> List[str]:
    files: List[str] = []

    for top in SEARCH_TOP_FOLDERS:
        top_dir = os.path.join(root_dir, top)
        print(f"[SCAN] {top_dir}")

        if not os.path.isdir(top_dir):
            print(f"[WARN] No existe carpeta: {top_dir}")
            continue

        for dirpath, _, filenames in os.walk(top_dir):
            for filename in filenames:
                full_path = os.path.join(dirpath, filename)

                if not filename.lower().endswith(".pdf"):
                    continue

                print(f"[PDF] {full_path}")
                ok, reason = should_scan_pdf(full_path, root_dir)
                if not ok:
                    print(f"[SKIP:{reason}] {full_path}")
                    continue

                print(f"[USE] {full_path}")
                files.append(full_path)

    return files


def extract_text_from_pdf(pdf_path: str) -> List[str]:
    reader = PdfReader(pdf_path)
    pages_text: List[str] = []

    for i, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
        except Exception as exc:
            print(f"[ERROR] {pdf_path} | página {i+1} | {exc}")
            text = ""
        pages_text.append(text)

    return pages_text


def find_all_occurrences(normalized_haystack: str, normalized_needle: str) -> List[int]:
    starts: List[int] = []
    start = 0
    while True:
        idx = normalized_haystack.find(normalized_needle, start)
        if idx == -1:
            break
        starts.append(idx)
        start = idx + 1
    return starts


def make_snippet(original_text: str, normalized_query: str) -> str:
    if not original_text:
        return ""

    normalized_text = normalize_text(original_text)
    idx = normalized_text.find(normalized_query)
    if idx == -1:
        snippet = original_text[:SHOW_SNIPPET_CHARS]
        return re.sub(r"\s+", " ", snippet).strip()

    # Aproximación: usar misma posición sobre texto original.
    start = max(0, idx - SHOW_SNIPPET_CHARS // 2)
    end = min(len(original_text), idx + len(normalized_query) + SHOW_SNIPPET_CHARS // 2)
    snippet = original_text[start:end]
    snippet = re.sub(r"\s+", " ", snippet).strip()
    if start > 0:
        snippet = "... " + snippet
    if end < len(original_text):
        snippet = snippet + " ..."
    return snippet


def search_in_pdfs(root_dir: str, query: str) -> None:
    query_norm = normalize_text(query)
    if not query_norm:
        print("ERROR: el texto de búsqueda está vacío.")
        return

    pdf_files = collect_pdf_files(root_dir)
    print()
    print(f"PDF válidos encontrados: {len(pdf_files)}")

    if not pdf_files:
        print("No se han encontrado PDFs válidos en B y C.")
        return

    total_hits = 0
    matched_files = 0

    print()
    print("=" * 100)
    print(f"BUSCANDO: {query}")
    print("=" * 100)

    for pdf_path in pdf_files:
        rel_path = os.path.relpath(pdf_path, root_dir)
        pages_text = extract_text_from_pdf(pdf_path)

        file_hits = []
        for page_index, page_text in enumerate(pages_text, start=1):
            page_norm = normalize_text(page_text)
            if not page_norm:
                continue

            starts = find_all_occurrences(page_norm, query_norm)
            if starts:
                file_hits.append({
                    "page": page_index,
                    "count": len(starts),
                    "snippet": make_snippet(page_text, query_norm),
                })

        if file_hits:
            matched_files += 1
            file_total = sum(hit["count"] for hit in file_hits)
            total_hits += file_total

            print()
            print(f"[MATCH] {rel_path}")
            print(f"        Coincidencias totales: {file_total}")
            for hit in file_hits:
                print(f"        - Página {hit['page']}: {hit['count']} coincidencia(s)")
                if hit["snippet"]:
                    print(f"          {hit['snippet']}")

    print()
    print("=" * 100)
    print(f"Ficheros con coincidencias: {matched_files}")
    print(f"Coincidencias totales: {total_hits}")
    print("=" * 100)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso:")
        print('  python find_text_in_pdfs.py "texto a buscar"')
        sys.exit(1)

    query = " ".join(sys.argv[1:]).strip()
    search_in_pdfs(ROOT_DIR, query)
