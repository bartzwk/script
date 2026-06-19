"""
find_text_in_pdfs.py
--------------------
Busca uno o varios textos en PDFs dentro de las carpetas "Workshop drawings"
de cada sección del proyecto FA1.

Uso:
    python find_text_in_pdfs.py "texto1" "texto2" "texto3"

Si no se dan argumentos en línea de comandos, se usa la lista SEARCH_TEXTS
definida en CONFIG más abajo.

Salida:
    Consola  + fichero TXT  →  find_results_<timestamp>.txt
"""

import logging
import os
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import List

logging.getLogger("pypdf").setLevel(logging.ERROR)

# Force UTF-8 stdout so Polish paths (e.g. "Popławski") don't crash on
# Windows when output is redirected/piped (default cp1252 can't encode 'ł').
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


# =========================================================
# TERMINAL HYPERLINKS (OSC 8)
# Clickable file links in Windows Terminal / VS Code terminal.
# Clicking opens the PDF with the OS default app. Disabled when the
# output is not a TTY (redirected/piped) so escape codes don't leak.
# =========================================================

def _enable_windows_ansi() -> None:
    """Enable ANSI/VT escape processing on the Windows console (no-op elsewhere)."""
    if os.name != "nt":
        return
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
        for handle_id in (-11, -12):  # STDOUT, STDERR
            handle = kernel32.GetStdHandle(handle_id)
            mode = ctypes.c_uint32()
            if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
                kernel32.SetConsoleMode(handle, mode.value | ENABLE_VIRTUAL_TERMINAL_PROCESSING)
    except Exception:
        pass


_enable_windows_ansi()
USE_HYPERLINKS = sys.stdout.isatty()


def file_link(path: str, label: str = None) -> str:
    """Return an OSC 8 hyperlink to a file, or the plain path when not a TTY."""
    text = label or path
    if not USE_HYPERLINKS:
        return text
    try:
        uri = Path(path).resolve().as_uri()
    except Exception:
        return text
    esc = "\x1b"
    return f"{esc}]8;;{uri}{esc}\\{text}{esc}]8;;{esc}\\"

try:
    from pypdf import PdfReader
except Exception:
    print("ERROR: instala pypdf primero:  pip install pypdf")
    raise

# =========================================================
# CONFIG — edita aquí si no usas argumentos de línea de comandos
# =========================================================
PROJECT_CONFIGS = {
    "FA1": {
        "root_dir":       r"P:\Aleksander Popławski 2\131-FA1\1. Dokumentacja projektowa",
        "search_folders": ["A1", "A2", "B", "C1", "C2"],
    },
    "CP1": {
        "root_dir":       r"P:\Aleksander Popławski 2\121-CP1\1. Dokumentacja projektowa",
        "search_folders": ["A", "A2", "B", "C"],
    },
}
DEFAULT_PROJECT = "FA1"

# Nombre (o fragmento) de la subcarpeta de planos dentro de cada sección.
# El script hace una búsqueda case-insensitive, así que no importan mayúsculas.
WORKSHOP_DRAWINGS_SUBFOLDER = "workshop drawings"

# Textos a buscar cuando NO se pasan argumentos CLI
SEARCH_TEXTS: List[str] = [
    # "PL-101",
    # "HEA 200",
]

# Contexto mostrado alrededor de cada coincidencia (caracteres)
SHOW_SNIPPET_CHARS = 120

# Carpeta donde se guarda el TXT de resultados (None → ./out junto a este script)
OUTPUT_DIR = None
# Carpeta de salida por defecto: ./out junto a este script
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
# =========================================================

# Unión de secciones de todos los proyectos — usado para detección de zona en queries
ALL_SECTIONS = sorted({s for cfg in PROJECT_CONFIGS.values() for s in cfg["search_folders"]})


# ─── helpers de texto ────────────────────────────────────

def strip_accents(text: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", str(text))
        if not unicodedata.combining(ch)
    )


def normalize(text: str) -> str:
    s = str(text or "")
    s = s.replace("\xa0", " ")
    s = strip_accents(s)
    s = s.lower()
    # Unify separators so codes match regardless of style:
    # 'CP1_B_GT_900' (Excel) == 'CP1-B-GT-900' (drawings) == 'CP1 B GT 900'
    s = re.sub(r"[_\-/]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# Códigos de empresa reconocidos como separadores de carpeta
COMPANY_CODES = {"GT", "KG"}


# ─── scope inference ─────────────────────────────────────

def infer_scope_from_queries(queries: List[str]):
    """
    Detecta zona (A1/A2/B/C1/C2) y empresa (GT/KG) en los textos buscados.
    Ejemplo: 'A2_GT_L-534'  →  zones=['A2'], companies=['GT']
    Devuelve listas vacías si no hay pistas → búsqueda completa.
    """
    found_zones: set = set()
    found_companies: set = set()
    for q in queries:
        tokens = {t.upper() for t in re.split(r"[_\-\s]+", q) if t}
        for zone in ALL_SECTIONS:
            if zone.upper() in tokens:
                found_zones.add(zone.upper())
        for company in COMPANY_CODES:
            if company in tokens:
                found_companies.add(company)
    return sorted(found_zones), sorted(found_companies)


# ─── recolección de PDFs ─────────────────────────────────

def collect_pdf_files(
    root_dir: str,
    sections_filter: List[str] = None,
    company_hints: List[str] = None,
) -> List[str]:
    """
    Devuelve todos los PDFs bajo
    <root_dir>/<section>/<Workshop drawings>/...

    sections_filter  — si se da, sólo escanea esas secciones.
    company_hints    — filtra carpetas de empresa: si un directorio se
                       identifica como carpeta de empresa (su nombre contiene
                       un código de COMPANY_CODES) pero NO coincide con los
                       hints, se poda. Carpetas estructurales ('Ceilings',
                       'Revision 0', 'Part 1'…) siempre se entran.
    """
    sections = sections_filter or ALL_SECTIONS
    hints = [c.lower() for c in (company_hints or [])]
    all_codes = {c.lower() for c in COMPANY_CODES}
    needle = WORKSHOP_DRAWINGS_SUBFOLDER.lower()
    files: List[str] = []

    for section in sections:
        section_dir = os.path.join(root_dir, section)
        if not os.path.isdir(section_dir):
            print(f"[WARN] No existe: {section_dir}")
            continue

        workshop_dir = None
        for entry in os.scandir(section_dir):
            if entry.is_dir() and needle in entry.name.lower():
                workshop_dir = entry.path
                break

        if workshop_dir is None:
            print(f"[WARN] No se encontró '{WORKSHOP_DRAWINGS_SUBFOLDER}' en {section_dir}")
            continue

        print(f"[SCAN] {workshop_dir}")
        for dirpath, dirnames, filenames in os.walk(workshop_dir):
            if hints:
                def _keep(d: str) -> bool:
                    dl = d.lower()
                    is_company_dir = any(code in dl for code in all_codes)
                    if is_company_dir:
                        return any(h in dl for h in hints)
                    return True  # carpeta estructural → siempre entrar

                dirnames[:] = [d for d in dirnames if _keep(d)]

            for fname in filenames:
                if fname.startswith("~$") or not fname.lower().endswith(".pdf"):
                    continue
                files.append(os.path.join(dirpath, fname))

    return files


# ─── extracción de texto ──────────────────────────────────

def extract_pages(pdf_path: str) -> List[str]:
    reader = PdfReader(pdf_path)
    pages = []
    for i, page in enumerate(reader.pages):
        try:
            pages.append(page.extract_text() or "")
        except Exception as exc:
            print(f"[ERROR] {pdf_path} | pág {i+1} | {exc}")
            pages.append("")
    return pages


def _is_token_char(ch: str) -> bool:
    return ch.isalnum() or ch == "_"


def find_all(haystack: str, needle: str) -> List[int]:
    """
    Returns all start positions of needle in haystack.
    If needle ends with an alphanumeric/underscore character, a match is only
    accepted when the following character is NOT alphanumeric/underscore — so
    'A1_GT_700' won't match inside 'A1_GT_7008'.
    """
    starts, pos = [], 0
    n = len(needle)
    check_right = bool(needle) and _is_token_char(needle[-1])
    while True:
        idx = haystack.find(needle, pos)
        if idx == -1:
            break
        after = idx + n
        if not check_right or after >= len(haystack) or not _is_token_char(haystack[after]):
            starts.append(idx)
        pos = idx + 1
    return starts


def snippet(original: str, norm_query: str) -> str:
    if not original:
        return ""
    norm_orig = normalize(original)
    idx = norm_orig.find(norm_query)
    half = SHOW_SNIPPET_CHARS // 2
    if idx == -1:
        raw = original[:SHOW_SNIPPET_CHARS]
    else:
        start = max(0, idx - half)
        end = min(len(original), idx + len(norm_query) + half)
        raw = ("... " if start > 0 else "") + original[start:end] + (" ..." if end < len(original) else "")
    return re.sub(r"\s+", " ", raw).strip()


# ─── búsqueda principal ───────────────────────────────────

def search_all(root_dir: str, sections: List[str], queries: List[str]) -> str:
    if not queries:
        print("ERROR: no hay textos para buscar.")
        return ""

    zones, companies = infer_scope_from_queries(queries)
    if zones or companies:
        print()
        if zones:
            print(f"  [SCOPE] Zona(s) detectada(s)  : {', '.join(zones)}")
        if companies:
            print(f"  [SCOPE] Empresa(s) detectada(s): {', '.join(companies)}")

    pdf_files = collect_pdf_files(
        root_dir,
        sections_filter=zones if zones else sections,
        company_hints=companies or None,
    )
    total = len(pdf_files)
    print(f"\nPDFs encontrados: {total}\n")

    if not pdf_files:
        msg = "No se encontraron PDFs en las carpetas configuradas."
        print(msg)
        return msg

    results: dict = {normalize(q): {} for q in queries}
    query_map = {normalize(q): q for q in queries}

    PROGRESS_WIDTH = 70

    for i, pdf_path in enumerate(pdf_files, start=1):
        short = os.path.basename(pdf_path)
        progress = f"  [{i}/{total}] {short}"
        print(f"\r{progress:<{PROGRESS_WIDTH}}", end="", flush=True)

        try:
            pages = extract_pages(pdf_path)
        except Exception as exc:
            print(f"\r[ERROR] {pdf_path} -> {exc}")
            continue

        file_had_match = False

        for q_norm, q_orig in query_map.items():
            hits_by_page = []
            for page_idx, page_text in enumerate(pages, start=1):
                page_norm = normalize(page_text)
                positions = find_all(page_norm, q_norm)
                if positions:
                    hits_by_page.append({
                        "page": page_idx,
                        "count": len(positions),
                        "snippet": snippet(page_text, q_norm),
                    })

            if hits_by_page:
                results[q_norm][pdf_path] = hits_by_page
                if not file_had_match:
                    print(f"\r{' ' * PROGRESS_WIDTH}\r", end="")
                    print(f"  [MATCH] {file_link(pdf_path)}")
                    file_had_match = True
                file_hits = sum(h["count"] for h in hits_by_page)
                print(f'          "{q_orig}" — {file_hits} coincidencia(s)')
                for h in hits_by_page:
                    print(f"          - Pág {h['page']}: {h['count']} vez/veces")
                    if h["snippet"]:
                        print(f"            {h['snippet']}")

    print(f"\r{' ' * PROGRESS_WIDTH}\r", end="")
    print(f"  Búsqueda completada. {total} PDFs analizados.\n")

    # ── informe para el fichero TXT ──
    lines: List[str] = []
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines.append("=" * 100)
    lines.append(f"  BÚSQUEDA EN PDFs — {timestamp}")
    lines.append(f"  Carpeta raíz : {root_dir}")
    lines.append(f"  PDFs analizados: {total}")
    lines.append("=" * 100)

    for q_norm, q_orig in query_map.items():
        matched = results[q_norm]
        total_hits = sum(sum(h["count"] for h in hits) for hits in matched.values())

        lines.append("")
        lines.append("─" * 100)
        lines.append(f'  TEXTO: "{q_orig}"')
        lines.append(f"  Ficheros con coincidencias: {len(matched)}   |   Coincidencias totales: {total_hits}")
        lines.append("─" * 100)

        if not matched:
            lines.append("  (sin resultados)")
        else:
            for pdf_path, hits in sorted(matched.items()):
                file_hits = sum(h["count"] for h in hits)
                lines.append("")
                lines.append(f"  [MATCH] {pdf_path}")
                lines.append(f"          Total coincidencias: {file_hits}")
                for h in hits:
                    lines.append(f"          - Pág {h['page']}: {h['count']} vez/veces")
                    if h["snippet"]:
                        lines.append(f"            {h['snippet']}")

    lines.append("")
    lines.append("=" * 100)
    return "\n".join(lines)


# ─── entrada ─────────────────────────────────────────────

if __name__ == "__main__":
    _args = sys.argv[1:]
    _project = DEFAULT_PROJECT
    if _args and _args[0].upper() in PROJECT_CONFIGS:
        _project = _args[0].upper()
        _args = _args[1:]

    _cfg = PROJECT_CONFIGS[_project]
    queries = [q.strip() for q in _args if q.strip()] or SEARCH_TEXTS

    if not queries:
        print(__doc__)
        print("ERROR: proporciona al menos un texto de búsqueda.")
        print(f"Proyectos disponibles: {', '.join(PROJECT_CONFIGS)}")
        sys.exit(1)

    print(f"\n  Proyecto: {_project}  ({_cfg['root_dir']})")
    report = search_all(_cfg["root_dir"], _cfg["search_folders"], queries)

    # Guardar TXT
    out_dir = OUTPUT_DIR or OUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    timestamp_file = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(out_dir, f"find_results_{timestamp_file}.txt")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"\n[GUARDADO] {file_link(out_path)}")
