# ──────────────────────────────────────────────────────────────────────────────
# Imports
# ──────────────────────────────────────────────────────────────────────────────

import os
import re
import psutil
import unicodedata
import xml.etree.ElementTree as ET
import hashlib
import pandas as pd
from ftfy import fix_text
from openpyxl import load_workbook
from openpyxl.styles import numbers
from openpyxl.utils import get_column_letter
from unidecode import unidecode
import win32com.client as win32
from win32com.client import gencache
from Dependencies.Titulos import determinar_titulo
import time
from .models import question, ANSWER_SETS, TITLE_CUES, COLOR_MAP
from .ProcesamientoDeTitulos import clasificar_tipo_spacy

# ──────────────────────────────────────────────────────────────────────────────
#  helpers
# ──────────────────────────────────────────────────────────────────────────────
_VERBOSE = True  # default matches current behavior (prints enabled)

def set_verbose(flag: bool) -> None:
    """Enable/disable module-wide verbose prints."""
    global _VERBOSE
    _VERBOSE = bool(flag)

def _log(msg: str, *, verbose: bool | None = None) -> None:
    """Print only if verbose is True. If None, use module default."""
    v = _VERBOSE if (verbose is None) else verbose
    if v:
        print(msg)

def _quote_sheet_name(name: str) -> str:
    # remove surrounding quotes if present, then rewrap in single quotes
    n = name.strip().strip("'")
    return f"'{n}'"

# ──────────────────────────────────────────────────────────────────────────────
# Pandas display (developer-friendly)
# ──────────────────────────────────────────────────────────────────────────────

pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", 5)
pd.set_option("display.width", None)
pd.set_option("display.max_colwidth", None)


# ──────────────────────────────────────────────────────────────────────────────
# closing programs
# ──────────────────────────────────────────────────────────────────────────────

def close_powerpoint():
    for proc in psutil.process_iter(["name", "pid"]):
        try:
            if proc.info["name"] == "POWERPNT.EXE":
                proc.kill()
                print(f"Closed PowerPoint (PID: {proc.pid})")
        except Exception as e:
            print(f"Failed to close PowerPoint: {e}")


def close_excel_ppt():
    """Kill any stray Excel/PowerPoint processes."""
    for proc in psutil.process_iter(["name", "pid"]):
        try:
            if proc.info["name"] in ("EXCEL.EXE", "POWERPNT.EXE"):
                proc.kill()
                print(f"Closed: {proc.info['name']} (PID: {proc.pid})")
        except Exception as e:
            print(f"Failed to close {proc.info.get('name')}: {e}")



# ──────────────────────────────────────────────────────────────────────────────
# Normalization helpers
# ──────────────────────────────────────────────────────────────────────────────

def normalize_text(x):
    """
    Normaliza texto de forma robusta (UTF-8 seguro):
    corrige glitches de encoding, reemplaza NBSP, colapsa espacios,
    quita diacríticos y devuelve en minúsculas con strip.
    """
    if pd.isna(x) or not isinstance(x, str):
        return x
    s = fix_text(x)
    s = s.replace("\u00A0", " ")
    s = " ".join(s.split())
    s = unidecode(s)
    return s.lower().strip()


def normalize_df(df, columns=None, normalize_headers=True):
    """
    Aplica normalize_text a columnas object. Si normalize_headers=True,
    normaliza también los encabezados.
    """
    out = df.copy()
    target_cols = list(columns) if columns else out.select_dtypes(include=["object"]).columns
    out[target_cols] = out[target_cols].apply(lambda s: s.map(normalize_text))
    if normalize_headers:
        out.columns = [normalize_text(c) if isinstance(c, str) else c for c in out.columns]
    return out

# ──────────────────────────────────────────────────────────────────────────────
# Tiny helpers (NEW)
# ──────────────────────────────────────────────────────────────────────────────

def collapse_ws(s: str) -> str:
    """Collapse tabs/newlines/multiple spaces into single spaces."""
    return " ".join(str(s or "").split())

def split_block_rows(a1_block: str, n_data_rows: int) -> list[str]:
    """
    'Sheet1!$B$12:$BS$17' + n=6 → [
        "'Sheet1'!$B$12:$BS$12",
        "'Sheet1'!$B$13:$BS$13",
        ...
        "'Sheet1'!$B$17:$BS$17"
    ]
    Returns [] on parse errors.
    """
    try:
        sheet_part, rng = a1_block.split("!", 1)
        # normalize sheet quoting to single quotes
        sheet = sheet_part.strip().strip("'")
        left, right = rng.split(":", 1)

        def parse_cell(c: str):
            m = re.match(r"\$?([A-Z]+)\$?(\d+)$", c.strip())
            if not m:
                raise ValueError(f"Bad A1 cell: {c!r}")
            col, row = m.group(1), int(m.group(2))
            return col, row

        c1, r1 = parse_cell(left)
        c2, r2 = parse_cell(right)

        # clamp to the block's bottom row
        last = min(r1 + max(0, n_data_rows - 1), r2)
        sheet_q = f"'{sheet}'"
        return [f"{sheet_q}!${c1}${r}:${c2}${r}" for r in range(r1, last + 1)]
    except Exception:
        return []


# ──────────────────────────────────────────────────────────────────────────────
# Excel I/O and formatting
# ──────────────────────────────────────────────────────────────────────────────

def insert_header(df: pd.DataFrame) -> pd.DataFrame:
    """
    Copia la primera fila y la inserta encima de cada fila con valor no nulo
    en la primera columna (excepto las primeras filas).
    """
    if len(df) < 1:
        return df

    data = df.values.tolist()
    columns = df.columns.tolist()
    header_row = data[0]

    non_null_indices = [
        i for i, row in enumerate(data[1:], start=1)
        if pd.notna(row[0]) and str(row[0]).strip() and i > 1
    ]

    for idx in reversed(non_null_indices):
        data.insert(idx, header_row.copy())

    return pd.DataFrame(data, columns=columns)


def export_to_excel(df: pd.DataFrame, file_save_location: str) -> None:
    df.to_excel(file_save_location, index=False)
    print(f"Data exported to {file_save_location}")


def apply_percentage_format(file_save_location: str) -> None:
    """
    Aplica formato 0% a celdas numéricas (filas de datos).
    """
    app_wb = load_workbook(file_save_location)
    app_ws = app_wb.active

    for row in app_ws.iter_rows(min_row=2, max_row=app_ws.max_row,
                                min_col=1, max_col=app_ws.max_column):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = "0%"

    app_wb.save(file_save_location)
    print("Excel file with percentage formatting has been saved.")


# ──────────────────────────────────────────────────────────────────────────────
# PowerPoint — linked charts and refresh
# ──────────────────────────────────────────────────────────────────────────────

def revincular_cuadros(ppt_template_path: str, excel_save_location: str, visible: bool = False) -> None:
    """
    Actualiza el SourceFullName de objetos vinculados (Type == 3) para que
    apunten al nuevo archivo de Excel. Sin uso de .Visible para builds restrictivas.
    """
    ppt_app = win32.Dispatch("PowerPoint.Application")
    if visible:
        ppt_app.DisplayAlerts = False
        ppt_app.WindowState = 2  # Minimizado

    presentation = ppt_app.Presentations.Open(ppt_template_path,)

    for slide in presentation.Slides:
        for shape in slide.Shapes:
            if getattr(shape, "Type", None) == 3:
                try:
                    source = shape.LinkFormat.SourceFullName
                    print(f"Found linked shape: {shape.Name} → {source}")
                    shape.LinkFormat.SourceFullName = excel_save_location
                    print(f"Updated link for shape: {shape.Name}")
                except Exception as e:
                    print(f"Skipping shape {shape.Name} due to error: {e}")

    presentation.Save()
    print(f"Presentation saved as: {ppt_template_path}")

    presentation.Close()
    ppt_app.Quit()
    print("PowerPoint closed successfully.")


import win32com.client as win32
import pythoncom
from win32com.client import gencache


def refresh_all_graphs(pptx_path: str, with_window=False):
    """
    Open PPTX, trigger recalculation/refresh on linked charts, save, close.
    Accepts a *path*; does not accept a Presentation object.
    Auto-repairs corrupted pywin32 COM cache (common gen_py SyntaxError).
    """
    if not isinstance(pptx_path, str):
        raise TypeError(f"refresh_all_graphs expects a file path (str), got {type(pptx_path)}")

    def safe_dispatch(app_name):
        """Try to dispatch COM app; rebuild cache if corruption detected."""
        try:
            return gencache.EnsureDispatch(app_name)
        except SyntaxError:
            cache_path = tempfile.gettempdir() + "\\gen_py"
            shutil.rmtree(cache_path, ignore_errors=True)
            return gencache.EnsureDispatch(app_name)

    pythoncom.CoInitialize()
    pp = None
    try:
        pp = safe_dispatch("PowerPoint.Application")
        pp.Visible = with_window
        pp.DisplayAlerts = 0

        try:
            pres = pp.Presentations.Open(pptx_path, ReadOnly=0, Untitled=0, WithWindow=with_window)
        except Exception as e:
            raise RuntimeError(
                f"PowerPoint failed to open '{pptx_path}'. "
                "This usually means a corrupted part (often chart XML ranges)."
            ) from e

        # Trigger relinks/refresh:
        for sld in pres.Slides:
            for shp in sld.Shapes:
                if hasattr(shp, "LinkFormat"):
                    try:
                        shp.LinkFormat.Update()
                    except Exception:
                        pass  # some shapes may not refresh cleanly

        pres.Save()
        pres.Close()

    finally:
        if pp:
            try:
                pp.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


# ──────────────────────────────────────────────────────────────────────────────
# Segmentation utilities
# ──────────────────────────────────────────────────────────────────────────────

def find_header_rows(df: pd.DataFrame, question_col: int = 0) -> list[int]:
    """
    Encuentra filas de encabezado asumiendo que se duplicó el encabezado
    inmediatamente por encima de cada fila de pregunta (columna A no vacía).
    Devuelve índices 0-based de encabezados.
    """
    col = df.iloc[:, question_col]
    candidates = [i - 1 for i, v in enumerate(col) if isinstance(v, str) and v.strip()]
    headers = sorted({i for i in candidates if 0 <= i < len(df)})
    return headers


def slice_blocks(header_rows: list[int], nrows: int) -> list[tuple[int, int]]:
    """
    Construye segmentos [start, end) no superpuestos donde start es el encabezado,
    y las filas de datos son start+1 .. end-1.
    """
    if not header_rows:
        return []
    hdrs = sorted(h for h in header_rows if 0 <= h < nrows)
    blocks: list[tuple[int, int]] = []
    for idx, s in enumerate(hdrs):
        e = hdrs[idx + 1] if idx + 1 < len(hdrs) else nrows
        if s + 1 < e:
            blocks.append((s, e))
    return blocks


def build_excel_range(start: int, end: int, data_start_col: int, df_width: int,
                      sheet_name: str = "Sheet1") -> str:
    """
    Construye un rango absoluto de Excel cubriendo sólo las filas de datos del bloque.
    first data row = start+1 (0-based) → Excel = start+2
    last  data row = end-1  (0-based)  → Excel = end
    Columnas: de data_start_col al último índice de columna del DataFrame.
    """
    first_row_1based = start + 2
    last_row_1based = end + 1
    first_col_letter = get_column_letter(data_start_col + 1)
    last_col_letter = get_column_letter(df_width)
    return f"{sheet_name}!${first_col_letter}${first_row_1based}:${last_col_letter}${last_row_1based}"


def segmentar_dataframe(df: pd.DataFrame, verbose: bool = True,
                        question_col: int = 0, data_start_col: int = 1):
    if verbose:
        print(f"Initializing segmentation on {len(df)} rows")
        print(f"Using column {question_col} for question headers")
        print(f"Data starts at column {data_start_col}")
        print(df.head(20))

    header_rows = find_header_rows(df, question_col=question_col)
    if verbose:
        print(f"Found {len(header_rows)} header rows at indices: {header_rows}")
        if not header_rows:
            print("No valid headers found - returning empty list")

    blocks = slice_blocks(header_rows, len(df))

    questions = []
    for i, (s, e) in enumerate(blocks):
        bloque = df.iloc[s:e]
        ubicacion = build_excel_range(s, e, data_start_col, df.shape[1], sheet_name="Sheet1")
        try:
            q = procesar_bloque(bloque, ubicacion)
            if q:
                questions.append(q)
                if verbose:
                    print(f"[block {i}] name: {q.nombre} | rows: {s+1}-{e-1} | range: {ubicacion}")
        except Exception as err:
            if verbose:
                print(f"Failed to process block {i} ({s}-{e}): {err}")

    if verbose:
        print(f"Segmentation complete. Found {len(questions)} valid questions")
    return questions


# ──────────────────────────────────────────────────────────────────────────────
# Block processing and classification
# ──────────────────────────────────────────────────────────────────────────────

ALLOWED_EXTRAS = {"otro", "otra", "otra cual", "otra cual?", "blanco", "nulo", "no recuerdo"}


def Determinar_tipo(respuestas, nombre):
    """
    Determina el tipo de pregunta:
    1) Por pistas en el título (TITLE_CUES).
    2) Por conjuntos de respuestas (ANSWER_SETS): igualdad exacta o subconjunto.
    """
    for tipo, cues in TITLE_CUES.items():
        if any(cue in nombre for cue in cues):
            return tipo

    resp_set = {r for r in respuestas if r not in ALLOWED_EXTRAS}

    for tipo, canon in ANSWER_SETS.items():
        if resp_set == canon:
            return tipo

    candidates = [(tipo, canon) for tipo, canon in ANSWER_SETS.items()
                  if resp_set.issubset(canon)]

    if not candidates:
        return "misc"

    tipo = min(candidates, key=lambda tc: len(tc[1]))[0]
    return tipo


def procesar_bloque(bloque: pd.DataFrame, ubicacion: str):
    print("Buscando el primer texto valido en la primera columna...")
    bloque = bloque.reset_index(drop=True)

    text = next(
        (val for val in bloque.iloc[0:, 0] if isinstance(val, str) and val.strip()),
        "Sin texto"
    )
    print(f"Texto encontrado: {text if text != 'Sin texto' else 'No se encontro texto valido'}")

    print("Extrayendo respuestas de la segunda columna...")
    respuestas = (
        bloque.iloc[:, 1]
        .dropna()
        .apply(str)
        .tolist()
    )
    print(f"Respuestas encontradas: {respuestas}")

    tipo = clasificar_tipo_spacy(respuestas, text)
    Titulo = determinar_titulo(text)

    print("Creando el objeto Cuestion...")
    # IMPORTANT: no-args constructor (fixes 'question() takes no arguments')
    new_question = question()  # ← changed

    # Fill core fields (bit by bit)
    new_question.nombre = text
    new_question.respuestas = respuestas
    new_question.ubicacion = ubicacion
    new_question.titulo = collapse_ws(Titulo)
    new_question.tipo = tipo

    # Fill new manifest-style fields only if the class provides them
    # 1) Excel block + per-row ranges
    if hasattr(new_question, "excel_range_block"):
        new_question.excel_range_block = ubicacion
    if hasattr(new_question, "excel_ranges_per_row"):
        n_data_rows = max(0, len(bloque) - 1)  # header + data rows
        new_question.excel_ranges_per_row = split_block_rows(ubicacion, n_data_rows)
        print("ranges per row is : ", new_question.excel_ranges_per_row)

    # 2) Legends (original + normalized)
    if hasattr(new_question, "legends_original"):
        new_question.legends_original = list(respuestas)
    if hasattr(new_question, "legends_norm"):
        new_question.legends_norm = [normalize_text(x) for x in respuestas]

    # 3) Palette hint (best effort)
    if hasattr(new_question, "palette_hint"):
        new_question.palette_hint = COLOR_MAP.get(tipo) if isinstance(COLOR_MAP, dict) else None

    # 4) Stable id
    if hasattr(new_question, "id"):
        try:
            new_question.id = compute_stable_id(text, ubicacion, getattr(new_question, "legends_norm", []))
        except Exception:
            pass

    print("cuestion.nombre is", new_question.nombre)
    print("cuestion.respuestas is", new_question.respuestas)
    print("cuestion.ubicacion is", new_question.ubicacion)
    print("question.titulo is", new_question.titulo)
    print("Retornando el objeto Cuestion...")
    return new_question

# ──────────────────────────────────────────────────────────────────────────────
# A1 range utilities
# ──────────────────────────────────────────────────────────────────────────────

_A1_ROW_RANGE = re.compile(r"^\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+$")

def _extract_sheet_from_ubicacion(ubicacion: str) -> str:
    sheet = ubicacion.split("!", 1)[0].strip()
    return _quote_sheet_name(sheet)

def _anchor_cols_rows(a1: str) -> str:
    def _part(p):
        m = re.match(r"^\$?([A-Za-z]{1,3})\$?(\d+)$", p)
        col, row = m.group(1), m.group(2)
        return f"${col.upper()}${row}"
    left, right = a1.split(":")
    return f"{_part(left)}:{_part(right)}"

def normalize_a1_row_range(rng: str, ubicacion_full: str) -> str:
    """
    Turn 'B44:BS44' into  ''Hoja1'!$B$44:$BS$44'
    If rng already has a sheet, it’s normalized (sheet quoted + anchors added).
    """
    if "!" in rng:  # already qualified
        sheet, rest = rng.split("!", 1)
        sheet = _quote_sheet_name(sheet)
        rest = rest.replace("$", "").strip()
        rest = _anchor_cols_rows(rest)
        return f"{sheet}!{rest}"

    sheet = _extract_sheet_from_ubicacion(ubicacion_full)
    rng_no_dollar = rng.replace("$", "").strip()
    anchored = _anchor_cols_rows(rng_no_dollar)
    return f"{sheet}!{anchored}"