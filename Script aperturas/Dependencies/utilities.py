import pandas as pd
import win32com.client as win32
from openpyxl import load_workbook
from openpyxl.styles import numbers
from openpyxl.utils import get_column_letter
import psutil
import re
import unicodedata
from unidecode import unidecode
from ftfy import fix_text
import xml.etree.ElementTree as ET

# Import models

from .models import question, ANSWER_SETS, TITLE_CUES

# Display options to show all data
pd.set_option('display.max_rows', None)  # Show all rows
pd.set_option('display.max_columns', 5)  # Show all columns
pd.set_option('display.width', None)  # Auto-detect terminal width
pd.set_option('display.max_colwidth', None)  # Show full column content

# Recopilador de encabezado (process DataFrame headers)


def insert_header(df):
    """
    Copies the first row and inserts it above every row with non-null values in the first column.
    
    Args:
        df (pd.DataFrame): Input DataFrame with a header row
        
    Returns:
        pd.DataFrame: Modified DataFrame with inserted header rows
    """
    if len(df) < 1:
        return df
    
    # Convert DataFrame to list of lists for easier manipulation
    data = df.values.tolist()
    columns = df.columns.tolist()
    
    # Get first row to copy (header)
    header_row = data[0]
    
    # Find indices where first column is not null (skip first row)
    non_null_indices = [
        i for i, row in enumerate(data[1:], start=1) 
        if pd.notna(row[0]) and str(row[0]).strip() and i > 1
    ]
    
    # Insert header above each non-null row (working backwards to preserve indices)
    for idx in reversed(non_null_indices):
        data.insert(idx, header_row.copy())
    
    # Convert back to DataFrame
    new_df = pd.DataFrame(data, columns=columns)
    return new_df

# Export DataFrame to Excel
def export_to_excel(df, file_save_location):
    df.to_excel(file_save_location, index=False)
    print(f"Data exported to {file_save_location}")

# Apply percentage formatting to Excel file
def apply_percentage_format(file_save_location):
    app_wb = load_workbook(file_save_location)
    app_ws = app_wb.active

    for row in app_ws.iter_rows(min_row=2, max_row=app_ws.max_row, min_col=1, max_col=app_ws.max_column):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = '0%'  # Format as integer percentage (no decimals)

    app_wb.save(file_save_location)
    print("Excel file with percentage formatting has been saved.")

# Revinculado_Cuadros (Linked charts processing)
def revincular_cuadros(ppt_template_path, excel_save_location):
    ppt_app = win32.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True
    presentation = ppt_app.Presentations.Open(ppt_template_path)

    for slide in presentation.Slides:
        for shape in slide.Shapes:
            if shape.Type == 3:  # Picture
                try:
                    source = shape.LinkFormat.SourceFullName
                    print(f"Found linked shape: {shape.Name} -> {source}")
                    shape.LinkFormat.SourceFullName = excel_save_location
                    print(f"Updated link for shape: {shape.Name}")
                except Exception as e:
                    print(f"Skipping shape {shape.Name} due to error: {e}")
    # Save and close if needed
    # presentation.Save()
    # ppt_app.Quit()

# Segment DataFrame based on rows and return list of questions (cuestion objects)

def segmentar_dataframe(df: pd.DataFrame, verbose: bool = True, question_col: int = 0, data_start_col: int = 1):
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

# Process each block to create a cuestion object
def procesar_bloque(bloque, ubicacion):
    """
    Procesa un bloque del DataFrame y devuelve un objeto Cuestion.
    El bloque se supone que contiene dos columnas: la primera con texto y la segunda con respuestas.
    """
    
    # Buscar el primer texto valido en la primera columna
    print("Buscando el primer texto valido en la primera columna...")
    bloque = bloque.reset_index(drop=True)
    
    text = next(
        (val for val in bloque.iloc[0:, 0] if isinstance(val, str) and val.strip()), 
        "Sin texto"
    )
    print(f"Texto encontrado: {text if text != 'Sin texto' else 'No se encontro texto valido'}")
    
    # Extraer respuestas de la segunda columna
    print("Extrayendo respuestas de la segunda columna...")
    respuestas = (
        bloque.iloc[:,1]
            .dropna()
            .apply(str)
            .tolist()
    )
    print(f"Respuestas encontradas: {respuestas}")

    tipo = Determinar_tipo(respuestas, text)
    
    # Crear objeto question con el nombre, respuestas, bloque y ubicacion
    print("Creando el objeto Cuestion...")
    new_question = question(text, respuestas, bloque, ubicacion, tipo)

    new_question.nombre = text
    new_question.respuestas = respuestas
    new_question.ubicacion = ubicacion
    
    # Debug: Imprimir atributos del objeto question para verificar
    print("cuestion.nombre is", new_question.nombre)
    print("cuestion.respuestas is", new_question.respuestas)
    print("cuestion.ubicacion is", new_question.ubicacion)

    # Tipo aun no definido, se asigna "na" por defecto
    print("Asignando tipo 'na' a la pregunta (aun no definido)...")
    new_question.tipo = tipo
    
    # Retornar el objeto cuestion
    print("Retornando el objeto Cuestion...")
    return new_question

def reformatear_ubicacion(quest):
    
    if "!" in quest.ubicacion:  # Skip if already formatted
        return quest.ubicacion
        
    # Add $ for absolute references and handle sheet name
    formatted = f"Sheet1!${quest.ubicacion.replace(':', ':$')}"
    
    # Ensure no duplicate $ (e.g., "B$21" → "$B$21")
    formatted = formatted.replace("$$", "$")
    
    quest.ubicacion = formatted
    
ALLOWED_EXTRAS = {"otro", "otra", "otra cual", "otra cual?", "blanco", "nulo", "no recuerdo"}

def Determinar_tipo(respuestas, nombre):
    # 1) TITLE CUES FIRST
    for tipo, cues in TITLE_CUES.items():
        if any(cue in nombre for cue in cues):
            return tipo

    # 2) ANSWER SETS
    resp_set = {r for r in respuestas if r not in ALLOWED_EXTRAS}

    # Exact equality
    for tipo, canon in ANSWER_SETS.items():
        if resp_set == canon:
            return tipo

    # Subset check (all responses must belong to a canonical set)
    candidates = [(tipo, canon) for tipo, canon in ANSWER_SETS.items()
                if resp_set.issubset(canon)]

    if not candidates:
        return "misc"

    # Pick the smallest superset (most specific classification)
    tipo = min(candidates, key=lambda tc: len(tc[1]))[0]
    return tipo


def close_excel_ppt():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] in ('EXCEL.EXE', 'POWERPNT.EXE'):
            try:
                proc.kill()
                print(f"Closed: {proc.info['name']} (PID: {proc.pid})")
            except Exception as e:
                print(f"Failed to close {proc.info['name']}: {e}")
    
def close_powerpoint():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'POWERPNT.EXE':
            try:
                proc.kill()
                print(f" Closed PowerPoint (PID: {proc.pid})")
            except Exception as e:
                print(f" Failed to close PowerPoint: {e}")


def normalize_text(x):
    if pd.isna(x) or not isinstance(x, str):
        return x
    s = fix_text(x)              # fix encoding glitches
    s = s.replace('\u00A0', ' ') # NBSP → space
    s = ' '.join(s.split())      # collapse whitespace
    s = unidecode(s)             # strip all diacritics (ñ→n, a→a, ç→c, …)
    return s.lower().strip()

def normalize_df(df, columns=None, normalize_headers=True):
    out = df.copy()
    target_cols = list(columns) if columns else out.select_dtypes(include=['object']).columns
    out[target_cols] = out[target_cols].apply(lambda s: s.map(normalize_text))
    if normalize_headers:
        out.columns = [normalize_text(c) if isinstance(c, str) else c for c in out.columns]
    return out

##################### SCRATCHBOOK #######################

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
for k, v in NS.items():
    ET.register_namespace(k, v)

def find_header_rows(df: pd.DataFrame, question_col: int = 0) -> list[int]:
    """
    Find header rows assuming you duplicated the header immediately ABOVE
    every question row (non-empty in column A).
    Returns 0-based indices of header rows.
    """
    col = df.iloc[:, question_col]
    candidates = [i - 1 for i, v in enumerate(col) if isinstance(v, str) and v.strip()]
    headers = sorted({i for i in candidates if 0 <= i < len(df)})
    return headers

def slice_blocks(header_rows: list[int], nrows: int) -> list[tuple[int, int]]:
    """
    Build non-overlapping [start, end) slices where 'start' is the header row,
    and data rows are start+1 .. end-1.
    """
    if not header_rows:
        return []
    hdrs = sorted(h for h in header_rows if 0 <= h < nrows)
    blocks: list[tuple[int, int]] = []
    for idx, s in enumerate(hdrs):
        e = hdrs[idx + 1] if idx + 1 < len(hdrs) else nrows
        if s + 1 < e:  # must contain at least one data row
            blocks.append((s, e))
    return blocks

def build_excel_range(start: int, end: int, data_start_col: int, df_width: int, sheet_name: str = "Sheet1") -> str:
    """
    Build an absolute Excel range covering ONLY the data rows of the block:
    first data row = start+1 (0-based) → Excel = start+2
    last data row  = end-1 (0-based)   → Excel = end
    Columns: from data_start_col to the last column of the DataFrame.
    """
    first_row_1based = start + 2
    last_row_1based = end + 1
    first_col_letter = get_column_letter(data_start_col + 1)
    last_col_letter = get_column_letter(df_width)
    return f"{sheet_name}!${first_col_letter}${first_row_1based}:${last_col_letter}${last_row_1based}"
