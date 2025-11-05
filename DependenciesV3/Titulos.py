#============================================================================
                ###DETERMINAR TITULO###
#============================================================================


import re
from unidecode import unidecode

# Titulos.py
import pandas as pd

XLSX_PATH = r"C:\Users\astor\Desktop\agora\Clientes\Olava post electoral noviembre\Apps\Titulos.xlsx"  # ← ajustá si hace falta

def determinar_titulo(nombre, tipo=None):
    """
    Busca 'nombre' en la columna 'pregunta' de Titulos.xlsx
    y devuelve el contenido de la columna 'titulo' correspondiente.
    Sin tolerancia de errores ni normalización.
    """
    df = pd.read_excel(XLSX_PATH, engine="openpyxl")
    return df.loc[df["Preguntas"] == nombre, "Titulos"].iloc[0]


#============================================================================
                ###CAMBIAR TITULO###
#============================================================================

def _bold_before_pipe(text, shape):
    """
    Applies bold formatting to all text before and including the first '|' in the shape.
    """
    try:
        if "|" not in text:
            return
        before, sep, after = text.partition("|")
        rng = shape.TextFrame.TextRange
        full_text = before + sep + after
        rng.Text = full_text
        # Bold up to and including the '|'
        rng.Characters(1, len(before) + len(sep)).Font.Bold = True
        # Unbold the rest
        rng.Characters(len(before) + len(sep) + 1, len(full_text) - len(before) - len(sep)).Font.Bold = False
    except Exception as e:
        print(f"Bold formatting failed: {e}")


# PowerPoint COM helpers (win32com)
MSO_TYPE_GROUP = 6           # msoGroup
MSO_TRUE = -1
MSO_PLACEHOLDER_TITLE = 1
MSO_PLACEHOLDER_CENTERTITLE = 3

HEADER_CANDIDATE_TERMS = ("plantilla", "titulo", "título")

def _iter_text_shapes(coll):
    """Yield text-capable shapes, descending into groups."""
    for shp in coll:
        try:
            if getattr(shp, "Type", None) == MSO_TYPE_GROUP and hasattr(shp, "GroupItems"):
                for sub in _iter_text_shapes(shp.GroupItems):
                    yield sub
                continue
            if getattr(shp, "HasTextFrame", 0):
                yield shp
        except Exception:
            continue

def _shape_text(shp):
    try:
        return (shp.TextFrame.TextRange.Text or "").strip()
    except Exception:
        return ""

def _is_visible(shp):
    try:
        return getattr(shp, "Visible", MSO_TRUE) == MSO_TRUE
    except Exception:
        return True

def _find_formal_title(slide):
    try:
        t = slide.Shapes.Title
        if t and getattr(t, "HasTextFrame", 0):
            return t
    except Exception:
        pass
    # scan for placeholder types
    for shp in _iter_text_shapes(slide.Shapes):
        try:
            pf = getattr(shp, "PlaceholderFormat", None)
            if pf and getattr(pf, "Type", None) in (MSO_PLACEHOLDER_TITLE, MSO_PLACEHOLDER_CENTERTITLE):
                return shp
        except Exception:
            continue
    return None

def _find_header_by_text(slide):
    """Look for a textbox whose text matches known header terms."""
    best = None
    for shp in _iter_text_shapes(slide.Shapes):
        if not _is_visible(shp): 
            continue
        txt = _shape_text(shp).lower()
        if not txt:
            continue
        if any(term in txt for term in HEADER_CANDIDATE_TERMS):
            best = shp
            break
    return best

def _find_header_by_position(slide, top_band_ratio=0.28):
    """Pick a big visible textbox in the top band (typical header zone)."""
    h = getattr(slide, "Master", None)
    # slide dimensions are not directly available here; rely on relative band
    candidates = []
    for shp in _iter_text_shapes(slide.Shapes):
        try:
            if not _is_visible(shp): 
                continue
            # Skip obvious footers/number/date
            nm = (getattr(shp, "Name", "") or "").lower()
            if any(k in nm for k in ("footer", "date", "slide number", "page number")):
                continue
            if not _shape_text(shp):  # needs to be text-y already
                continue
            # Heuristic: “top band”
            top = float(getattr(shp, "Top", 0))
            height = float(getattr(shp, "Height", 0))
            width = float(getattr(shp, "Width", 0))
            if width < 80 or height < 20:
                continue
            # Use Top relative to the tallest shape area (we don’t know page height; assume titles are near top)
            if top <= (getattr(slide, "SlideShowTransition", None) and 0) or top >= 10**6:
                # ignore weird values
                pass
            # Keep candidates; we’ll sort by (is higher, width)
            candidates.append((top, width, shp))
        except Exception:
            continue

    if not candidates:
        return None

    # Prefer highest (smallest Top), then widest
    candidates.sort(key=lambda t: (t[0], -t[1]))
    # Optionally clamp to a top band (first quartile-ish)
    top_min = candidates[0][0]
    band_threshold = top_min + (candidates[-1][0] - top_min) * top_band_ratio
    for top, width, shp in candidates:
        if top <= band_threshold:
            return shp
    # fallback to absolute best
    return candidates[0][2]

def _set_text(shp, text):
    try:
        shp.TextFrame.TextRange.Text = text
        return True
    except Exception:
        return False

def change_slide_title(slide, nombre, tipo, *, prefer_text_header=True):
    """
    Robustly change the slide's top-left 'header' (like 'PLANTILLA COMUN') or formal title.
    Returns True if something was updated.
    """
    # Build target text using your existing logic
    try:
        target = determinar_titulo(nombre, tipo)
    except Exception:
        target = str(nombre)

    # 1) If you want to prioritize fixed header text (like PLANTILLA…), try by-text first
    if prefer_text_header:
        shp = _find_header_by_text(slide)
        if shp and _set_text(shp, target):
            _bold_before_pipe(target, shp)
            print(f"Header (by text) updated to: {target}")
            return True

    # 2) Formal title placeholder
    shp = _find_formal_title(slide)
    if shp and _set_text(shp, target):
        _bold_before_pipe(target, shp)
        print(f"Title placeholder updated to: {target}")
        return True

    # 3) Positional heuristic: top band, wide textbox
    shp = _find_header_by_position(slide)
    if shp and _set_text(shp, target):
        _bold_before_pipe(target, shp)
        print(f"Header (by position) updated to: {target}")
        return True

    print("No suitable header/title textbox found.")
    return False

def change_all_titles(pres, nombre, tipo):
    ok = 0
    for s in pres.Slides:
        ok += bool(change_slide_title(s, nombre, tipo))
    print(f"Titles updated on {ok}/{pres.Slides.Count} slides.")