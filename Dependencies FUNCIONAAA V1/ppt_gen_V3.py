# -*- coding: utf-8 -*-

# ──────────────────────────────────────────────────────────────────────────────
# Imports
# ──────────────────────────────────────────────────────────────────────────────

import os
import re
import zipfile
import shutil
from io import BytesIO
from copy import deepcopy

import win32com.client as win32
from lxml import etree

from Dependencies.utilities import refresh_all_graphs
from Dependencies.Palletes import (
    TYPE_PALETTES,
    TYPE_LABEL_COLORS,
    SPECIAL_LABEL_COLORS,
    pick_series_color_by_type,
)
from Dependencies.models import COLOR_MAP
from Dependencies.Titulos import determinar_titulo, change_slide_title


# ──────────────────────────────────────────────────────────────────────────────
# Optional dependency: Pillow (image verification during unzip)
# ──────────────────────────────────────────────────────────────────────────────

try:
    from PIL import Image
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pillow"])
    from PIL import Image


# ──────────────────────────────────────────────────────────────────────────────
# PPT unzipper and housekeeping
# ──────────────────────────────────────────────────────────────────────────────

def prepare_output_dir(pptx_path: str) -> str:
    """
    Create a clean extraction directory next to the pptx: <base>_unzipped.
    Deletes any previous directory with the same name.
    """
    base, _ = os.path.splitext(pptx_path)
    extract_dir = base + "_unzipped"
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir, exist_ok=True)
    return extract_dir


def extract_member_safe(zip_ref: zipfile.ZipFile, info: zipfile.ZipInfo,
                        extract_dir: str, deleted_files: list[str]) -> bool:
    """
    Extract a single ZIP member safely. On failure, remove partial file,
    record it in deleted_files, and continue.
    """
    target_path = os.path.join(extract_dir, info.filename)
    os.makedirs(os.path.dirname(target_path), exist_ok=True)
    try:
        zip_ref.extract(info, extract_dir)
        return True
    except Exception as e:
        if os.path.exists(target_path):
            os.remove(target_path)
        print(f"Skipped unreadable entry '{info.filename}': {e}")
        deleted_files.append(info.filename)
        return False


def verify_and_clean_images(extract_dir: str, deleted_files: list[str]) -> None:
    """
    Verify JPEG/PNG media; delete unreadable images to avoid PPT repair prompts.
    """
    for root, _, files in os.walk(extract_dir):
        for file in files:
            if not file.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            path = os.path.join(root, file)
            try:
                with open(path, "rb") as f:
                    Image.open(BytesIO(f.read())).verify()
            except Exception:
                rel = os.path.relpath(path, extract_dir)
                print(f"Deleted broken image: {rel}")
                os.remove(path)
                deleted_files.append(rel)


def write_log(extract_dir: str, deleted_files: list[str]) -> None:
    """
    If any entries were skipped/deleted during unzip/verification, write a log.
    """
    if not deleted_files:
        return
    log_path = os.path.join(extract_dir, "deleted_media.txt")
    with open(log_path, "w", encoding="utf-8") as log:
        log.write("Deleted or unreadable files:\n")
        log.write("\n".join(deleted_files))
    print(f"Log written to: {log_path}")


def unzip_pptx_clean(ppt_export_location: str) -> str:
    """
    Unzip a .pptx into <base>_unzipped, skipping unreadable entries and
    deleting broken images. Returns the extraction directory path.
    """
    extract_dir = prepare_output_dir(ppt_export_location)
    deleted_files: list[str] = []

    with zipfile.ZipFile(ppt_export_location, "r") as zip_ref:
        for info in zip_ref.infolist():
            extract_member_safe(zip_ref, info, extract_dir, deleted_files)

    verify_and_clean_images(extract_dir, deleted_files)
    write_log(extract_dir, deleted_files)

    print(f"Extracted clean version to:\n{extract_dir}")
    return extract_dir


# ──────────────────────────────────────────────────────────────────────────────
# PowerPoint (COM) — duplicate template slides
# ──────────────────────────────────────────────────────────────────────────────

def CopyTemplateSlides(ppt_template_path: str,
                       preguntas_parseadas,
                       excel_save_location: str | None,
                       visible: bool = False) -> None:
    """
    Duplica slides de plantilla, aplica títulos y rellena en cada 'q' los campos:
      - slide_index (COM)
      - slide_id     (COM + verificación vía presentation.xml)
      - chart_rel_id (rels del slide)
      - chart_xml_path (ruta 'ppt/charts/chartN.xml')
    """
    ppt_app = win32.Dispatch("PowerPoint.Application")
    if visible:
        ppt_app.Visible = True
    ppt_app.DisplayAlerts = False

    pres = ppt_app.Presentations.Open(ppt_template_path, WithWindow=False)

    base_slide_1 = pres.Slides(1)
    base_slide_2 = pres.Slides(2)

    print(f"Total preguntas: {len(preguntas_parseadas)}")

    created_slides = []  # guardo objetos Slide para extraer SlideIndex/SlideID
    for i, q in enumerate(preguntas_parseadas, start=1):
        source_slide = base_slide_2 if getattr(q, "tipo", None) == "multiple" else base_slide_1
        new_slide = source_slide.Duplicate()
        new_slide.MoveTo(pres.Slides.Count)
        created_slides.append(new_slide)  # SlideRange
        print(f"{i}. Created slide for '{q.nombre}' (tipo: {getattr(q, 'tipo', None)})")

        # Título
        try:
            titulo = determinar_titulo(getattr(q, "nombre", ""), getattr(q, "tipo", ""))
            ok = change_slide_title(new_slide, getattr(q, "nombre", ""), getattr(q, "tipo", ""))
            if not ok:
                for shape in new_slide.Shapes:
                    try:
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            text = (shape.TextFrame.TextRange.Text or "").strip().upper()
                            if "TITULO" in text or "TÍTULO" in text:
                                shape.TextFrame.TextRange.Text = titulo
                                ok = True
                                break
                    except Exception:
                        continue
            print(f"Slide {i}: title {'updated' if ok else 'NOT updated'} -> {titulo}")
        except Exception as e:
            print(f"Warning setting title on slide {i}: {e}")

    # Elimino las dos plantillas base
    try:
        pres.Slides(2).Delete()
        pres.Slides(1).Delete()
        print("Deleted the two template slides.")
    except Exception as e:
        print(f"Error deleting templates: {e}")

    # Guardar
    save_path = excel_save_location or ppt_template_path
    pres.SaveAs(save_path)
    print(f"Saved presentation as: {save_path}")

    # ---- Completar manifest en cada pregunta ----
    # 1) slide_index / slide_id desde COM (tras el borrado, índices finales)
    #    SlideRange puede contener varias; iteramos asegurando orden de creación
    for q, srange in zip(preguntas_parseadas, created_slides):
        # 'srange' es SlideRange; tomar el primer elemento
        s = srange(1)
        try:
            q.slide_index = int(s.SlideIndex)
        except Exception:
            q.slide_index = None
        try:
            q.slide_id = str(s.SlideID)  # entero largo → str
        except Exception:
            q.slide_id = None

    # 2) chart_rel_id / chart_xml_path leyendo el paquete .pptx
    try:
        import zipfile
        from lxml import etree

        def _read_xml(zf, path):
            return etree.fromstring(zf.read(path))

        with zipfile.ZipFile(save_path, "r") as z:
            names = set(z.namelist())

            # Mapa rId → Target de presentation.xml.rels
            pres_rels = _read_xml(z, "ppt/_rels/presentation.xml.rels")
            rel_map = {}
            for r in pres_rels.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                rel_map[r.get("Id")] = r.get("Target")

            # Orden de slides + SlideID (atributo p:sldId @id)
            pres_xml = _read_xml(z, "ppt/presentation.xml")
            ns = {
                "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
                "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            }
            slide_parts = []  # [(slide_path, slide_id_int)]
            for sld in pres_xml.findall(".//p:sldIdLst/p:sldId", ns):
                rid = sld.get("{%s}id" % ns["r"]) or sld.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                sp = rel_map.get(rid, "")
                if sp:
                    slide_path = "ppt/" + sp.lstrip("/")
                    slide_id_attr = sld.get("id")  # SlideID "duro" del DOC (no el COM)
                    slide_parts.append((slide_path, slide_id_attr))

            # Para cada slide, leer su .rels y buscar el primer chart
            chart_links = []  # [(chart_rel_id, chart_xml_path)]
            for slide_path, _sid in slide_parts:
                rels_path = slide_path.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
                chart_rel_id, chart_xml_path = None, None
                if rels_path in names:
                    srels = _read_xml(z, rels_path)
                    for r in srels.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                        if r.get("Type", "").endswith("/relationships/chart"):
                            chart_rel_id = r.get("Id")
                            tgt = r.get("Target", "").replace("../", "").lstrip("/")
                            chart_xml_path = "ppt/" + tgt if not tgt.startswith("ppt/") else tgt
                            break
                chart_links.append((chart_rel_id, chart_xml_path))

            # Ahora, las primeras N slides del archivo corresponden exactamente
            # a las N slides creadas (ya borramos las 2 plantillas).
            # Emparejo 1 a 1 por orden.
            for q, (rel_id, cpath) in zip(preguntas_parseadas, chart_links):
                q.chart_rel_id = rel_id
                q.chart_xml_path = cpath

            # Extra: si quieres alinear por SlideID exacto:
            # creo índice: slide_id_doc -> (rel_id, cpath)
            sid_to_chart = {}
            for (slide_path, sid_doc), (rel_id, cpath) in zip(slide_parts, chart_links):
                sid_to_chart[str(sid_doc)] = (rel_id, cpath)

            # Si algún q ya tenía slide_id (vía COM), lo revalido/reescribo
            for q in preguntas_parseadas:
                if getattr(q, "slide_id", None) in sid_to_chart:
                    rel_id, cpath = sid_to_chart[q.slide_id]
                    q.chart_rel_id = rel_id
                    q.chart_xml_path = cpath

        print("Manifest PPT → preguntas: mapeo chart_rel_id / chart_xml_path completado.")
    except Exception as e:
        print(f"Warning building chart manifest: {e}")

    pres.Close()
    ppt_app.Quit()
    print("PowerPoint closed successfully.")


# ──────────────────────────────────────────────────────────────────────────────
# XML Namespaces & utilities
# ──────────────────────────────────────────────────────────────────────────────

REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

def load_xml(path: str):
    """
    Parse an XML file (preserving whitespace). Returns (tree, root).
    """
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(path, parser)
    return tree, tree.getroot()


def save_xml(tree: etree._ElementTree, path: str) -> None:
    """
    Write XML with declaration, UTF-8, and standalone='yes'.
    """
    tree.write(path, xml_declaration=True, encoding="UTF-8", standalone="yes")


def split_rows(ubicacion: str, respuestas: list[str]) -> list[str]:
    """
    Convert a multi-row Excel range into per-row ranges (one per respuesta).
    Example: 'Sheet1!$B$10:$E$14' → per-row 'Sheet1!$B$10:$E$10' ... '$E$14'.
    """
    sheet, coords = ubicacion.split("!")
    sheet = sheet.strip("'")
    start, end = coords.split(":")
    start_col = re.findall(r"\$?([A-Z]+)", start)[0]
    start_row = int(re.findall(r"\$?(\d+)", start)[0])
    end_col = re.findall(r"\$?([A-Z]+)", end)[0]

    ranges: list[str] = []
    for i, _ in enumerate(respuestas):
        row = start_row + i + 1
        ranges.append(f"{sheet}!${start_col}${row}:${end_col}${row}")
    return ranges


def list_chart_files(pptx_dir: str) -> list[str]:
    """
    Return sorted paths to chart XML files under ppt/charts.
    """
    chart_dir = os.path.join(pptx_dir, "ppt", "charts")
    charts = sorted(
        os.path.join(chart_dir, f)
        for f in os.listdir(chart_dir)
        if f.startswith("chart")
    )
    return charts


def match_charts_to_questions(chart_files: list[str], preguntas) -> list[tuple[str, object]]:
    """
    Pair chart files with questions by order.
    """
    return list(zip(chart_files, preguntas))


# ──────────────────────────────────────────────────────────────────────────────
# Series color and chart editing helpers
# ──────────────────────────────────────────────────────────────────────────────

def set_series_colorDEPRECATED(ser, hex_rgb: str) -> None:
    """
    Set solid RGB color for series fill and line stroke. hex_rgb='RRGGBB'.
    """
    # Fill
    spPr = ser.find("c:spPr", NS)
    if spPr is None:
        spPr = etree.SubElement(ser, f"{{{NS['c']}}}spPr")

    solidFill = spPr.find("a:solidFill", NS)
    if solidFill is None:
        solidFill = etree.SubElement(spPr, f"{{{NS['a']}}}solidFill")

    scheme = solidFill.find("a:schemeClr", NS)
    if scheme is not None:
        solidFill.remove(scheme)

    srgb = solidFill.find("a:srgbClr", NS)
    if srgb is None:
        srgb = etree.SubElement(solidFill, f"{{{NS['a']}}}srgbClr")
    srgb.set("val", hex_rgb)

    # Stroke
    ln = spPr.find("a:ln", NS)
    if ln is None:
        ln = etree.SubElement(spPr, f"{{{NS['a']}}}ln")

    lfill = ln.find("a:solidFill", NS)
    if lfill is None:
        lfill = etree.SubElement(ln, f"{{{NS['a']}}}solidFill")

    lscheme = lfill.find("a:schemeClr", NS)
    if lscheme is not None:
        lfill.remove(lscheme)

    lrgb = lfill.find("a:srgbClr", NS)
    if lrgb is None:
        lrgb = etree.SubElement(lfill, f"{{{NS['a']}}}srgbClr")
    lrgb.set("val", hex_rgb)

#===========================================DEBUG

# Optional toggle at module top

# Optional toggle at module top
DEBUG_COLOR = True

def set_series_colorDEPRECATED2(ser, hex_rgb: str) -> None:
    """
    Set solid RGB color for series fill and line stroke. hex_rgb='RRGGBB'.
    Debug prints show structure before/after and flag common OOXML violations.
    """
    from lxml import etree

    def _lname(e):
        return etree.QName(e.tag).localname if e is not None else None

    def _names(children):
        return [etree.QName(x.tag).localname for x in children]

    # --- Identify series & parent chart for context ---
    idx_node   = ser.find("c:idx", NS)
    order_node = ser.find("c:order", NS)
    tx_node    = (ser.find("c:tx/c:strRef/c:strCache/c:pt/c:v", NS)
                  or ser.find("c:tx/c:v", NS))
    ser_idx    = idx_node.get("val") if idx_node is not None else "?"
    ser_order  = order_node.get("val") if order_node is not None else "?"
    ser_title  = (tx_node.text or "").strip() if tx_node is not None else ""

    chart = ser.getparent()  # e.g., c:barChart / c:colChart / c:lineChart
    chart_name = _lname(chart)

    # --- Snapshot BEFORE ---
    before_ser_children   = _names(list(ser))
    before_chart_children = _names(list(chart)) if chart is not None else []

    if DEBUG_COLOR:
        print(f"[color-debug] BEGIN ser idx={ser_idx} order={ser_order} title={ser_title!r} "
              f"hex={hex_rgb} chart={chart_name}")
        print(f"[color-debug]   ser children BEFORE: {before_ser_children}")
        if chart is not None:
            print(f"[color-debug]   chart children BEFORE: {before_chart_children}")

    # --- Existence flags BEFORE (so we know what we’re creating) ---
    spPr_before     = ser.find("c:spPr", NS) is not None
    solidFill_before= ser.find("c:spPr/a:solidFill", NS) is not None
    scheme_before   = ser.find("c:spPr/a:solidFill/a:schemeClr", NS) is not None
    srgb_before     = ser.find("c:spPr/a:solidFill/a:srgbClr", NS) is not None
    ln_before       = ser.find("c:spPr/a:ln", NS) is not None
    lfill_before    = ser.find("c:spPr/a:ln/a:solidFill", NS) is not None
    lscheme_before  = ser.find("c:spPr/a:ln/a:solidFill/a:schemeClr", NS) is not None
    lrgb_before     = ser.find("c:spPr/a:ln/a:solidFill/a:srgbClr", NS) is not None

    # --- Basic input sanity ---
    if DEBUG_COLOR and not (isinstance(hex_rgb, str) and len(hex_rgb) == 6 and all(c in "0123456789ABCDEFabcdef" for c in hex_rgb)):
        print(f"[color-debug][WARN] hex_rgb looks odd: {hex_rgb!r} (expected 6 hex chars)")

    # =========================
    # Original behavior (unchanged)
    # =========================

    # Fill
    spPr = ser.find("c:spPr", NS)
    if spPr is None:
        spPr = etree.SubElement(ser, f"{{{NS['c']}}}spPr")

    solidFill = spPr.find("a:solidFill", NS)
    if solidFill is None:
        solidFill = etree.SubElement(spPr, f"{{{NS['a']}}}solidFill")

    scheme = solidFill.find("a:schemeClr", NS)
    if scheme is not None:
        solidFill.remove(scheme)

    srgb = solidFill.find("a:srgbClr", NS)
    if srgb is None:
        srgb = etree.SubElement(solidFill, f"{{{NS['a']}}}srgbClr")
    srgb.set("val", hex_rgb.upper())

    # Stroke
    ln = spPr.find("a:ln", NS)
    if ln is None:
        ln = etree.SubElement(spPr, f"{{{NS['a']}}}ln")

    lfill = ln.find("a:solidFill", NS)
    if lfill is None:
        lfill = etree.SubElement(ln, f"{{{NS['a']}}}solidFill")

    lscheme = lfill.find("a:schemeClr", NS)
    if lscheme is not None:
        lfill.remove(lscheme)

    lrgb = lfill.find("a:srgbClr", NS)
    if lrgb is None:
        lrgb = etree.SubElement(lfill, f"{{{NS['a']}}}srgbClr")
    lrgb.set("val", hex_rgb.upper())

    # --- Snapshot AFTER ---
    after_ser_children   = _names(list(ser))
    after_chart_children = _names(list(chart)) if chart is not None else []

    # --- Existence flags AFTER ---
    spPr_after     = ser.find("c:spPr", NS) is not None
    solidFill_after= ser.find("c:spPr/a:solidFill", NS) is not None
    srgb_after     = ser.find("c:spPr/a:solidFill/a:srgbClr", NS) is not None
    ln_after       = ser.find("c:spPr/a:ln", NS) is not None
    lfill_after    = ser.find("c:spPr/a:ln/a:solidFill", NS) is not None
    lrgb_after     = ser.find("c:spPr/a:ln/a:solidFill/a:srgbClr", NS) is not None

    # --- Invariant checks & helpful warnings ---
    warn = []

    # A) Ensure no conflicting fill children under spPr/solidFill
    sf_children = ser.findall("c:spPr/a:solidFill/*", NS)
    if any(_lname(x) != "srgbClr" for x in sf_children):
        warn.append("spPr/solidFill has non-srgb children")

    # B) spPr should precede cat/val/xVal/yVal/bubbleSize
    names_after = after_ser_children
    if "spPr" in names_after:
        sp_idx = names_after.index("spPr")
        for later in ("cat", "val", "xVal", "yVal", "bubbleSize"):
            if later in names_after and sp_idx > names_after.index(later):
                warn.append(f"spPr appears after {later}")

    # C) All <c:ser> should be before <c:axId> at chart level
    if chart is not None:
        ser_pos = [i for i, x in enumerate(chart) if _lname(x) == "ser"]
        ax_pos  = [i for i, x in enumerate(chart) if _lname(x) == "axId"]
        if ser_pos and ax_pos and max(ser_pos) >= min(ax_pos):
            warn.append("a <c:ser> appears after <c:axId> (chart child order)")

    if DEBUG_COLOR:
        print(f"[color-debug]   CREATED: spPr:{not spPr_before and spPr_after} "
              f"solidFill:{not solidFill_before and solidFill_after} "
              f"srgb:{not srgb_before and srgb_after} "
              f"ln:{not ln_before and ln_after} "
              f"lfill:{not lfill_before and lfill_after} "
              f"lrgb:{not lrgb_before and lrgb_after}")
        print(f"[color-debug]   ser children AFTER:  {after_ser_children}")
        if chart is not None:
            print(f"[color-debug]   chart children AFTER: {after_chart_children}")
        if warn:
            print(f"[color-debug][WARN] {' | '.join(warn)}")
        print(f"[color-debug] END ser idx={ser_idx} order={ser_order}\n")

def set_series_color(ser, hex_rgb: str, stroke: str = "solid") -> None:
    """
    Set series FILL to solid srgb (hex_rgb).
    Stroke: "solid" → solid line same color; "none" → no line.
    """
    from lxml import etree
    hex_rgb = (hex_rgb or "").upper()

    # --- ensure <c:spPr>
    spPr = ser.find("c:spPr", NS)
    if spPr is None:
        spPr = etree.SubElement(ser, f"{{{NS['c']}}}spPr")

    # --- FILL (series body) → solid srgb
    for q in ("a:noFill", "a:gradFill", "a:pattFill"):
        n = spPr.find(q, NS)
        if n is not None:
            spPr.remove(n)
    solidFill = spPr.find("a:solidFill", NS)
    if solidFill is None:
        solidFill = etree.SubElement(spPr, f"{{{NS['a']}}}solidFill")
    # keep only srgbClr inside solidFill
    for child in list(solidFill):
        if etree.QName(child).localname != "srgbClr":
            solidFill.remove(child)
    srgb = solidFill.find("a:srgbClr", NS)
    if srgb is None:
        srgb = etree.SubElement(solidFill, f"{{{NS['a']}}}srgbClr")
    srgb.set("val", hex_rgb)

    # --- LINE (border)
    ln = spPr.find("a:ln", NS)
    if ln is None:
        ln = etree.SubElement(spPr, f"{{{NS['a']}}}ln")

    # remove any existing fill mode under <a:ln>
    for q in ("a:noFill", "a:solidFill", "a:gradFill", "a:pattFill"):
        n = ln.find(q, NS)
        if n is not None:
            ln.remove(n)

    if stroke == "none":
        # no border at all
        etree.SubElement(ln, f"{{{NS['a']}}}noFill")
    else:
        # solid border same color as fill
        lfill = etree.SubElement(ln, f"{{{NS['a']}}}solidFill")
        lrgb  = etree.SubElement(lfill, f"{{{NS['a']}}}srgbClr")
        lrgb.set("val", hex_rgb)



def ensure_extlst(ser) -> None:
    """
    Ensure each <c:ser> ends with a minimal <c:extLst>.
    """
    extlst = ser.find("c:extLst", NS)
    if extlst is None:
        extlst = etree.SubElement(ser, f"{{{NS['c']}}}extLst")
        etree.SubElement(
            extlst,
            f"{{{NS['c']}}}ext",
            uri="{A1602332-AB5F-4743-BC9F-DDF7F2FEE8C2}",
        )


def set_black_fill(ser) -> None:
    """
    Force series fill to solid black (#000000).
    """
    spPr = ser.find("c:spPr", NS)
    if spPr is None:
        spPr = etree.SubElement(ser, f"{{{NS['c']}}}spPr")
    solidFill = spPr.find("a:solidFill", NS)
    if solidFill is None:
        solidFill = etree.SubElement(spPr, f"{{{NS['a']}}}solidFill")
    srgb = solidFill.find("a:srgbClr", NS)
    if srgb is None:
        srgb = etree.SubElement(solidFill, f"{{{NS['a']}}}srgbClr")
    srgb.set("val", "000000")


def set_series_title_literal(ser, title: str) -> None:
    """
    Replace the series title with a literal <c:v>.
    """
    tx = ser.find("c:tx", NS)
    if tx is None:
        tx = etree.SubElement(ser, f"{{{NS['c']}}}tx")
    for child in list(tx):
        tx.remove(child)
    v = etree.SubElement(tx, f"{{{NS['c']}}}v")
    v.text = title


def set_series_values_range(ser, excel_range: str) -> None:
    """
    Set <c:val>/<c:numRef>/<c:f> to the given A1 range. Remove any <c:numCache>.
    """
    numref = ser.find("c:val/c:numRef", NS)
    if numref is None:
        val = ser.find("c:val", NS) or etree.SubElement(ser, f"{{{NS['c']}}}val")
        numref = etree.SubElement(val, f"{{{NS['c']}}}numRef")
    f = numref.find("c:f", NS)
    if f is None:
        f = etree.SubElement(numref, f"{{{NS['c']}}}f")
    f.text = excel_range
    num_cache = numref.find("c:numCache", NS)
    if num_cache is not None:
        numref.remove(num_cache)


def _all_series_must_live_in_barchart(root) -> None:
    """
    If there are stray <c:ser> directly under <c:plotArea>, move them into
    <c:barChart> (if present). Does not create/delete series.
    """
    bar = root.find(".//c:barChart", NS)
    if bar is None:
        return
    plot_area = root.find(".//c:plotArea", NS)
    for ser in list(plot_area.findall("./c:ser", NS)):
        plot_area.remove(ser)
        bar.append(ser)


# ──────────────────────────────────────────────────────────────────────────────
# Chart update core
# ──────────────────────────────────────────────────────────────────────────────

def update_chart_for_question(chart_path: str, question) -> None:
    """
    Update one chart XML according to a question object.
    """
    print(f"\nProcessing chart file: {os.path.basename(chart_path)}")
    print(f"  Question: {question.nombre}")
    print(f"  Ubicacion: {question.ubicacion}")
    print(f"  Respuestas: {question.respuestas}")
    print(f"  Tipo: {question.tipo}")

    tree, root = load_xml(chart_path)
    assign_series_from_question(root, question)
    save_xml(tree, chart_path)
    print(f"  Saved updates to {os.path.basename(chart_path)}")


def ensure_external_data(root) -> None:
    """
    Ensure <c:externalData r:id='rId1'> exists under <c:chart>.
    Does not edit .rels files, only the chart XML node.
    """
    ext_data = root.find(".//c:externalData", NS)
    if ext_data is None:
        chart = root.find(".//c:chart", NS)
        if chart is not None:
            ext_data = etree.SubElement(chart, f"{{{NS['c']}}}externalData")
            ext_data.set(f"{{{NS['r']}}}id", "rId1")
            auto = etree.SubElement(ext_data, f"{{{NS['c']}}}autoUpdate")
            auto.set("val", "0")
            print("  Added externalData link (rId1).")
    else:
        rid = ext_data.get(f"{{{NS['r']}}}id")
        if rid != "rId1":
            ext_data.set(f"{{{NS['r']}}}id", "rId1")
            print("  Fixed externalData r:id -> rId1")


def assign_series_from_questionDEPRECATED(root, question) -> None:
    """
    Align the chart's series with the question:
      - Ensure series count matches respuestas (clone/remove as needed).
      - Set title, values range, order.
      - Pick and apply color.
      - Ensure extLst.
      - Relocate stray series into <c:barChart>.
    """
    series = root.findall(".//c:ser", NS)
    row_ranges = split_rows(question.ubicacion, question.respuestas)
    needed = len(question.respuestas)
    current = len(series)
    plot_area = root.find(".//c:plotArea", NS)

    print(f"  Found {current} series, need {needed}")

    # Add series by cloning the last one
    if needed > current and series:
        template_ser = series[-1]
        for i in range(needed - current):
            new_ser = deepcopy(template_ser)
            for tag in new_ser.findall(".//c:idx", NS) + new_ser.findall(".//c:order", NS):
                tag.set("val", str(current + i))
            set_black_fill(new_ser)
            plot_area.append(new_ser)
            series.append(new_ser)
            print(f"    Added new cloned series {current + i + 1}")

    # Remove extra series
    elif needed < current:
        for extra in series[needed:]:
            parent = extra.getparent()
            parent.remove(extra)
            print("    Removed extra series")
        series = series[:needed]

    # Update each series
    for i, (ser, resp, rng) in enumerate(zip(series, question.respuestas, row_ranges)):
        set_series_title_literal(ser, resp)
        set_series_values_range(ser, rng)

        order = ser.find("c:order", NS)
        if order is not None:
            order.set("val", str(i))

        # Color selection
        qtype = (getattr(question, "tipo", "") or "").lower()
        color_hex = pick_series_color_by_type(
            COLOR_MAP,
            TYPE_PALETTES,
            TYPE_LABEL_COLORS,
            SPECIAL_LABEL_COLORS,
            qtype,
            resp,
            i,
        )
        set_series_color(ser, color_hex)
        ensure_extlst(ser)
        print(f"    Series {i+1}: '{resp}'  {rng}  color={color_hex}")

    _all_series_must_live_in_barchart(root)


# ──────────────────────────────────────────────────────────────────────────────
# Repack PPTX
# ──────────────────────────────────────────────────────────────────────────────

def rezip_pptx(source_dir: str, output_path: str) -> None:
    """
    Repack a directory back into a .pptx, preserving relative paths.
    """
    if not os.path.exists(os.path.join(source_dir, "[Content_Types].xml")):
        raise ValueError(f"{source_dir} is not the PPTX root (missing [Content_Types].xml).")

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(source_dir):
            for f in files:
                full_path = os.path.join(root, f)
                rel_path = os.path.relpath(full_path, source_dir)
                z.write(full_path, rel_path)
    print(f"Repacked PPTX: {output_path}")


# ──────────────────────────────────────────────────────────────────────────────
# Orchestration: unzip → edit charts → rezip
# ──────────────────────────────────────────────────────────────────────────────



def update_charts_in_pptx(ppt_export_location: str, preguntas_parseadas) -> str:
    """
    Pipeline (con manifest por pregunta):
      1) Unzip PPTX en carpeta temporal.
      2) Para cada pregunta, usar q.chart_xml_path → actualizar ese chart.
      3) Reempaquetar como '*_updated.pptx'.
    """
    tmp_dir = unzip_pptx_clean(ppt_export_location)

    updated, skipped = 0, 0

    def _abs_chart_path(tmp_root: str, rel_path: str) -> str | None:
        """
        Normaliza y resuelve 'ppt/charts/chartN.xml' (o variantes) al path absoluto dentro del unzip.
        Devuelve None si no existe.
        """
        if not rel_path:
            return None

        rel_norm = rel_path.replace("\\", "/").lstrip("/")

        # Si no empieza con 'ppt/', forzamos
        if not rel_norm.lower().startswith("ppt/"):
            rel_norm = f"ppt/{rel_norm}"

        # Limpieza de posibles '../'
        rel_norm = os.path.normpath(rel_norm).replace("\\", "/")

        # Construyo absoluto
        abs_path = os.path.normpath(os.path.join(tmp_root, *rel_norm.split("/")))
        if os.path.exists(abs_path):
            return abs_path

        # Intento alternativo por si vino con 'ppt/../charts/...'
        rel_alt = rel_norm.replace("ppt/../", "ppt/")
        abs_alt = os.path.normpath(os.path.join(tmp_root, *rel_alt.split("/")))
        if os.path.exists(abs_alt):
            return abs_alt

        return None

    for q in preguntas_parseadas:
        rel = getattr(q, "chart_xml_path", None)
        chart_path = _abs_chart_path(tmp_dir, rel)

        if not chart_path:
            print(f"SKIP: '{getattr(q, 'nombre', '')}' sin chart resoluble "
                  f"(chart_xml_path={rel}, slide_index={getattr(q,'slide_index',None)}, slide_id={getattr(q,'slide_id',None)})")
            skipped += 1
            continue

        print(f"Updating {os.path.basename(chart_path)} "
              f"for slide {getattr(q,'slide_index',None)} (id {getattr(q,'slide_id',None)}): {getattr(q,'nombre','')}")
        update_chart_for_question(chart_path, q)
        updated += 1

    output_path = ppt_export_location.replace(".pptx", "_updated.pptx")
    rezip_pptx(tmp_dir, output_path)

    #print(f"Charts updated: {updated}, skipped: {skipped}")
    #print(f"New file created: {output_path}")
    return output_path


##DEBUG

def _norm_label_for_debug(s: str) -> str:
    # simple, local normalization (matches Palletes._normalize_label intent)
    import unicodedata
    s = (s or "").strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s

def _color_debug(qtype: str, legend_text: str, idx: int) -> tuple[str, str]:
    """
    Returns (color_hex, debug_str) without changing how the color is actually chosen.
    We call your picker to keep behavior identical, but also compute a readable reason.
    """
    # Always get the actual color using the canonical function:
    color_hex = pick_series_color_by_type(
        COLOR_MAP, TYPE_PALETTES, TYPE_LABEL_COLORS, SPECIAL_LABEL_COLORS,
        (qtype or "").lower(), legend_text, idx
    )

    qt = (qtype or "").lower()
    norm = _norm_label_for_debug(legend_text)

    # 1) Special labels override everything
    if norm in SPECIAL_LABEL_COLORS:
        return color_hex, f"special label → '{legend_text}' = {SPECIAL_LABEL_COLORS[norm]}"

    # 2) Type-specific explicit mapping
    tmap = TYPE_LABEL_COLORS.get(qt, {})
    if norm in tmap:
        return color_hex, f"type label map ({qt}) → '{legend_text}' = {tmap[norm]}"

    # 3) Palette path via COLOR_MAP → TYPE_PALETTES
    pal_id = COLOR_MAP.get(qt)
    if pal_id is not None:
        pal = TYPE_PALETTES.get(pal_id, [])
        if idx < len(pal):
            return color_hex, f"palette {pal_id} index {idx} → {pal[idx]} | palette={pal}"
        else:
            return color_hex, f"palette {pal_id} exhausted → deterministic fallback (idx={idx})"

    # 4) Unknown type → deterministic
    return color_hex, f"unknown type '{qt}' → deterministic fallback"

def assign_series_from_question(root, question) -> None:
    """
    Align the chart's series with the question:
      - Ensure series count matches respuestas (clone/remove as needed).
      - Set title, values range, order.
      - Pick and apply color (with debug on palette source).
      - Ensure extLst.
      - Relocate stray series into <c:barChart>.
    """
    series = root.findall(".//c:ser", NS)
    row_ranges = split_rows(question.ubicacion, question.respuestas)
    needed = len(question.respuestas)
    current = len(series)
    plot_area = root.find(".//c:plotArea", NS)

    print(f"  Found {current} series, need {needed}")

    # Add series by cloning the last one
    if needed > current and series:
        template_ser = series[-1]
        for i in range(needed - current):
            new_ser = deepcopy(template_ser)
            for tag in new_ser.findall(".//c:idx", NS) + new_ser.findall(".//c:order", NS):
                tag.set("val", str(current + i))
            set_black_fill(new_ser)
            plot_area.append(new_ser)
            series.append(new_ser)
            print(f"    Added new cloned series {current + i + 1}")

    # Remove extra series
    elif needed < current:
        for extra in series[needed:]:
            parent = extra.getparent()
            parent.remove(extra)
            print("    Removed extra series")
        series = series[:needed]

    # Update each series
    for i, (ser, resp, rng) in enumerate(zip(series, question.respuestas, row_ranges)):
        set_series_title_literal(ser, resp)
        set_series_values_range(ser, rng)

        order = ser.find("c:order", NS)
        if order is not None:
            order.set("val", str(i))

        qtype = (getattr(question, "tipo", "") or "").lower()

        # Get the exact color AND a human-readable reason for how it was chosen
        color_hex, why = _color_debug(qtype, resp, i)

        set_series_color(ser, color_hex, stroke="none")
        ##ensure_extlst(ser)
        ##print(f"    Series {i+1}: '{resp}'  {rng}  color={color_hex}  [{why}]")

    _all_series_must_live_in_barchart(root)
