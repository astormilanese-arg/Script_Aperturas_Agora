# ──────────────────────────────────────────────────────────────────────────────
# DIAGNOSTIC: find chart XML problems that make PowerPoint .Open() fail
# ──────────────────────────────────────────────────────────────────────────────
import zipfile
from io import BytesIO
from lxml import etree
import re
from collections import defaultdict

A1_RE = re.compile(r"^'?(?P<sheet>[^'!]+)'?!(?P<ref>\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)$")

def _pp(s):
    try:
        return s.encode("utf-8", "ignore").decode("utf-8", "ignore")
    except Exception:
        return str(s)

def _is_valid_a1(f: str) -> tuple[bool, str]:
    if not f or "!" not in f:
        return False, "missing '!' (no sheet qualifier)"
    m = A1_RE.match(f)
    if not m:
        return False, "not a valid A1 range (quote the sheet + full A1:A1)"
    sheet = m.group("sheet")
    if " " in sheet and not (f.startswith("'") and "'!" in f):
        return False, "sheet has spaces but is not quoted"
    return True, ""

def diagnose_pptx_after_edit(pptx_path: str, preguntas, verbose: bool = True) -> dict:
    """
    Pure read-only diagnostics:
      - Zip integrity
      - Slide->chart rels exist
      - Each chart XML parses
      - Flags: dangling <c:externalData>, <c:ser> under plotArea, missing idx/order
      - numRef/c:f formulas validity (quoted sheet / A1 pattern)
      - Series count vs respuestas (if chart is mapped via manifest)
    Returns a summary dict and prints concise lines when verbose=True.
    """
    summary = {
        "zip_ok": None,
        "slides": 0,
        "charts_found": 0,
        "charts_with_errors": 0,
        "errors": defaultdict(list),  # key=chart path, val=list[str]
    }

    # Build manifest map chart_xml_path -> question (first occurrence)
    q_by_chart: dict[str, object] = {}
    for q in (preguntas or []):
        p = getattr(q, "chart_xml_path", None)
        if p:
            norm = p.replace("\\", "/").lstrip("/")
            if not norm.startswith("ppt/"):
                norm = "ppt/" + norm
            q_by_chart[norm] = q

    def log(msg): 
        if verbose: print(msg)

    # 0) Zip integrity
    try:
        with zipfile.ZipFile(pptx_path, "r") as z:
            bad = z.testzip()
            summary["zip_ok"] = (bad is None)
            if bad is not None:
                log(f"[diag][zip] CORRUPT entry → {bad}")
            else:
                log(f"[diag][zip] OK")
    except Exception as e:
        log(f"[diag][zip] cannot open/test: {e}")
        summary["zip_ok"] = False
        return summary  # bail early if the package isn't even readable

    with zipfile.ZipFile(pptx_path, "r") as z:
        names = set(z.namelist())

        # 1) Slide parts + rels quick check
        if "ppt/presentation.xml" in names and "ppt/_rels/presentation.xml.rels" in names:
            pres = etree.fromstring(z.read("ppt/presentation.xml"))
            rels = etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
            ns = {
                "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
                "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            }
            rel_map = {}
            for r in rels.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                rel_map[r.get("Id")] = r.get("Target")
            slide_targets = []
            for sld in pres.findall(".//p:sldIdLst/p:sldId", ns):
                rid = sld.get("{%s}id" % ns["r"]) or sld.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                tgt = rel_map.get(rid, "")
                if tgt:
                    slide_targets.append("ppt/" + tgt.lstrip("/"))
            summary["slides"] = len(slide_targets)
            log(f"[diag][rels] slides={len(slide_targets)}")

            # 2) For each slide, check chart relationship and that target exists
            for sp in slide_targets:
                rels_path = sp.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
                if rels_path not in names:
                    continue
                srels = etree.fromstring(z.read(rels_path))
                found = False
                for r in srels.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                    if r.get("Type", "").endswith("/relationships/chart"):
                        target = r.get("Target", "").replace("../", "").lstrip("/")
                        chart_path = "ppt/" + target if not target.startswith("ppt/") else target
                        if chart_path not in names:
                            summary["errors"][rels_path].append(f"rel to missing chart: {chart_path}")
                            summary["charts_with_errors"] += 1
                            log(f"[diag][rels] missing chart target → {chart_path}")
                        found = True
                if not found:
                    # not all slides must have a chart; that's fine
                    pass

        # 3) Validate each chart XML
        chart_paths = sorted(p for p in names if p.startswith("ppt/charts/chart") and p.endswith(".xml"))
        summary["charts_found"] = len(chart_paths)
        for cp in chart_paths:
            errs = []
            try:
                root = etree.fromstring(z.read(cp))
            except Exception as e:
                errs.append(f"XML parse error: {e}")
                summary["errors"][cp] += errs
                summary["charts_with_errors"] += 1
                log(f"[diag][chart] {cp} → PARSE ERROR: {e}")
                continue

            # 3a) Dangling externalData
            ext = root.find(".//{http://schemas.openxmlformats.org/drawingml/2006/chart}externalData")
            if ext is not None and not ext.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"):
                errs.append("dangling <c:externalData> without r:id")

            # 3b) Stray series directly under plotArea
            plot_area = root.find(".//{http://schemas.openxmlformats.org/drawingml/2006/chart}plotArea")
            if plot_area is not None:
                stray = plot_area.findall("./{http://schemas.openxmlformats.org/drawingml/2006/chart}ser")
                if stray:
                    errs.append(f"{len(stray)} <c:ser> directly under <c:plotArea>")

            # 3c) Series checks
            series = root.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/chart}ser")
            if not series:
                errs.append("no <c:ser> elements")

            idx_vals, order_vals = [], []
            for s in series:
                idx = s.find("{http://schemas.openxmlformats.org/drawingml/2006/chart}idx")
                order = s.find("{http://schemas.openxmlformats.org/drawingml/2006/chart}order")
                if idx is None or order is None:
                    errs.append("<c:ser> missing idx/order")
                else:
                    idx_vals.append(idx.get("val"))
                    order_vals.append(order.get("val"))
                f = s.find(".//{http://schemas.openxmlformats.org/drawingml/2006/chart}numRef/{http://schemas.openxmlformats.org/drawingml/2006/chart}f")
                if f is None or not (f.text and f.text.strip()):
                    errs.append("<c:ser> missing numRef/f")
                else:
                    ok, why = _is_valid_a1(f.text.strip())
                    if not ok:
                        errs.append(f"invalid A1 formula: {f.text.strip()} ({why})")

            # 3d) If manifest maps this chart, compare series count vs respuestas
            q = q_by_chart.get(cp)
            if q is not None:
                need = len(getattr(q, "respuestas", []) or [])
                cur = len(series)
                if need != cur:
                    errs.append(f"series count mismatch: have={cur} need={need} (question='{_pp(getattr(q,'nombre',''))[:60]}')")

            if errs:
                summary["errors"][cp] += errs
                summary["charts_with_errors"] += 1
                if verbose:
                    log(f"[diag][chart] {cp}")
                    for e in errs:
                        log(f"  - {e}")

    # Compact printout if anything broke
    if summary["charts_with_errors"] > 0 and verbose:
        log("\n[diag] SUMMARY of chart issues:")
        for cp, es in summary["errors"].items():
            for e in es:
                log(f"  {cp}: {e}")

    return summary