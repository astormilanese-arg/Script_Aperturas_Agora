
def update_charts_in_pptxDEPRECATED2(ppt_export_location: str, preguntas_parseadas) -> str:
    """
    Pipeline:
      1) Unzip PPTX into a working folder.
      2) List all chart XML files under ppt/charts.
      3) Pair one-to-one with preguntas_parseadas.
      4) Apply XML updates per pair.
      5) Repack into '*_updated.pptx'.
    """
    tmp_dir = unzip_pptx_clean(ppt_export_location)
    chart_files = list_chart_files(tmp_dir)
    pairs = match_charts_to_questions(chart_files, preguntas_parseadas)

    for chart_path, q in pairs:
        print(f"Updating {os.path.basename(chart_path)} for question: {q.nombre}")
        update_chart_for_question(chart_path, q)

    output_path = ppt_export_location.replace(".pptx", "_updated.pptx")
    rezip_pptx(tmp_dir, output_path)

    # Optionally refresh all graphs after rezip:
    # refresh_all_graphs(output_path)

    print(f"New file created: {output_path}")
    return output_path

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

def set_series_colorDEPRECATED3(ser, hex_rgb: str, stroke: str = "solid") -> None:
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

