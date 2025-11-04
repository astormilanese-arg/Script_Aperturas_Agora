import hashlib, random, unicodedata

# 1) Special legends (highest priority, independent of type)
SPECIAL_LABEL_COLORS = {
    "otro":    "A6A6A6",
    "en blanco": "D9D9D9",
    "no sabe": "BFBFBF",
    "ns/nc":   "BFBFBF",
    "no lo conoce o no tiene opinion": "A6A6A6",
    "no lo se aun": "A6A6A6",
    "si": "023D9B",
    "no": "E85833",
    "voy a votar seguro": "023D9B",
    "es bastante probable que vaya a votar": "1D73FC",
    "es poco probable que vaya a votar": "FF008E",
    "no voy a votar": "C00000",
    "libertario": "7030A0",
    "al pro": "FFC000",
    "al kirchnerismo": "0070C0",
    "al peronismo": "03715C",
    "a la izquierda": "FF008E",
    "al radicalismo": "C00000"
}

# 2) Type-specific explicit color per value (by legend text)
#    Example scaffolding — fill with your real mappings when you have them.
TYPE_LABEL_COLORS = {
    "multiple": {
        # "si": "4472C4",
        # "no": "ED7D31",
    },
    "simple":   {
        # "aprueba": "70AD47",
        # "desaprueba": "ED7D31",
    },
    # add other types...
}

# 3) Type fallback palettes (ordered lists). Used when label not explicitly mapped.
TYPE_PALETTES = {
    1: ["000000","023D9B","1D73FC","FF008E","7357BE","AF1956","E85833","007F48","02BFA3"],
    2: ["04967A","03BD85","EA3F28","C1273A"],
    3: ["023D9B","1D73FC","FF008E","C1273A"],
    4: ["04967A","03BD85","52D1A6","ED692F","EA3F28","C1273A"],
    5: ["023D9B","1D73FC","BFBFBF","FF008E","C1273A"],
    6: ["000000","023D9B","1D73FC","E85833","C1273A"],
    7: ["7030A0","9751CB","04967A","03BD85","E85833","F19B85"],
    8: ["9E480E","4472C4","ED7D31","70AD47","5B9BD5","FFC000","A5A5A5","264478","636363","997300"],
    9: ["4472C4","ED7D31","FFC000","70AD47","5B9BD5","A5A5A5","264478","9E480E","636363","997300"],
    10:["007F48","02BFA3","1D73FC","7357BE","AF1956","E85833","FFC000","A5A5A5","636363","997300"],
}

def _normalize_label(s):
    s = (s or "").strip().lower()
    # strip accents
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s

def _deterministic_hex(seed_text):
    """
    Return a stable hex color (RRGGBB) derived from seed_text.
    Avoids extremes for readability (keep channels in 40..215).
    """
    h = hashlib.sha1(seed_text.encode("utf-8")).digest()
    r, g, b = h[0], h[1], h[2]
    lo, hi = 40, 215
    r = lo + (r % (hi - lo + 1))
    g = lo + (g % (hi - lo + 1))
    b = lo + (b % (hi - lo + 1))
    return f"{r:02X}{g:02X}{b:02X}"

def pick_series_color_by_type(color_map, palettes, type_label_colors, special_label_colors,
                              qtype, legend_text, idx):
    """
    Decide a hex color (RRGGBB) for one series:
      - special_label_colors override everything
      - then per-type explicit TYPE_LABEL_COLORS
      - then type palette from COLOR_MAP -> PALETTES
        * 'misc': indices 0..9 use palette (cycle if short), idx>=10 -> deterministic
        * others: palette by index; if exhausted -> deterministic
    """
    norm = _normalize_label(legend_text)
    # 1) specials
    if norm in special_label_colors:
        return special_label_colors[norm]

    # 2) per-type explicit mapping
    tmap = type_label_colors.get((qtype or "").lower(), {})
    if norm in tmap:
        return tmap[norm]

    # 3) palette by type
    t = (qtype or "").lower()

    if t == "misc":
        base = TYPE_PALETTES.get("misc", [])
        if idx < 10:
            if base:
                return base[idx % len(base)]
            return _deterministic_hex(f"misc:{norm}:{idx}")
        return _deterministic_hex(f"misc:{norm}:{idx}")

    # non-misc types
    pal_id = color_map.get(t)
    if pal_id is not None:
        pal = palettes.get(pal_id, [])
        if idx < len(pal):
            return pal[idx]
        return _deterministic_hex(f"{t}:{norm}:{idx}")

    # unknown type: fall back to deterministic
    return _deterministic_hex(f"unknown:{t}:{norm}:{idx}")

