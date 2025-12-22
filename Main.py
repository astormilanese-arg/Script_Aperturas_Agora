# ──────────────────────────────────────────────────────────────────────────────
# Imports
# ──────────────────────────────────────────────────────────────────────────────

import sys
print(">>> Running with Python:", sys.executable)

import time
import builtins
import xml.etree.ElementTree as ET

import pandas as pd

from Dependencies.config import (
    app_Location,
    ppt_template_path,
    excel_save_location,
    ppt_export_location,
)

from Dependencies.models import question, COLOR_MAP
from Dependencies.parser import parse_data
from Dependencies.Diagnostic import diagnose_pptx_after_edit
from Dependencies.ppt_gen import revincular_powerpoint
from Dependencies.ppt_gen_V3 import (
    update_charts_in_pptx,
    CopyTemplateSlides,
    unzip_pptx_clean,
    rezip_pptx,
)
from Dependencies.utilities import (
    close_excel_ppt,
    close_powerpoint,
    refresh_all_graphs,
)


# ──────────────────────────────────────────────────────────────────────────────
# UTF-8 STDOUT/STDERR configuration
# ──────────────────────────────────────────────────────────────────────────────

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


import win32com.client.gencache
import shutil, os, tempfile

gen_py = os.path.join(tempfile.gettempdir(), "gen_py")
if os.path.exists(gen_py):
    shutil.rmtree(gen_py)

# ──────────────────────────────────────────────────────────────────────────────
# Timestamped print (global replacement)
# ──────────────────────────────────────────────────────────────────────────────

_start_time = time.time()
builtins._original_print = builtins.print  # keep reference


def timestamped_print(*args, **kwargs):
    """
    Print with elapsed seconds since program start prefixed.
    """
    elapsed = time.time() - _start_time
    prefix = f"[{elapsed:8.2f}s]"
    builtins._original_print(prefix, *args, **kwargs)


# Replace print globally
builtins.print = timestamped_print


# ──────────────────────────────────────────────────────────────────────────────
# Pandas display options (developer-friendly)
# ──────────────────────────────────────────────────────────────────────────────

pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", 5)
pd.set_option("display.width", None)
pd.set_option("display.max_colwidth", None)


# ──────────────────────────────────────────────────────────────────────────────
# Orchestration
# ──────────────────────────────────────────────────────────────────────────────

def main():
    # Ensure Office processes are closed (twice, as in original script)
    close_excel_ppt()
    close_powerpoint()

    # Load application dataframe
    app_dataframe = pd.read_excel(app_Location, engine="openpyxl")

    # Parse into question objects
    preguntas_parseadas = parse_data(app_dataframe)
    print(preguntas_parseadas)

    # Inspect parsed questions
    for i, quest in enumerate(preguntas_parseadas):
        print(f"  Nombre: {quest.nombre}")
        print(f"  Respuestas: {quest.respuestas}")
        print(f"  Ubicacion: {quest.ubicacion}")
        print(f"  Tipo: {quest.tipo}")
        print(f" Titulo: {quest.titulo}")

    # Re-link template pictures/charts to Excel
    revincular_powerpoint(ppt_template_path, excel_save_location)

    # Duplicate slides by question and export presentation
    CopyTemplateSlides(ppt_template_path, preguntas_parseadas, ppt_export_location, visible=False)
    print(ppt_export_location)

    # Update charts in exported presentation; returns output path
    output_path = update_charts_in_pptx(ppt_export_location, preguntas_parseadas)

    summary = diagnose_pptx_after_edit(output_path, preguntas_parseadas, verbose=True)

    #Refresh all links/charts in-place
    refresh_all_graphs(output_path, with_window=True)

# ──────────────────────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    main()
