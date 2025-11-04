from Dependencies.parser import parse_data
from Dependencies.utilities import close_excel_ppt, close_powerpoint
import pandas as pd
from Dependencies.config import app_Location, ppt_template_path, excel_save_location, ppt_export_location
from Dependencies.ppt_gen import revincular_powerpoint, create_slides, borrar_plantillas
from Dependencies.models import question, TEMPLATE_MAP
import xml.etree.ElementTree as ET


close_excel_ppt()
close_excel_ppt()

# Display options to show all data
pd.set_option('display.max_rows', None)  # Show all rows
pd.set_option('display.max_columns', 5)  # Show all columns
pd.set_option('display.width', None)  # Auto-detect terminal width
pd.set_option('display.max_colwidth', None)  # Show full column content

app_dataframe = pd.read_excel(app_Location, engine="openpyxl")

preguntas_parseadas = parse_data(app_dataframe)

#print((preguntas_parseadas))

for i, quest in enumerate(preguntas_parseadas):

    print(f"  Nombre: {quest.nombre}")
    print(f"  Respuestas: {quest.respuestas}")
    print(f"  Ubicacion: {quest.ubicacion}")  
    print(f"  Tipo: {quest.tipo}")
        
revincular_powerpoint(ppt_template_path, excel_save_location)

create_slides(preguntas_parseadas, excel_save_location, ppt_export_location, ppt_template_path)

borrar_plantillas(ppt_export_location)
