
import pandas as pd
import os
from .utilities import  export_to_excel, apply_percentage_format, segmentar_dataframe, insert_header, normalize_df
from .config import excel_save_location, app_Location
from .models import question

stripped_excel_location = excel_save_location + "stripped.xlsx"
print(stripped_excel_location)

def parse_data(df):
    
    # Process the DataFrame
    processed_df = insert_header(df)
    print("preguntas encabezadas")
    
    Stripped_df = normalize_df(processed_df)
    
    #print(processed_df.head(20))

    # Export the processed DataFrame to Excel
    export_to_excel(processed_df, excel_save_location)
    
    # Export the Stripped DataFrame to Excel
    export_to_excel(Stripped_df, stripped_excel_location)

    #Apply percentage formatting
    apply_percentage_format(excel_save_location)

    apply_percentage_format(stripped_excel_location)

    questions = segmentar_dataframe(processed_df)
    
    return questions
