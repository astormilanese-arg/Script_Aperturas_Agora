###Import Libraries

import pandas as pd
import win32com.client as win32
import win32gui
import win32con
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from openpyxl import load_workbook
from openpyxl.styles import numbers
import os
from .config import ppt_export_location, ppt_template_path, excel_save_location
from .models import question, TEMPLATE_MAP
import time
import psutil
from pywinauto import Application, keyboard
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.timings import Timings
import pyautogui
import pygetwindow as gw
import win32api
from .utilities import Determinar_tipo
import os, re, zipfile, shutil, tempfile, urllib.parse
import xml.etree.ElementTree as ET
from pathlib import Path
import binascii
import base64
import os



###Powerpoint_Revinculado###

def revincular_powerpoint (ppt_template_path, excel_save_location):
    
    # Open PowerPoint
    ppt_app = win32.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True
    presentation = ppt_app.Presentations.Open(ppt_template_path)
    print(ppt_template_path)
    # Go through all slides and shapes
    
    for slide in presentation.Slides:
        for shape in slide.Shapes:
            
            if shape.Type == 3:  # Picture
                
                try:
                    
                    source = shape.LinkFormat.SourceFullName
                    print(f"Found linked shape: {shape.Name} -> {source}")
                    #Changing the source
                    shape.LinkFormat.SourceFullName = excel_save_location
                    print(f"Updated link for shape: {shape.Name}")
                    
                except Exception as e:
                    
                    print(f"Skipping shape {shape.Name} due to error: {e}")

##Save and close if needed
    presentation.SaveAs(ppt_template_path)
    print(f"Presentation saved as: {ppt_template_path}")
    ppt_app.Quit()

def create_slides(preguntas_parseadas, excel_save_location, ppt_export_location, ppt_template_path):
    
    # Open Excel
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True
    workbook = excel_app.Workbooks.Open(excel_save_location)
    worksheet = workbook.Sheets(1)
    
    # Initialize applications
    ppt_app = win32.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True
    
    # Open presentation
    presentation = ppt_app.Presentations.Open(ppt_template_path)
    last_slide = presentation.Slides(presentation.Slides.Count)

    
    # Click center of screen to ensure PowerPoint is active
    screen_width, screen_height = pyautogui.size()
    
    pyautogui.hotkey('win', 'up')  # Simulates Win + Up
    
    for question in preguntas_parseadas:
        
        respuestas = question.respuestas
        nombre = question.nombre
        slide_index = TEMPLATE_MAP.get(question.tipo, 1)

        tipo = question.tipo
        
        print(question.ubicacion)
                
        time.sleep(0.1)
        
        # Duplicate slide acording to the type
        
        selected_slide = presentation.Slides(slide_index)
        
        new_slide = selected_slide.Duplicate()[0]  # Duplicate the selected slide
        
        new_slide.MoveTo(presentation.Slides.Count)
        
        nombre = question.nombre
        
        #voy al ultimo slide
        last_slide_index = presentation.Slides.Count
        ppt_app.ActiveWindow.View.GotoSlide(last_slide_index)
                
        time.sleep(0.5)
        # Find and modify shapes without selecting them
        pyautogui.click(screen_width // 2, screen_height // 2)
        time.sleep(0.5)  # Wait for click to register
        # Execute Alt → J → C → E shortcut
        pyautogui.hotkey('alt')  # Press Alt key
        time.sleep(0.1)
        pyautogui.press('j')     # Press J
        time.sleep(0.1)
        pyautogui.press('c')     # Press C
        time.sleep(0.1)
        pyautogui.press('e')     # Press E
        time.sleep(0.25)  # Wait for ribbon to update
        pyautogui.typewrite(question.ubicacion)
        time.sleep(0.25)
        pyautogui.press('enter')
        time.sleep(0.5)
        #pyautogui.click(1380, 80)
        pyautogui.keyDown('alt')
        pyautogui.press('space')
        pyautogui.keyUp('alt')
        time.sleep(0.25)  # Wait for suggestions to appear
        pyautogui.press('down')
        time.sleep(0.05)  # Wait for suggestions to appear       
        pyautogui.press('down')
        time.sleep(0.05)  # Wait for suggestions to appear      
        pyautogui.press('down')
        time.sleep(0.05)  # Wait for suggestions to appear      
        pyautogui.press('enter')
        time.sleep(0.5)
        change_slide_title(new_slide, nombre, tipo)
        time.sleep(0.5)

    # Save and close
    time.sleep(1)  # Extra delay before save
    presentation.SaveAs(ppt_export_location)
    workbook.Close(False)
    excel_app.Quit()
    ppt_app.Quit()
                
def change_slide_title(slide, nombre, tipo):
    if slide.Shapes.Count < 1:
        print("Slide has no shapes.")
        return

    shape = slide.Shapes(1)  # COM is 1-based
    if shape.HasTextFrame:
        
        nombre = determinar_titulo(nombre, tipo)
        
        shape.TextFrame.TextRange.Text = nombre
        print(f"Updated first shape's text to: {shape.TextFrame.TextRange.Text}")
    else:
        print("First shape has no text frame.")

def determinar_titulo(nombre, tipo):
    
    print("determinando titulo")
    
    nombre = nombre.replace("?", "").strip()
    nombre = nombre.replace("¿", "").strip()
    nombre = nombre.replace(".", "").strip()
    
    #sacamos las "q" iniciales
    if nombre.startswith("q"):
        nombre = nombre[1:]
    
    nombre = ''.join((x for x in nombre if not x.isdigit()))
    nombre = f"Apertura {nombre}"
    
    if "gestion" in nombre:
        _, _, after = nombre.partition(" gestion ")
        tema = after.strip()
        return f"Apertura evaluación de la gestión {tema}"
        
    if "imagen" in nombre:
        _, _, after = nombre.partition("de")
        tema = after.strip()
        return f"Apertura imagen de {tema}"
        
    if "clase social" in nombre:
        return f"Apertura clase social autopercibida"
        
    if tipo == "potencialidad":
        
        _, _, after = nombre.partition(" a ")
        tema = after.strip()
        return f"Apertura probabilidad de votar a {tema}"    
    
    if tipo == "estimulo":
        
        return f"Apertura evaluación del video: {nombre}"    
    
    if tipo == "principal motivo del voto":
        
        return f"Apertura principal motivo del voto"    
    
    if tipo == "afirmacion":
        _, _, after = nombre.partition(" : ")
        tema = after.strip()
        return f"Apertura grado de acuerdo con la afirmación: {nombre}"  
    
    return nombre



def borrar_plantillas(ppt_export_location):
    print("borrando")
    ppt = win32.Dispatch("PowerPoint.Application")
    presentation = ppt.Presentations.Open(ppt_export_location)

    for i in range(10, 0, -1):
        if i <= presentation.Slides.Count:
            presentation.Slides(i).Delete()

    presentation.Save()
    presentation.Close()
    ppt.Quit()
    
def slide_selector(question) -> int:
    """
    Given a question object, return the corresponding slide index
    defined in TEMPLATE_MAP. Defaults to 1 if not found.
    """
    tipo = question.tipo
    if not tipo:
        return 1  # safe fallback
    
    # Normalize just in case (strip spaces, lowercasing if needed)
    tipo = tipo.strip().lower()
    
    
    
    return TEMPLATE_MAP.get(tipo, 1)

###################### REWORKED ######################