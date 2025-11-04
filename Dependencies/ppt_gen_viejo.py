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
from .models import question
import time
import psutil
from pywinauto import Application, keyboard
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.timings import Timings
import pyautogui
import pygetwindow as gw
import win32api

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


def create_slides1DEPRE(preguntas_parseadas, excel_save_location, ppt_export_location, ppt_template_path):
    # Open PowerPoint
    ppt_app = win32.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True

    # Open the template presentation
    presentation = ppt_app.Presentations.Open(ppt_template_path)
    print(f"Opened PowerPoint template: {ppt_template_path}")

    # Get the last slide in the presentation
    last_slide_index = presentation.Slides.Count  # Get last slide index
    last_slide = presentation.Slides(last_slide_index)

    # Open Excel (keep the same instance for all updates)
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = False  # Set to True if you want Excel to be visible
    workbook = excel_app.Workbooks.Open(excel_save_location)
    worksheet = workbook.Sheets(1)  # Modify if necessary to choose the correct sheet

    for question in preguntas_parseadas:
        # Duplicate the last slide
        new_slide = last_slide.Duplicate()[0]  # [0] to get the slide from SlideRange
        print(question.nombre, question.ubicacion )
        # Find pictures in the new slide and update their links
        for shape in new_slide.Shapes:
            if shape.Type == 3:  # Picture (Linked Excel Data)
                try:
                    # Log original link
                    print(f"Original link: {shape.LinkFormat.SourceFullName}")
                    
                    # Update the link to the new Excel file (point to the new data series)
                    shape.LinkFormat.SourceFullName = excel_save_location

                    # Now, set the new Excel data range (ubicacion)
                    range_to_use = worksheet.Range(question.ubicacion)

                    # Update the link and make sure the picture reflects the new Excel range
                    shape.LinkFormat.AutoUpdate = 2  # Enable auto-update of the link
                    # Force the refresh (this may help to load the range properly)
                    shape.LinkFormat.Update()

                    print(f"Updated linked picture data to range {question.ubicacion} for {shape.Name}")

                except Exception as e:
                    print(f"Could not update link for picture {shape.Name}: {e}")

    # Save the updated PowerPoint as a new file
    presentation.SaveAs(ppt_export_location)
    #ppt_app.Quit()

    # Close Excel file after updates
    workbook.Close(SaveChanges=False)
    excel_app.Quit()
    
def create_slides3DEPRE(preguntas_parseadas, excel_save_location, ppt_export_location, ppt_template_path):
    # Open PowerPoint
    ppt_app = win32.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True
    presentation = ppt_app.Presentations.Open(ppt_template_path)
    
    # Get the template slide (last slide)
    last_slide = presentation.Slides(presentation.Slides.Count)
    
    # Open Excel (visible for debugging)
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True
    workbook = excel_app.Workbooks.Open(excel_save_location)
    
    for question in preguntas_parseadas:
        # Duplicate the template slide
        new_slide = last_slide.Duplicate()[0]
        time.sleep(2)
        
        # Activate the slide window (critical for selection)
        presentation.Windows(1).View.GotoSlide(new_slide.SlideIndex)
        time.sleep(1)
        
        for shape in new_slide.Shapes:
            if shape.HasChart == -1:  # Check if shape is a chart
                try:
                    # Get the chart object
                    chart = shape.Chart
                    
                    # Activate the chart data (this will open Excel)
                    print(f"Activating chart data for range: {question.ubicacion}")
                    chart.ChartData.Activate()
                    time.sleep(2)  # Wait for Excel to open
                    
                    # Programmatically update the range
                    chart_data = chart.ChartData
                    chart_workbook = chart_data.Workbook
                    chart_worksheet = chart_workbook.Sheets(1)
                    chart_range = chart_worksheet.Range(question.ubicacion)
                    chart.SetSourceData(chart_range)
                    
                    print(f"Updated chart range to: {question.ubicacion}")
                    
                    # Close the data workbook
                    chart_workbook.Close(True)
                    time.sleep(1)
                    
                except Exception as e:
                    print(f"Error updating chart: {str(e)}")
                    continue
    
    # Save and clean up
    presentation.SaveAs(ppt_export_location)
    workbook.Close(False)
    excel_app.Quit()
    ppt_app.Quit()
    
def create_slides2(preguntas_parseadas, excel_save_location, ppt_export_location, ppt_template_path):
    
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
        
        print(question.ubicacion)
        
        type =  question.tipo
        
        time.sleep(1)
        
        # Duplicate slide acording to the type
        
        selected_slide = presentation.Slides(slide_selector(type))  # Example: selecting the second slide in the presentation [USAR UNA FUNCION QUE ASIGNE UN NUMERO DEPENDIENDO DEL TIPO]
        
        new_slide = selected_slide.Duplicate()[0]  # Duplicate the selected slide
        
        new_slide.MoveTo(presentation.Slides.Count)
        
        nombre = question.nombre
        
        #voy al ultimo slide
        last_slide_index = presentation.Slides.Count
        ppt_app.ActiveWindow.View.GotoSlide(last_slide_index)
                
        time.sleep(1)
        # Find and modify shapes without selecting them
        pyautogui.click(screen_width // 2, screen_height // 2)
        time.sleep(1)  # Wait for click to register
        # Execute Alt → J → C → E shortcut
        pyautogui.hotkey('alt')  # Press Alt key
        time.sleep(0.2)
        pyautogui.press('j')     # Press J
        time.sleep(0.2)
        pyautogui.press('c')     # Press C
        time.sleep(0.2)
        pyautogui.press('e')     # Press E
        time.sleep(0.5)  # Wait for ribbon to update
        pyautogui.typewrite(question.ubicacion)
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(1)
        #pyautogui.click(1380, 80)
        pyautogui.keyDown('alt')
        pyautogui.press('space')
        pyautogui.keyUp('alt')
        time.sleep(0.5)  # Wait for suggestions to appear
        pyautogui.press('down')
        time.sleep(0.1)  # Wait for suggestions to appear       
        pyautogui.press('down')
        time.sleep(0.1)  # Wait for suggestions to appear      
        pyautogui.press('down')
        time.sleep(0.1)  # Wait for suggestions to appear      
        pyautogui.press('enter')
        time.sleep(1)
        change_slide_title(new_slide, nombre)
        time.sleep(1)

    # Save and close
    time.sleep(1)  # Extra delay before save
    presentation.SaveAs(ppt_export_location)
    workbook.Close(False)
    excel_app.Quit()
    ppt_app.Quit()
                
def change_slide_title(slide, nombre):
    if slide.Shapes.Count < 1:
        print("Slide has no shapes.")
        return

    shape = slide.Shapes(1)  # COM is 1-based
    if shape.HasTextFrame:
        shape.TextFrame.TextRange.Text = nombre
        print(f"Updated first shape's text to: {shape.TextFrame.TextRange.Text}")
    else:
        print("First shape has no text frame.")
        
def slide_selector(type):
    
    slide = 5
    
    if type == "valorativa":
        slide = 1
        
    if type == "acuerdo":
        slide = 2

    if type == "afirmativa":
        slide = 3
        
    if type == "potencialidad":
        slide = 4

    if type == "conocimiento":
        slide = 6
        
        
    return slide