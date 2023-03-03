import PySimpleGUI as sg
import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate

# Instructions for use button.
def instructions():
    sg.popup_scrolled(f"HOW TO USE THE TEMPLATE GENERATOR ")

# Define the layout of the GUI
layout = [
    [sg.Text('Select a Word Template file:', size=(20, 1)),
     sg.InputText('Select a template from TEMPLATES Folder', key='WORD_TEMPLATE'),
     sg.FileBrowse(file_types=(("Word Template Files", "*.docx"),))],
    [sg.Text('Select an Excel file:', size=(20, 1)), sg.InputText('Select an Excel file', key='EXCEL_FILE'),
     sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),))],
    [sg.Exit(), sg.Button('Generate Documents'), sg.Button("Instructions")]
]

# Create the window with the defined layout
window = sg.Window('Document Generator', layout)

# Event loop
while True:
    event, values = window.read()
    if event in (sg.WINDOW_CLOSED,"Exit"):
        break
    if event == "Instructions":
        instructions()
    elif event == 'Generate Documents':
        # Get the selected file paths from the input fields
        word_template_path = values['WORD_TEMPLATE']
        excel_path = values['EXCEL_FILE']

        # Set the output directory
        base_dir = Path.home() / "Desktop"
        output_dir = base_dir / "OUTPUT"
        output_dir.mkdir(exist_ok=True)

        # Read in the Excel file and generate documents from each row
        df = pd.read_excel(excel_path, sheet_name="Sheet1")
        for record in df.to_dict(orient="records"):
            doc = DocxTemplate(word_template_path)
            doc.render(record)
            output_path = output_dir / f"{record['SITECODE']}.docx"
            doc.save(output_path)

        sg.popup(f"Documents generated in {output_dir}")

# Close the window
window.close()