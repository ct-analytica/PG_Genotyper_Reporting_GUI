import PySimpleGUI as sg
import webview
import datetime, time
import pandas as pd
import sys
import csv, os

# Import CPIC dataframe from Google Cloud Storage
cpic_data = f'https://storage.googleapis.com/pg-genotyping-cpic-2d6-2d9/2d6_2d9_joined.csv'
cpic_df = pd.read_csv(cpic_data)

# CSV workflow Scripts
working_directory = os.getcwd()
def convert_csv_array(csv_address):
    file = open(csv_address)
    reader = csv.DictReader(file)
    header = reader.fieldnames
    rows = [row for row in reader]
    file.close()
    return header, rows

# Setting the theme
theme_dict = {'BACKGROUND': '#2B475D',
              'TEXT': '#FFFFFF',
              'INPUT': '#F2EFE8',
              'TEXT_INPUT': '#000000',
              'SCROLL': '#F2EFE8',
              'BUTTON': ('#000000', '#C2D4D8'),
              'PROGRESS': ('#FFFFFF', '#C7D5E0'),
              'BORDER': 0,'SLIDER_DEPTH': 0, 'PROGRESS_DEPTH': 0}

sg.theme_add_new('Dashboard', theme_dict)
sg.theme('Dashboard')

BORDER_COLOR = '#C7D5E0'
DARK_HEADER_COLOR = '#1B2838'
BPAD_TOP = ((20,20), (20, 10))
BPAD_LEFT = ((20,10), (0, 0))
BPAD_LEFT_INSIDE = (0, (10, 0))
BPAD_RIGHT = ((10,20), (10, 0))

# Creating tabs
tab_al_typ = [[sg.T('Proceed to Thermofishers Allele Typer', expand_x=True, expand_y=True)],
              [sg.T('maybe we can upload genotyper file to fix sample names?')],
              [sg.Push(), sg.Button('ThermoFisher', size=(14, 2), button_color=('#000000', '#C2D4D8')), sg.Push()]]

tab_diplo = [[sg.T('Upload AlleleTyper Exports', expand_x=True, expand_y=True)],
             [sg.Push(), sg.InputText(key="-FILE_PATH-"), sg.FileBrowse(initial_folder=working_directory, file_types=('CSV Files', '*.csv'))],
             [sg.Button('Submit', button_color=('#000000', '#C2D4D8')), sg.Exit(), sg.Push()]]

tab_genx = [[sg.T('Convert files for Genxys upload', expand_x=True, expand_y=True)],
            [sg.Push(), sg.Button('Convert', size=(14, 2), button_color=('#000000', '#C2D4D8')), sg.Push()]]

# Creating layout with tabs
layout = [[sg.Stretch()],
          [sg.TabGroup([
            [sg.Tab('Allele Typer', tab_al_typ)],
            [sg.Tab('Diplotype Calculator', tab_diplo)],
            [sg.Tab('Genxys File Conversion', tab_genx)]
          ], background_color=DARK_HEADER_COLOR, tab_location='topleft', expand_x=True, expand_y=True)],
          [sg.Stretch()]]

window = sg.Window('Reporting Tools', layout, resizable=True, finalize=True)

while True:
    event, values = window.read()
    print(event, values)
    if event == sg.WIN_CLOSED:
        break
    elif event == 'ThermoFisher':
        webview.create_window('Thermofisher', 'https://apps.thermofisher.com/alleletyper')
        webview.start()
    elif event == 'Submit':
        csv_address = values["-FILE_PATH-"]
        header, rows = convert_csv_array(csv_address)

        # Creating the layout for the CSV table
        table_layout = [[sg.Table(values=rows[:25], headings=header, display_row_numbers=True, auto_size_columns=True,
                                  num_rows=min(25, len(rows)), key='-TABLE-', enable_events=True)],
                        [sg.Button('Update', button_color=('#000000', '#C2D4D8')), sg.Exit(), sg.Push()]]

        # Create new window to display CSV as table
        table_window = sg.Window('Editable CSV Table', table_layout, resizable=True, finalize=True)

        # Loop to keep table window open and update values
        while True:
            event, values = table_window.read()
            values_list = list(values.values())
            if event == sg.WIN_CLOSED:
                break
            elif event == 'Update':
                updated_data = values['-TABLE-']
                updated_header = header
                # Update CSV with new values
                with open(csv_address, 'w', newline='') as csvfile:
                    csvwriter = csv.writer(csvfile)
                    csvwriter.writerow(updated_header)
                    csvwriter.writerows(updated_data)

                # Reload updated CSV
                header, rows = convert_csv_array(csv_address)

                # Update table with new data
                table_window['-TABLE-'].update(values=rows, headings=header, num_rows=min(25, len(rows)))

window.close()