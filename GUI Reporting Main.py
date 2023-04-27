import PySimpleGUI as sg
import webview
import datetime, time
import pandas as pd
import sys
import csv
import os
import operator


# Setting the theme
sg.theme('Dark Blue')

##################################################################################################################
# Import CPIC dataframe from Google Cloud Storage
# This will be used in diplotype calculator processes
# Currently only supported for CYP2D6 and CYP2C9
# ToDo: Create dataframe manipulation for the allele typer file import
cpic_data = f'https://storage.googleapis.com/pg-genotyping-cpic-2d6-2d9/2d6_2d9_joined.csv'
cpic_df = pd.read_csv(cpic_data)

##################################################################################################################
# Scripts for working with .CSVs
# This will contain code for creating an interactive window to work with .csv's
working_directory = os.getcwd()
csv.field_size_limit(2147483647)   # Allows for tables to be as large as needed

# Allows the csv table to be sorted
def sort_table(table, cols, descending=False):
    for col in reversed(cols):
        try:
            table = sorted(table, key=operator.itemgetter(col), reverse=descending)
        except Exception as e:
            sg.popup_error('Error in sort_table', 'Exception in sort_table', e)
    return table

def read_csv_file(filename):
    data = []
    header_list = []
    allele_cols = ['sample ID', 'CYP2D6', 'CYP2C9', 'CYP2C19', 'SLCO1B1', 'CYP2B6', 'CYP3A4', 'CYP3A5', 'VKORC1']
    if filename is not None and filename != '':
        try:
            with open(filename, 'r', newline='', encoding=None) as infile:
                reader = csv.reader(infile)
                for i in range(11):
                    next(reader)  # skip the first 10 rows
                header_list = next(reader)
                allele_indices = [header_list.index(col) for col in allele_cols if col in header_list]
                data = [[row[i] for i in allele_indices] for row in reader]
        except Exception as e:
            print(e)
            sg.popup_error('Error In Reading File', e)
    return data, allele_cols

######################################################################################################################
# Creating tabs
tab_al_typ = [[sg.T('Proceed to Thermofishers Allele Typer', expand_x=True, expand_y=True)],
              [sg.T('maybe we can upload genotyper file to fix sample names?')],
              [sg.Push(), sg.Button('ThermoFisher', size=(14, 2), button_color=('#000000', '#C2D4D8')), sg.Push()]]

tab_diplo = [[sg.T('Upload & View AlleleTyper Exports', expand_x=True, expand_y=True)],
             [sg.Push(), sg.InputText(key="-FILE_PATH-"), sg.FileBrowse(initial_folder=os.path.dirname(__file__), file_types=('CSV Files', '*.csv'))],
             [sg.Button('Submit', key='-SUBMIT-', button_color=('#000000', '#C2D4D8')), sg.Button('Cancel', button_color=('#000000', '#C2D4D8')), sg.Push()]]

### Callback function for the 'Submit' button on diplo tab
def handle_submit(file_path):
    csv_window(file_path)
def submit_callback(values):
    if values["-SUBMIT-"]:
        file_path = values["-FILE_PATH-"]
        handle_submit(file_path)


tab_genx = [[sg.T('Convert files for Genxys upload', expand_x=True, expand_y=True)],
            [sg.Push(), sg.Button('Convert', size=(14, 2), button_color=('#000000', '#C2D4D8')), sg.Push()]]

################################################################################################################
# Creating layout with tabs
layout = [[sg.Stretch()],
          [sg.TabGroup([
            [sg.Tab('Allele Typer', tab_al_typ)],
            [sg.Tab('Diplotype Calculator', tab_diplo)],
            [sg.Tab('Genxys File Conversion', tab_genx)]
          ], background_color='#1B2838', tab_location='topleft', expand_x=True, expand_y=True)],
          [sg.Stretch()]]

window = sg.Window('Reporting Tools', layout, resizable=True, finalize=True)

##################################################################################################################
#Creating CSV Table Layout and function
def csv_window(file_path):
    data, header_list = read_csv_file(file_path)

    sg.popup_quick_message('Building your main window.... one moment....', background_color='#1c1e23', text_color='white', keep_on_top=True, font='_ 30')

    # ------ Window Layout ------
    layout = [  [sg.Text('Click a heading to sort on that column or enter a filter and click a heading to search for matches in that column')],
                [sg.Text(f'{len(data)} Records in table', font='_ 18')],
                [sg.Text(k='-RECORDS SHOWN-', font='_ 18')],
                [sg.Text(k='-SELECTED-')],
                [sg.T('Filter:'), sg.Input(k='-FILTER-', focus=True, tooltip='Not case sensative\nEnter value and click on a col header'),
                 sg.B('Reset Table', tooltip='Resets entire table to your original data'),
                 sg.Checkbox('Sort Descending', k='-DESCENDING-'), sg.Checkbox('Filter Out (exclude)', k='-FILTER OUT-', tooltip='Check to remove matching entries when filtering a column'), sg.Push()],
                [sg.Table(values=data, headings=header_list, max_col_width=25,
                        auto_size_columns=True, display_row_numbers=True, vertical_scroll_only=True,
                        justification='right', num_rows=50,
                        key='-TABLE-', selected_row_colors='red on yellow', enable_events=True,
                        expand_x=True, expand_y=True,
                        enable_click_events=True)],
                [sg.Sizegrip()]]

    # ------ Create Window ------
    window = sg.Window('CSV Table Display', layout, right_click_menu=sg.MENU_RIGHT_CLICK_EDITME_VER_EXIT,  resizable=True, size=(800, 600), finalize=True)
    window.bind("<Control_L><End>", '-CONTROL END-')
    window.bind("<End>", '-CONTROL END-')
    window.bind("<Control_L><Home>", '-CONTROL HOME-')
    window.bind("<Home>", '-CONTROL HOME-')
    original_data = data        # save a copy of the data
    # ------ Event Loop ------
    while True:
        event, values = window.read()
        # print(event, values)
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        if values['-TABLE-']:           # Show how many rows are slected
            window['-SELECTED-'].update(f'{len(values["-TABLE-"])} rows selected')
        else:
            window['-SELECTED-'].update('')
        if event[0] == '-TABLE-':
        # if isinstance(event, tuple):
            filter_value = values['-FILTER-']
            # TABLE CLICKED Event has value in format ('-TABLE=', '+CLICKED+', (row,col))
            if event[0] == '-TABLE-':
                if event[2][0] == -1 and event[2][1] != -1:  # Header was clicked and wasn't the "row" column
                    col_num_clicked = event[2][1]
                    # if there's a filter, first filter based on the column clicked
                    if filter_value not in (None, ''):
                        filter_out = values['-FILTER OUT-']     # get bool filter out setting
                        new_data = []
                        for line in data:
                            if not filter_out and (filter_value.lower() in line[col_num_clicked].lower()):
                                new_data.append(line)
                            elif filter_out and (filter_value.lower() not in line[col_num_clicked].lower()):
                                new_data.append(line)
                        data = new_data
                    new_table = sort_table(data, (col_num_clicked, 0), values['-DESCENDING-'])
                    window['-TABLE-'].update(new_table)
                    data = new_table
                    window['-RECORDS SHOWN-'].update(f'{len(new_table)} Records shown')
                    window['-FILTER-'].update('')           # once used, clear the filter
                    window['-FILTER OUT-'].update(False)  # Also clear the filter out flag
        elif event == 'Reset Table':
            data = original_data
            window['-TABLE-'].update(data)
            window['-RECORDS SHOWN-'].update(f'{len(data)} Records shown')
        elif event == '-CONTROL END-':
            window['-TABLE-'].set_vscroll_position(100)
        elif event == '-CONTROL HOME-':
            window['-TABLE-'].set_vscroll_position(0)
        elif event == 'Edit Me':
            sg.execute_editor(__file__)
        elif event == 'Version':
            sg.popup_scrolled(__file__, sg.get_versions(), location=window.current_location(), keep_on_top=True, non_blocking=True)
    window.close()



###############################################################################################################

#Event Loop
while True:
    event, values = window.read()
    print(event, values)
    if event == sg.WIN_CLOSED:
        break
    elif event == 'ThermoFisher':                       #This creates an in program window for ThermoFishers AlleleTyper Software
        ### Todo: Find a way to implement our own AlleleTyper script and gain independence/distance from Thermo's software
        webview.create_window('Thermofisher', 'https://apps.thermofisher.com/alleletyper')
        webview.start()
    elif event == '-SUBMIT-':
        event, values = window.read()
        file_path = values['-File_PATH-']
        csv_window(file_path)
    elif event == 'Cancel':
        sg.popup('Closing Program')
        exit()

if __name__ == '__main__':
    event, values = window.read()

window.close()

        ###ToDo: Add code to convert diplotype calculator/alleletyper files for Genxys