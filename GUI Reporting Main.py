import PySimpleGUI as sg
import webview
import pandas as pd
import csv
import os
import operator
import pyperclip
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox

# Setting the theme
sg.theme('Dark Blue')

#############################################################################
# Scripts for working with .CSVs
# This will contain code for creating an interactive window to work with .csv's
working_directory = os.getcwd()
csv.field_size_limit(2147483647)   # Allows for tables to be as large as needed


# Allows the csv table to be sorted in various ways
def sort_table(table, cols, descending=False):
    for col in reversed(cols):
        try:
            table = sorted(table, key=operator.itemgetter(col), reverse=descending)
        except Exception as e:
            sg.popup_error('Error in sorting the table', 'Exception in sorting the table', e)
    return table

##################################################################################################################
# CSV reader functions
# This function interacts with the CSV once it is uploaded
# When uploaded, it will trim the columns to only what is prevalent for reporting
# Rows are skipped so that the header is missed in the alleletyper export file
def read_csv_file(filename):
    data = []
    allele_cols = ['sample ID', 'CYP2D6', 'CYP2C9', 'SLCO1B1', 'CYP2B6', 'CYP3A4', 'CYP3A5', 'VKORC1']
    if filename is not None and filename != '':
        try:
            with open(filename, 'r', newline='', encoding=None) as infile:
                reader = csv.reader(infile)
                header_list = []
                for row in reader:
                    if any(col in allele_cols for col in row):
                        header_list = row
                        break
                allele_indices = [header_list.index(col) for col in allele_cols if col in header_list]
                next(reader)  # move the reader to the next line to read the data
                data = [[row[i] if i < len(row) else '' for i in allele_indices] for row in reader]
        except Exception as e:
            print(e)
            sg.popup_error('Error In Reading File', e)
        # Reset allele_cols to default value
        allele_cols = ['sample ID', 'CYP2D6', 'CYP2C9', 'SLCO1B1', 'CYP2B6', 'CYP3A4', 'CYP3A5', 'VKORC1']
        # Update allele_cols to include only columns present in the CSV file
        allele_cols = [header_list[i] for i in allele_indices if i < len(header_list)]
        header_list = [col for col in header_list if col in allele_cols]
    return data, allele_cols


#####################################################################################################################
# Setting up
######################################################################################################################
# Tab Layouts

tab_al_typ = [[sg.T('Proceed to Thermofishers Allele Typer', expand_x=True, expand_y=True, justification='center')],
              [sg.Push(), sg.T('This will direct you to ThermoFishers website, in order to complete this step.', justification='center'), sg.Push()],
              [sg.Push(), sg.Button('ThermoFisher', size=(14, 2), button_color=('#000000', '#C2D4D8')), sg.Push()]]

tab_diplo = [[sg.Push(), sg.T('Upload Your AlleleTyper Export.', expand_x=True, expand_y=True)],
             [sg.Push(), sg.T('You will be able to view, calculate the diplotype, and save a copy of the file.', expand_x=True, expand_y=True)],
             [sg.Push(), sg.InputText(key="-FILE_PATH-"), sg.FileBrowse(initial_folder=os.path.dirname(__file__), file_types=('CSV Files', '*.csv')), sg.Push()],
             [sg.Push(), sg.Button('Submit', key='-SUBMIT-', button_color=('#000000', '#C2D4D8'), bind_return_key=True), sg.Button('Cancel', button_color=('#000000', '#C2D4D8')), sg.Push()]]

controls = ['NA12878', 'NA18573', 'NA18971', 'NA19143', 'NA19144']

tab_ctrl = [[sg.Push(), sg.T('This tool is used to extract the sections needed to update the Control Batch Log from the Genotyper Export.', expand_x=True, expand_y=True)],
             [sg.Push(), sg.T('Upload The Genotyper Export', expand_x=True, expand_y=True)],
             [sg.Push(), sg.T('Please Choose Control Used In Batch.', expand_x=True, expand_y=True)],
             [sg.Push(), sg.OptionMenu(values=controls, key='-CONTROL-'), sg.Push()],
             [sg.Push(), sg.T('Information will be copied to your clipboard', expand_x=True, expand_y=True)],
             [sg.Push(), sg.InputText(key="-GENO_XPORT-"), sg.FileBrowse(initial_folder=os.path.dirname(__file__), file_types=[('CSV Files', '*.csv'), ('All Files', '*.*')]), sg.Push()],
             [sg.Push(), sg.Button('Submit', key='-SUBCTRL-', button_color=('#000000', '#C2D4D8'), bind_return_key=True), sg.Button('Cancel', button_color=('#000000', '#C2D4D8')), sg.Push()]]

tab_genx = [[sg.Push(), sg.T('Choose Conversion Method For Genxys', expand_x=True, expand_y=True, justification='center'), sg.Push()],
            [sg.Push(), sg.Button('Conversion For .CNV', key='-GEN-', size=(25, 2), button_color=('#000000', '#C2D4D8')), sg.Push()],
            [sg.Push(), sg.Button('Conversion For .OA', key='-OVA-', size=(25, 2), button_color=('#000000', '#C2D4D8')), sg.Push()]]

################################################################################################################
# Creating main window Layout
# This will be the homebase for workflow
layout = [[sg.Stretch()],
          [sg.TabGroup([
            [sg.Tab('Allele Typer', tab_al_typ)],
            [sg.Tab('Diplotype Calculator', tab_diplo)],
            [sg.Tab('Batch Control Log', tab_ctrl)],
            [sg.Tab('Genxys File Conversions', tab_genx)]
          ], background_color='#1B2838', tab_location='topleft', expand_x=True, expand_y=True)],
          [sg.Stretch()]]

window = sg.Window('Reporting Tools', layout, resizable=True, finalize=True)

##################################################################################################################
# Creating CSV Table Layout and function
def csv_window(file_path):
    data, header_list = read_csv_file(file_path)
    sg.popup_quick_message('Building your main window.... one moment....', background_color='#1c1e23', text_color='white', keep_on_top=True, font='_ 30')
    save_loc = None

# Diplotype Calculator Functions
    def diplotype_calc(file_path):
        nonlocal save_loc
        data, allele_cols = read_csv_file(file_path)


        # check which columns are in the csv, and add some if they are not present
        req_cols = ['sample ID', 'CYP2D6', 'CYP2C9']
        for col in req_cols:
            if col not in allele_cols:
                data.insert(allele_cols.index(allele_cols[-1])+1, [col] + ['' for _ in range(len(data)-2)])
                allele_cols.insert(allele_cols.index(allele_cols[-1])+1, col)


        allele_df = pd.DataFrame(data, columns=allele_cols)[['sample ID', 'CYP2D6', 'CYP2C9']]

        # Import CPIC dataframe from Google Cloud Storage into pandas dataframe
        # Currently only supported for CYP2D6 and CYP2C9
        cpic_data = f'https://storage.googleapis.com/pg-genotyping-cpic-2d6-2d9/2d6_2d9_joined.csv'
        cpic_df = pd.read_csv(cpic_data)

        # Merge the two DataFrames based on the common columns CYP2D6 and CYP2C9
        merged_df = pd.merge(allele_df, cpic_df, on=['CYP2D6', 'CYP2C9'])

        # Create two new columns to store the Metabolizer Status and Activity Score
        merged_df['Metabolizer Status'] = ''
        merged_df['Activity Score'] = ''

        # Create two dictionaries to store the metabolizer status and activity score for each call
        d6_status_dict = dict(zip(cpic_df['CYP2D6'], cpic_df['Metabolizer Status']))
        d9_status_dict = dict(zip(cpic_df['CYP2C9'], cpic_df['Metabolizer Status']))
        d6_activity_dict = dict(zip(cpic_df['CYP2D6'], cpic_df['Activity Score']))
        d9_activity_dict = dict(zip(cpic_df['CYP2C9'], cpic_df['Activity Score']))

        # Create two new columns to store the Metabolizer Status and Activity Score
        allele_df['2D6 Metabolizer Status'] = ''
        allele_df['2C9 Metabolizer Status'] = ''
        allele_df['Activity Score'] = ''

        # Iterate over each row of the allele DataFrame
        for index, row in allele_df.iterrows():
            # Get the cyp2d6 and cyp2c9 calls for the current sample ID
            cyp2d6_call = row['CYP2D6']
            cyp2c9_call = row['CYP2C9']

            # Use the dictionaries to get the metabolizer status and activity score for the current calls
            cyp2d6_status = d6_status_dict.get(cyp2d6_call, 'UND')
            cyp2c9_status = d9_status_dict.get(cyp2c9_call, 'UND')
            cyp2d6_activity = d6_activity_dict.get(cyp2d6_call, 'UND')
            cyp2c9_activity = d9_activity_dict.get(cyp2c9_call, 'UND')

            # Use the comparison results to populate the Metabolizer Status and Activity Score columns of the allele DataFrame
            allele_df.loc[index, '2D6 Metabolizer Status'] = f'{cyp2d6_status}'
            allele_df.loc[index, '2C9 Metabolizer Status'] = f'{cyp2c9_status}'
            allele_df.loc[index, 'Activity Score'] = f'{cyp2d6_activity}, {cyp2c9_activity}'

        # convert all of the updated data back into a list of lists
        updated_data = allele_df.values.tolist()

        # for saving the modified dataframe
        if save_loc:
            allele_df.to_csv(save_loc, index=False)

        ### ToDo: Fix this so it works for alleletyper exports that have fewer initial columns in their csv
        ### ToDo: Column name updates are an absolute mess. Come back and clean it up
        # This is a really cheap way of updating the column names
        # I cannot figure out how to update the headers correctly, so I have it update in a solid status
        updated_header = ['index', 'sample ID', 'CYP2D6', 'CYP2C9', '2D6 Metabolizer Status', '2C9 Metabolizer Status',
                          'Activity Score', '', '']
        window['-TABLE-'].update(values=updated_data)

        for i, col in enumerate(updated_header):
            window['-TABLE-'].Widget.heading(i, text=col)



    # ------ Window Layout ------
    layout = [  [sg.Text('Click a heading to sort on that column or enter a filter and click a heading to search for matches in that column')],
                [sg.Text(f'{len(data)} Records in table', font='_ 18')],
                [sg.Text(k='-RECORDS SHOWN-', font='_ 18')],
                [sg.Text(k='-SELECTED-')],
                [sg.T('Filter:'), sg.Input(k='-FILTER-', focus=True, tooltip='Not case sensative\nEnter value and click on a col header'),
                 sg.B('Reset Table', tooltip='Resets entire table to your original data'), sg.Button('Calculate Diplotype', key='-DIPLO-'),
                 sg.B('Save Copy', key='-SAVE-'),
                 sg.Checkbox('Sort Descending', k='-DESCENDING-'), sg.Checkbox('Filter Out (exclude)', k='-FILTER OUT-', tooltip='Check to remove matching entries when filtering a column'),
                 sg.Push()],
                [sg.Table(values=data, headings=header_list, max_col_width=25,
                        auto_size_columns=True, display_row_numbers=True, vertical_scroll_only=True,
                        justification='right', num_rows=50,
                        key='-TABLE-', selected_row_colors='red on yellow', enable_events=True,
                        expand_x=True, expand_y=True,
                        enable_click_events=True)],
                [sg.Sizegrip()]]

    # ------ Create Window ------
    window = sg.Window('CSV Table Display', layout, right_click_menu=sg.MENU_RIGHT_CLICK_EDITME_VER_EXIT,  resizable=True, size=(1000, 600), finalize=True)
    window.bind("<Control_L><End>", '-CONTROL END-')
    window.bind("<End>", '-CONTROL END-')
    window.bind("<Control_L><Home>", '-CONTROL HOME-')
    window.bind("<Home>", '-CONTROL HOME-')
    original_data = data        # save a copy of the data
    # ------ Event Loop ------
    ### ToDo: Fix the issues with column sorting. When Pressed, table data resets.
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
        # if is instance(event, tuple):
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
        elif event == '-DIPLO-' or event == 'calculate':
            diplotype_calc(file_path)
            window['-RECORDS SHOWN-'].update(f'{len(data)} Records Shown')
        elif event == '-SAVE-':
            save_loc = sg.popup_get_file('Save CSV', save_as=True, default_extension='.csv', file_types=(("CSV Files", "*.csv"),))
            if save_loc:
                diplotype_calc(file_path)
                sg.popup('File Saved Successfully!')
        elif event == '-CONTROL END-':
            window['-TABLE-'].set_vscroll_position(100)
        elif event == '-CONTROL HOME-':
            window['-TABLE-'].set_vscroll_position(0)
        elif event == 'Edit Me':
            sg.execute_editor(__file__)
        elif event == 'Version':
            sg.popup_scrolled(__file__, sg.get_versions(), location=window.current_location(), keep_on_top=True, non_blocking=True)
    window.close()

##################################################################################################################
# Creating a function for adding Genotyper Batch control information to clipboard
# A control is selected from the dropdown menu, and the genotyper export is uploaded
# The csv then checks the 'Sample Id' column to see if it matches the control that was selected.
# if everything matches, the specified columns and their cells are copied to the clipboard
# That data can then be pasted into the Batch Control Log workbook

def submit_ctrl_callback(values):
    selected_control = values['-CONTROL-']
    geno_file_path = values['-GENO_XPORT-']
    if geno_file_path.endswith('.csv'):
        # here we put the csv file into pandas dataframe
        df = pd.read_csv(geno_file_path, skiprows=17, skipfooter=1000, engine='python', encoding='ANSI')

        # create a dictionary mapping column names to column indices
        column_map = {'Sample ID': None, 'Assay Name': None, 'Assay ID': None, 'Gene Symbol': None,
                      'NCBI SNP Reference': None, 'Plate Barcode': None, 'Call': None}
        for i, col_name in enumerate(df.columns):
            if col_name in column_map:
                column_map[col_name] = i

        # The order we want the extracted data in
        extracted_data = df[['Sample ID', 'Assay Name', 'Assay ID', 'Gene Symbol', 'NCBI SNP Reference',
                             'Plate Barcode', 'Call']]
        for col_name, col_idx in column_map.items():
            if col_idx is not None and col_name not in extracted_data.columns:
                extracted_data.insert(loc=col_idx, column=col_name, value=df.iloc[:, col_idx])

        # filter rows that contain the selected control
        selected_rows = extracted_data[extracted_data['Sample ID'] == selected_control]
        if len(selected_rows) > 0:
            pyperclip.copy(selected_rows.to_csv(index=False, sep='\t'))
            sg.popup('Data Was Copied to Clipboard.')
        else:
            sg.popup('No Data Found For Selected Control.')
    else:
        sg.popup('Please Select A Valid CSV file.')

###############################################################################################################
# This is the file conversion function for turning the copy caller .txt export to a .cnv file
def convert_gen():
    # Creating a file selection window
    layout = [
        [sg.Text('Select a file to convert:')],
        [sg.Input(key='file_path'), sg.FileBrowse()],
        [sg.Text('Select a folder to save the converted file:')],
        [sg.Input(key='folder_path'), sg.FolderBrowse()],
        [sg.Button('Convert', key='CON'), sg.Button('Cancel')]
    ]
    window = sg.Window('Convert File', layout)

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Cancel'):
            break

        if event == 'CON':
            # Get the .txt file path and choose the folder the converted file will be exported to
            file_path = values['file_path']
            folder_path = values['folder_path']
            print(f"File path: {file_path}")
            print(f"Folder path: {folder_path}")
            # Checks if file is a .txt file, if not, it will return the error message
            if not file_path.endswith('.txt'):
                sg.popup('Invalid file selected. Please choose a .txt file.')
                continue

            # Convert file
            new_file_name, old_file_ext = os.path.splitext(os.path.basename(file_path))
            new_file_path = os.path.join(folder_path, new_file_name + '.cnv')
            with open(file_path, 'r') as f_in:
                file_data = f_in.read()
            # Conversion process
            converted_data = file_data.replace('.txt', '.cnv')
            with open(new_file_path, 'w') as f_out:
                f_out.write(converted_data)

            sg.popup('File converted and saved successfully!')
            break

    window.close()

###############################################################################################################
# This is the file conversion function for turning the Genotyper results into an .OA File
def convert_oa():
    # Creating a file selection window
    layout = [
        [sg.Text('Select a file to convert:')],
        [sg.Input(key='file_path'), sg.FileBrowse()],
        [sg.Text('Select a folder to save the converted file:')],
        [sg.Input(key='folder_path'), sg.FolderBrowse()],
        [sg.Button('Convert', key='CONOA'), sg.Button('Cancel')]
    ]
    window = sg.Window('Convert File', layout)

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Cancel'):
            break

        if event == 'CONOA':
            # Get the .txt file path and choose the folder the converted file will be exported to
            file_path = values['file_path']
            folder_path = values['folder_path']
            print(f"File path: {file_path}")
            print(f"Folder path: {folder_path}")
            # Checks if file is a .txt file, if not, it will return the error message
            if not file_path.endswith('.csv'):
                sg.popup('Invalid file selected. Please choose a .csv file.')
                continue

            # Convert file
            new_file_name, old_file_ext = os.path.splitext(os.path.basename(file_path))
            new_file_path = os.path.join(folder_path, new_file_name + '.oa')
            with open(file_path, 'r') as f_in:
                file_data = f_in.read()
            # Conversion process
            converted_data = file_data.replace('.csv', '.oa')
            with open(new_file_path, 'w') as f_out:
                f_out.write(converted_data)

            sg.popup('File converted and saved successfully!')
            break

    window.close()

###############################################################################################################
# Main Event Loop for the GUI

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':      # Closes program when needed
        break


    elif event == 'ThermoFisher':  # This creates an in program window for ThermoFishers AlleleTyper Software
        ### Todo: Find a way to implement our own AlleleTyper script and gain independence/distance from Thermo's software
        webview.create_window('Thermofisher', 'https://apps.thermofisher.com/alleletyper')
        webview.start()

    elif event == '-SUBMIT-':              # This 'Submit' button event is for the Diplotype Calculator Tab
        event, values = window.read()
        file_path = values['-FILE_PATH-']
        if file_path:
            csv_window(file_path)
            window.refresh()

    elif event == 'Cancel':                 # When 'Cancel' is pressed, it closes the program
        sg.popup('Closing Program')
        exit()

    elif event == '-SUBCTRL-':
        submit_ctrl_callback(values)

    elif event == '-GEN-':                  # This is For the Genxys file Conversion function on the file conversion tab
        convert_gen()

    elif event == '-OVA-':
        convert_oa()

if __name__ == '__main__':
    event, values = window.read()

window.close()
