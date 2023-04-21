import PySimpleGUI as sg
import webbrowser
import webview
import time
import pandas as pd
import sys

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

# Creating frames
# top_banner = [[sg.Text('Precision Genetics', font='Any 20', background_color='#1B2838', enable_events=True, grab=False), sg.Push(background_color='#1B2838'),
#                sg.Text('Reporting Tools', font='Any 20', background_color='#1B2838')],]

# Creating tabs
tab_al_typ = [[sg.T('Proceed to Thermofishers Allele Typer')],
              [sg.T('maybe we can upload genotyper file to fix sample names?')],
              [sg.Push(), sg.Button('ThermoFisher', size=(14, 2), button_color=('#000000', '#C2D4D8')), sg.Push()]]

tab_diplo = [[sg.T('Upload Files For Calculations')],
             [sg.Push(),sg.Button('Upload Files', size=(14, 2), button_color=('#000000', '#C2D4D8')),
              sg.Button('Export Results', size=(14, 2), button_color=('#000000', '#C2D4D8')), sg.Push()]]

tab_genx = [[sg.T('Convert files for Genxys upload')],
            [sg.Push(), sg.Button('Convert', size=(14, 2), button_color=('#000000', '#C2D4D8')), sg.Push()]]

# Creating layout with tabs
layout =[[sg.TabGroup([
            [sg.Tab('Allele Typer', tab_al_typ, expand_x=True, expand_y=True)],
            [sg.Tab('Diplotype Calculator', tab_diplo, expand_x=True, expand_y=True)],
            [sg.Tab('Genxys File Conversion', tab_genx, expand_x=True, expand_y=True)]
         ], background_color=DARK_HEADER_COLOR, tab_location='topleft', expand_x=True, expand_y=True)]]

window = sg.Window('Reporting Tools', layout, resizable=True)

while True:
        event, values = window.read()
        print(event, values)
        if event == sg.WIN_CLOSED:
            break
        elif event == 'ThermoFisher':
            webview.create_window('Thermofisher', 'https://apps.thermofisher.com/alleletyper')
            webview.start()
window.close()