"""
This script requires the following libraries to be installed:

- numpy
- matplotlib
- scipy
- pandas
- PySimpleGUI
- pygetwindow
- pyautogui
- pillow (PIL)

You can install these libraries using pip:

pip install numpy matplotlib scipy pandas PySimpleGUI pygetwindow pyautogui pillow
"""
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib 
from scipy.stats import norm
import time
import datetime as dt
from datetime import date
import pandas as pd
import PySimpleGUI as sg
import math
import statistics as stat
import re
import os
import pygetwindow as gw
import pyautogui
from PIL import Image
from tkinter import Tk, filedialog

sg.theme('BlueMono')
sg.set_options(font='Arial 12')
unit=''
layout1=[
     [sg.Text('Insert measurement unit: ', key='-UNIT TXT-', visible=True),
      sg.Input(key='-UNIT-', size=(15,1), visible=True)   
     ],     
     [sg.Text('Evaluation method: ', expand_x=True, justification='left'),
      sg.Combo(['Type A', 'Type B'], default_value='Type A', enable_events=True, key='-TYPE MENU-')]
]

Type_A_leftcolumn=[
     [sg.Text('Uncertainty evaluation method employed: ', key='-SEL TYPE TEXT-', visible=False),
      sg.Text(' ', key='-SEL TYPE-',visible=False)
     ],
     [sg.Text('If the measured values are stored in a file, press the button to select the file: ', key='-SELFILE TEXT-', visible=False),
      sg.Button('Select file', key='-FILE BUTTON-', enable_events=True, visible=False)],
     [sg.Text('Selected file: ', key='-SEL FILE TXT-',visible=False),
      sg.Text(" ",key='-SEL FILE-', visible=False)],
     [sg.Text('Number of measured values: ', key='-NUM MEAS VAL TXT-', visible=False),
      sg.Text(" ", key='-NUM MEAS VAL-', visible=False)  
     ],
     [sg.Text('Insert calibration std. uncertainty: ', key='-CAL UNC TXT-', visible=False),
      sg.Input(key='-CAL UNC-', size=(5,1), visible=False),
      sg.Text(' ', key='-UNIT CAL-', visible=False),
      sg.Combo(['Normal', 'Uniform'], default_value='Normal', key='-CAL UNC MENU-', visible=False),
      sg.Button('ADD', key='-ADD UNC-', enable_events=True, visible=False),
      sg.Button('SKIP', key='-SKIP UNC-', enable_events=True, visible=False)],
           [sg.Text('Mean value: ', key='-MEANVAL TXT-', visible=False),
      sg.Text(' ', key='-MEANVAL-', visible=False),
      sg.Text(' ', key='-UNITMV-', visible=False)],
     [sg.Text('Standard uncertainty ',key='-STDUNC TXT-', visible=False),
      sg.Text(' ', key='-STDUNC-', visible=False),
      sg.Text(' ', key='-UNITSU-', visible=False)],
     [sg.Text('Interval with 95.45% coverage probability; ', key='-EXP UNC INT TXT-', visible=False),
      sg.Text(' ', key='-INI INT-', visible=False),
      sg.Text(' ', key='-UNIT2-', visible=False),
      sg.Text(', ', key='-VIRG-', visible=False),
      sg.Text(' ', key='-FIN INT-', visible=False),
      sg.Text(' ', key='-UNIT3-', visible=False)
     ],
     [sg.Text('Expanded uncertainty U=', key='-EXP UNC TXT-', visible=False),
      sg.Text(' ', key='-EXP UNC-', visible=False),
      sg.Text(' ', key='-UNIT4-', visible=False)
     ],
     [sg.Text('Coverage factor k=', key='-COV FACT TXT-', visible=False),
      sg.Text(' ', key='-COV FACT-', visible=False)
     ],
     [sg.Button('SAVE', key='-SAVE-',enable_events=True, visible=False),
      sg.Button('CLOSE', key='-CLOSE-', enable_events=True, visible=False)]
]
Type_A_rightcolumn=[
     [sg.Canvas(size=(300,250),key='-CANVAS1-', visible=False)],
     [sg.Canvas(size=(300,250),key='-CANVAS2-', visible=False)]
]
layout2=[
    [sg.Column(Type_A_leftcolumn),
    sg.Column(Type_A_rightcolumn)]
]
type_B_leftcolumn=[
    [sg.Text('Uncertainty evaluation method employed: ', key='-SEL TYPE TEXTB-', visible=True),
      sg.Text(' ', key='-SEL TYPEB-',visible=False)
     ],
    [sg.CalendarButton("Press to insert measurement date", close_when_date_chosen=True, target='-DATAMEAS-', location=(800,600), no_titlebar=False, format='%Y-%m-%d'),
     sg.Input(key='-DATAMEAS-',size=(15,1))],
    [sg.Text('Measured value: '),
     sg.Input(key='-MEASVAL-', size=(10,1)),
     sg.Text(' ', key='-UNITB-')],
    [sg.CalendarButton("Press to insert last calibration date", close_when_date_chosen=True, target='-DATACAL-', location=(800,600), no_titlebar=False, format='%Y-%m-%d'),
     sg.Input(key='-DATACAL-',size=(15,1))],
    [sg.Text('Expected Drift: '),
     sg.Input(key='-DRIFT VAL-',size=(10,1)),
     sg.Text(' ', key='-DRIFT UNIT-'),
     sg.Combo(['Absolute', 'per year', 'per month', 'per day'], default_value='Absolute', key='-DRIFT MENU-'),
     sg.Combo(['Deterministic', 'Probabilistic'], default_value='Probabilistic', key='-DRIFT PROC MENU-')],
    [sg.Text('Definitional std. uncertainty: '),
     sg.Input(key='-DEF UNC-', size=(10,1)),
     sg.Text(' ', key='-DEF UNC UNIT-'),
     sg.Combo(['Normal', 'Uniform'], default_value='Normal', key='-DEF UNC MENU-')],
    [sg.Text('Instrumental std. uncertainty: '),
     sg.Input(key='-INST UNC-', size=(10,1)),
     sg.Text(' ', key='-INST UNC UNIT-'),
     sg.Combo(['Normal', 'Uniform'], default_value='Normal', key='-INST UNC MENU-')],
     [sg.Text('Resolution: '),
      sg.Input(key='-RES VAL-', size=(10,1)),
      sg.Text(' ', key='-RES UNIT-')],    
     [sg.Button('NEXT', enable_events=True, key='-NEXT1-')],
     [sg.Text('Uncertainty estimation', key='-RESULT TXT-', visible=False)],
     [sg.Text('Elapsed time since last calibration: ', key='-EL TIM TXT-', visible=False),
      sg.Text(' ', key='-EXP TIM-', visible=False),
      sg.Text(' ', key='-EXP TIM UNIT-', visible= False)],
     [sg.Text('Estimated drift: ', key='-EST DRIFT TXT-', visible=False),
      sg.Text(' ', key='-EST DRIFT VAL-', visible=False),
      sg.Text(' ', key='-EST DRIFT UNIT-', visible=False)],
     [sg.Text('Corrected value for drift: ', key='-CORR VAL TXT-', visible= False),
      sg.Text(' ', key='-CORR VAL-', visible=False),
      sg.Text(' ', key='-CORR VAL UNIT-', visible=False)],
     [sg.Text('Evaluated standard uncertainty: ', key='-STD UNC TXT-', visible=False),
      sg.Text(' ', key='-STD UNC VAL-', visible=False),
      sg.Text(' ', key='-STD UNC UNIT-', visible=False)],
     [sg.Text('Interval with 95.45% coverage probability; ', key='-EXP UNC INT TXTB-', visible=False),
      sg.Text(' ', key='-INI INTB-', visible=False),
      sg.Text(' ', key='-INI INTB UNIT-', visible=False),
      sg.Text(', ', key='-VIRGB-', visible=False),
      sg.Text(' ', key='-FIN INTB-', visible=False),
      sg.Text(' ', key='-FIN INTB UNIT-', visible=False)
     ],
     [sg.Text('Expanded uncertainty U=', key='-EXP UNC TXTB-', visible=False),
      sg.Text(' ', key='-EXP UNCB-', visible=False),
      sg.Text(' ', key='-EXP UNCB UNIT-', visible=False)
     ],
     [sg.Text('Coverage factor k=', key='-COV FACT TXTB-', visible=False),
      sg.Text(' ', key='-COV FACTB-', visible=False)
     ],
     [sg.Button('SAVE', key='-SAVEB-', enable_events=True, visible=False),
      sg.Button('CLOSE', key='-CLOSEB-', enable_events=True, visible=False)]      
]
type_B_rightcolumn = [
     [sg.Text('Histogram of the distribution of values that can reasonably be attributed to the measurand', key='-HISTB-', visible=False)],
     [sg.Canvas(size=(300,250),key='-CANVAS3-', visible=False)],
     [sg.Text('Cumulative probability function', key='-CUM PROB TXT-', visible=False)],
     [sg.Canvas(size=(300,250),key='-CANVAS4-', visible=False)]
]
layout3=[
    [sg.Column(type_B_leftcolumn),
    sg.Column(type_B_rightcolumn)]
]
matplotlib.use('TkAgg')

def draw_figure(canvas, figure):
   tkcanvas = FigureCanvasTkAgg(figure, canvas)
   tkcanvas.draw()
   widget = tkcanvas.get_tk_widget()
   widget.pack(side='top', fill='x', expand=True)
   #tkcanvas.get_tk_widget().pack(side='top', fill='x', expand=True)
   return tkcanvas

def delete_figure(fig):
    fig.get_tk_widget().forget()
    plt.close('all')

def create_and_plot_histogram(data, bins, unit):
    """
    Creates and plots a histogram for the given data.
    """
    fig, ax = plt.subplots(figsize=(5.5, 3.67))
    frequencies, edges, _ = ax.hist(data, bins, color='blue', edgecolor='blue', alpha=0.7, density=True)
    ax.set_title('Histogram')
    ax.set_xlabel(f'Measured Values [{unit}]')
    ax.set_ylabel('Relative Frequency')
    ax.grid(linestyle='--', alpha=0.7)
    return frequencies, edges, fig

def plot_cpf(X, Y, unit):
    """
    Creates and plots the Cumulative Probability Function (CPF).
    """
    fig, ax = plt.subplots(figsize=(5.5, 3.67))
    ax.plot(X, Y, 'b', label='CPF')
    ax.set_title('Cumulative Probability Function')
    ax.set_xlabel(f'Measured Values [{unit}]')
    ax.set_ylabel('Cumulative Probability')
    ax.legend()
    ax.grid(linestyle='--', alpha=0.7)
    return fig

def open_file():
    """Opens a dialog wondow to select a text file."""
    file_path = sg.popup_get_file(
        "Select a text file",
        file_types=(
            ("Text files", "*.txt"),
            ("All files", "*.*")
        )
    )

    if file_path:
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()
                data = [float(line.strip()) for line in lines if line.strip()]
                file.close()
                return file_path, data

        except Exception as e:
            window2["-TEXT-"].update(f"File open error: {e}")

def save_window_as_image(window_title, dati, cpfmat):
    # Find the window by title
    win = gw.getWindowsWithTitle(window_title)
    if not win:
        print("Window not found!")
        return
    win = win[0]

    # Get the window's position and size
    x, y, width, height = win.left, win.top, win.width, win.height

    # Open save dialog to choose output path
    root = Tk()
    root.withdraw()  # Hide the root window
    output_path = filedialog.asksaveasfilename(
        defaultextension=".png",
        filetypes=[("PNG files", "*.png"), ("All files", "*.*")]
    )
    root.destroy()

    if not output_path:
        print("Save operation canceled.")
        return

    # Capture the region using pyautogui
    screenshot = pyautogui.screenshot(region=(x, y, width, height))

    # Save the image
    screenshot.save(output_path)
    output_path_dati=output_path.replace(".png", "_dati.txt")
    with open(output_path_dati, "w") as file:
        for dato in dati:
            file.write(f"{dato}\n")

    output_path_cpf=output_path.replace(".png", "_cpf.txt")
    with open(output_path_cpf, "w") as file:
        for riga in cpfmat:
            file.write(";".join(map(str, riga)) + "\n")

tkcanvas=""
tkcanvas2=""
tkcanvas3=""
tkcanvas4=""

window1=sg.Window('Uncertainty evaluation',layout1, location=(50,100))


while True: 
    event, values= window1.read()

    if event == sg.WIN_CLOSED:
        break
    if event == '-TYPE MENU-':
        seltype=str(values['-TYPE MENU-'])
        unit=str(values['-UNIT-'])

        if seltype=='Type A':
            window2=sg.Window('Type A uncertainty evaluation', layout2, location=(600,100), finalize=True)
            while True:
                    window2['-SEL TYPE TEXT-'].update(visible=True)
                    window2['-SEL TYPE-'].update(seltype,visible=True)            
                    window2['-SELFILE TEXT-'].update(visible=True)
                    window2['-FILE BUTTON-'].update(visible=True)
                    event, values = window2.read()                
                    
                    if event==sg.WIN_CLOSED:
                        break

                    if event == '-FILE BUTTON-':
                        file_pathA, dataA=open_file()
                        window2['-SEL FILE TXT-'].update(visible=True)
                        window2['-SEL FILE-'].update(file_pathA, visible=True)
                        N=len(dataA)
                        window2['-NUM MEAS VAL TXT-'].update(visible=True)
                        window2['-NUM MEAS VAL-'].update(N,visible=True)

                        window2['-CAL UNC TXT-'].update(visible=True)
                        window2['-CAL UNC-'].update(visible=True)
                        window2['-UNIT CAL-'].update(unit, visible=True)
                        window2['-CAL UNC MENU-'].update(visible=True)
                        window2['-ADD UNC-'].update(visible=True)
                        window2['-SKIP UNC-'].update(visible=True)
                        event, values=window2.read()
                        if event == '-ADD UNC-':
                            calunc=float(values['-CAL UNC-'])
                            #adds randomly generated numbers to the measured values to consider calibration uncertainty
                            if values['-CAL UNC MENU-'] == 'Normal':
                                random_cal_err = np.random.normal(loc=0, scale=calunc, size=N)
                            else:
                                random_cal_err = np.random.uniform(-calunc*math.sqrt(3), calunc*math.sqrt(3), N) 
                            dataC=[]
                            dataC=[0 for j in range(N)]
                            for j in range(N):
                                dataC[j]=dataA[j]+random_cal_err[j]
                        elif event=='-SKIP UNC-':
                            dataC=dataA
                        meanval=stat.mean(dataC)
                        sigma=stat.stdev(dataC)
                        #Plot histogram
                        bins=int(N/10)
                        if tkcanvas != "":
                            delete_figure(tkcanvas)
                        frequencies, edges, plt.gcf = create_and_plot_histogram(dataC, bins, unit)          
                        tkcanvas=draw_figure(window2['-CANVAS1-'].TKCanvas, plt.gcf)
                        window2['-CANVAS1-'].update(visible=True)
                        window2['-MEANVAL TXT-'].update(visible=True)
                        window2['-MEANVAL-'].update("{:.4f}".format(meanval), visible=True)
                        window2['-UNITMV-'].update(' ['+unit+']', visible=True)
                        window2['-STDUNC TXT-'].update(visible=True)
                        window2['-STDUNC-'].update("{:.2f}".format(sigma), visible=True)
                        window2['-UNITSU-'].update(' ['+unit+']', visible=True)

                        #evaluates the cumulative probabiliy function

                        deltacl=(edges[bins]-edges[0])/bins
                        #areahist=np.sum(frequencies*deltacl)
                        cpf=np.zeros(bins)
                        cpf[1:]=np.cumsum(frequencies[1:]*deltacl)
                        X = edges[:-1]  # Remove the last edge to align with cpf length
                        cpfmatr = np.column_stack((X, cpf))
                        if len(X) != len(cpf):
                            print(f"Length mismatch: X={len(X)}, CPF={len(cpf)}")
                        else:
                            if tkcanvas2 != "":
                                delete_figure(tkcanvas2)  # Remove previous figure if exists
                        try:
                            fig = plot_cpf(X, cpf, unit)
                            tkcanvas2 = draw_figure(window2['-CANVAS2-'].TKCanvas, fig)
                            window2['-CANVAS2-'].update(visible=True)
                        except Exception as e:
                            print(f"Error plotting CPF: {e}")

                        #Evaluates expanded uncertainty with p=95.45%
                        covprob = 0.9545
                        iniprob = (1-covprob)/2
                        finprob = 1-iniprob
                        j=0
                        while cpf[j] < iniprob:
                            j = j+1
                        intini = X[j]
                        while cpf[j] < finprob:
                            j=j+1
                        intfin = X[j]
                        U = (intfin-intini)/2
                        k = U/sigma
                        window2['-EXP UNC INT TXT-'].update(visible=True)
                        window2['-INI INT-'].update("{:.4f}".format(intini), visible=True)
                        window2['-UNIT2-'].update(unit, visible=True)
                        window2['-VIRG-'].update(visible=True)
                        window2['-FIN INT-'].update("{:.4f}".format(intfin), visible=True)
                        window2['-UNIT3-'].update(unit, visible=True)
                        window2['-EXP UNC TXT-'].update(visible=True)
                        window2['-EXP UNC-'].update("{:.2f}".format(U), visible=True)
                        window2['-UNIT4-'].update(unit, visible=True)
                        window2['-COV FACT TXT-'].update(visible=True)
                        window2['-COV FACT-'].update("{:.3f}".format(k), visible=True)
                        window2['-SAVE-'].update(visible=True)
                        window2['-CLOSE-'].update(visible=True)
                        event, values = window2.read()
                        if event=='-SAVE-':
                            save_window_as_image('Type A uncertainty evaluation', dataC, cpfmatr)
                            break
                        if event=='-CLOSE-':
                            break
            window2.close()
            window1.close()

        elif seltype=='Type B':
            window3=sg.Window('Type B uncertainty evaluation', layout3, location=(600,100), finalize=True)
            window3['-SEL TYPEB-'].update(seltype,visible=True)
            window3['-UNITB-'].update(unit,visible=True)
            window3['-DRIFT UNIT-'].update(unit, visible=True)
            window3['-DEF UNC UNIT-'].update(unit, visible=True)
            window3['-INST UNC UNIT-'].update(unit, visible=True)
            window3['-RES UNIT-'].update(unit,visible=True)
            while True:
                event, values = window3.read()
                if event=='-NEXT1-':
                    measval=float(values['-MEASVAL-'])
                    dmeas=date.fromisoformat(values['-DATAMEAS-'])
                    dcal=date.fromisoformat(values['-DATACAL-'])
                    #Evaluates the time elapsed since calibration
                    exptimcal=(dmeas-dcal)
                    expdays=exptimcal.days
                    expmonths=expdays*100/30
                    expmonths=int(expmonths)
                    expmonths=expmonths/100
                    expyears=expdays*100/365
                    expyears=int(expyears)
                    expyears=expyears/100
                    #evaluates drift
                    driftval=float(values['-DRIFT VAL-'])
                    if values['-DRIFT MENU-'] == 'Absolute':
                        drift=driftval
                        exptim=''
                        timunit='NA'
                    elif values['-DRIFT MENU-'] == 'per year':
                        drift=driftval*expyears
                        exptim=expyears
                        timunit='years'
                    elif values['-DRIFT MENU-'] == 'per month':
                        drift=driftval*expmonths
                        exptim=expmonths
                        timunit='months'
                    elif values['-DRIFT MENU-'] == 'per day':
                        drift=driftval*expdays
                        exptim=expdays
                        timunit='days'
                    window3['-RESULT TXT-'].update(visible=True)
                    window3['-EL TIM TXT-'].update(visible=True)
                    window3['-EXP TIM-'].update(exptim, visible=True)
                    window3['-EXP TIM UNIT-'].update(timunit, visible=True)
                    window3['-EST DRIFT TXT-'].update(visible=True)
                    window3['-EST DRIFT VAL-'].update('{:.4f}'.format(drift), visible=True)
                    window3['-EST DRIFT UNIT-'].update(unit, visible=True)
                    #Evaluates the corrected value
                    if values['-DRIFT PROC MENU-'] == 'Deterministic':
                        corrval=measval-drift
                    else:
                        corrval=measval
                    window3['-CORR VAL TXT-'].update(visible=True)
                    window3['-CORR VAL-'].update("{:.3f}".format(corrval), visible=True)
                    window3['-CORR VAL UNIT-'].update(unit, visible=True)
                    #Standard uncertainty evaluation
                    if values['-DRIFT PROC MENU-'] == 'Probabilistic':
                        driftu = drift/(2*math.sqrt(3))
                    else:
                        driftu = 0
                    defu=float(values['-DEF UNC-'])
                    instu=float(values['-INST UNC-'])
                    resu=float(values['-RES VAL-'])/(2*math.sqrt(3))
                    u=math.sqrt(defu**2+instu**2+driftu**2+resu**2)
                    window3['-STD UNC TXT-'].update(visible=True)
                    window3['-STD UNC VAL-'].update('{:.2f}'.format(u), visible=True)
                    window3['-STD UNC UNIT-'].update(unit, visible=True)

                    #Evaluation of the expected distribution of the measured values by a Monte Carlo simulation
                    
                    N=20000
                    if values['-DRIFT PROC MENU-'] == 'Probabilistic':
                        err_drift = np.random.uniform(-drift, drift, N)
                    else:
                        err_drift = np.zeros(N)
                    if values['-DEF UNC MENU-'] == 'Uniform':
                        err_def = np.random.uniform(-defu*math.sqrt(3), defu*math.sqrt(3), N)
                    elif values['-DEF UNC MENU-'] == 'Normal':
                        err_def = np.random.normal(0, defu, N)
                    if values['-INST UNC MENU-'] == 'Uniform':
                        err_inst = np.random.uniform(-instu*math.sqrt(3), instu*math.sqrt(3), N)
                    elif values['-INST UNC MENU-'] == 'Normal':
                        err_inst = np.random.normal(0, instu, N)
                    res = float(values['-RES VAL-'])
                    err_res = np.random.uniform(-res/2, res/2, N)
                    distval = np.zeros(N)
                    for i in range(N):
                        distval[i] = corrval + err_drift[i] + err_def[i] + err_inst[i] + err_res[i]
                    
                    #Displays the histogram

                    bins=int(N/10)
                    window3['-HISTB-'].update(visible=True)
                    if tkcanvas3 != "":
                        delete_figure(tkcanvas)
                    frequencies_B, edges_B, plt.gcf = create_and_plot_histogram(distval, bins, unit)          
                    tkcanvas3=draw_figure(window3['-CANVAS3-'].TKCanvas, plt.gcf)
                    window3['-CANVAS3-'].update(visible=True)

                    #Evaluates the cumulative probability function

                    deltaclB = (edges_B[bins]-edges_B[0])/bins
                    cpfB=np.zeros(bins)
                    cpfB[1:]=np.cumsum(frequencies_B[1:]*deltaclB)
                    XB = edges_B[:-1]  # Remove the last edge to align with cpf length
                    cpfmatrB = np.column_stack((XB, cpfB))
                    if len(XB) != len(cpfB):
                        print(f"Length mismatch: X={len(X)}, CPF={len(cpf)}")
                    else:
                        if tkcanvas4 != "":
                            delete_figure(tkcanvas4)  # Remove previous figure if exists
                        try:
                            fig = plot_cpf(XB, cpfB, unit)
                            tkcanvas4 = draw_figure(window3['-CANVAS4-'].TKCanvas, fig)
                            window3['-CUM PROB TXT-'].update(visible=True)
                            window3['-CANVAS4-'].update(visible=True)
                        except Exception as e:
                            print(f"Error plotting CPF: {e}")
                        
                    #Evaluates expanded uncertainty with p=95.45%

                    covprob = 0.9545
                    iniprob = (1-covprob)/2
                    finprob = 1-iniprob
                    j=0
                    while cpfB[j] < iniprob:
                        j = j+1
                    intini = XB[j]
                    while cpfB[j] < finprob:
                        j=j+1
                    intfin = XB[j]
                    U_B = (intfin-intini)/2
                    k_B = U_B/u
                    window3['-EXP UNC INT TXTB-'].update(visible=True)
                    window3['-INI INTB-'].update("{:.3f}".format(intini), visible=True)
                    window3['-INI INTB UNIT-'].update(unit, visible=True)
                    window3['-VIRGB-'].update(visible=True)
                    window3['-FIN INTB-'].update("{:.3f}".format(intfin), visible=True)
                    window3['-FIN INTB UNIT-'].update(unit, visible=True)
                    window3['-EXP UNC TXTB-'].update(visible=True)
                    window3['-EXP UNCB-'].update("{:.2f}".format(U_B), visible=True)
                    window3['-EXP UNCB UNIT-'].update(unit, visible=True)
                    window3['-COV FACT TXTB-'].update(visible=True)
                    window3['-COV FACTB-'].update("{:.3f}".format(k_B), visible=True)
                    window3['-SAVEB-'].update(visible=True)
                    window3['-CLOSEB-'].update(visible=True)
                    event, values = window3.read()
                    if event=='-SAVEB-':
                        save_window_as_image('Type B uncertainty evaluation', distval, cpfmatrB)
                        break
                    if event=='-CLOSEB-':
                        break
                
                else:
                    window3.close()
                    break
            window3.close()
            window1.close()

window1.close()