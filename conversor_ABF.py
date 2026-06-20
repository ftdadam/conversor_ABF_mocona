# -*- coding: utf-8 -*-
"""

@author: fdadam
"""

import os
import tkinter as tk
# from tkinter import *
from tkinter import filedialog
import pandas as pd
from datetime import datetime
from pathlib import Path
import time
import numpy as np
import ttkbootstrap as ttk

script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)


def logdate():
    return datetime.now().strftime("[%Y/%m/%d, %H:%M:%S]")

def openFiles():
    global filesList
    global checkVarList

    filesFromDiag = filedialog.askopenfilenames(initialdir=current_dir_StringVar,title='Seleccionar Inputs',
                                        filetypes=(("Archivos xlsx","*.xlsx"),))
    filesFromDiag = list(filesFromDiag)

    checkVarFromDiag = []
    for _ in range(len(filesFromDiag)):
        # print(_)
        aux = tk.BooleanVar()
        aux.set(True)
        checkVarFromDiag.append(aux)

    checkVarList = checkVarList + checkVarFromDiag

    filesList = filesList + filesFromDiag    
    displayFilenamesList = [f'{f.split(sep='/')[-1]}' for f in filesList]

    if(filesList):
        # Clear Frame
        for widget in fileChecklistFrame.winfo_children():
            widget.destroy()

        title_StringVar.set('Archivos Seleccionados:')
        
        for ptr in range(len(displayFilenamesList)):
            ttk.Checkbutton(fileChecklistFrame,
                            text=displayFilenamesList[ptr],
                            variable=checkVarList[ptr]).pack(anchor='w')

        # print([var.get() for var in checkVarList])

def openLog():
    global logfile
    global logfile_path
    logfile.close()
    os.system(f'start {logfile_path}')
    logfile = open(logfile_path,'a+')

def removeFiles():
    global filesList
    global checkVarList
    
    # Clear Frame
    for widget in fileChecklistFrame.winfo_children():
        widget.destroy()

    filesList = []
    checkVarList = []
    title_StringVar.set('No Hay Archivos de Entrada Seleccionados')

def tipo_canto(canto):
    if(isinstance(canto,str)):
        return int(canto.split('-')[0].strip())
    else:
        return np.nan

def definirCantos():
    global config_path
    global cantos_dir
    fileFromDiag = filedialog.askopenfilename(initialdir=current_dir_StringVar,title='Seleccionar Archivo de Cantos',
                                              filetypes=(("Archivos xlsx","*.xlsx"),))
    with open(config_path, "w") as file:
        file.write(fileFromDiag)  
    file.close()
    cantos_dir = fileFromDiag
    cantos_dir_StringVar.set(f'\U0001F6C8 Archivo de cantos: {cantos_dir}')  
    return 0

def convertDfScheimberg(df_og):
    global cantos_dir
    df = df_og.copy()
    
    # Switch LARGO y ANCHO if grain==N
    
    largo_aux = df['cutting-length'].copy().values
    ancho_aux = df['cutting-width'].copy().values
    
    for ptr in df.index:
        if( (df['grain'][ptr] == 'N') or (df['grain'][ptr] == 0)):
            largo_aux[ptr] = df['cutting-width'].copy().values[ptr]
            ancho_aux[ptr] = df['cutting-length'].copy().values[ptr]
            
            
    df['cutting-length'] = largo_aux
    df['cutting-width'] = ancho_aux
    
    # Drop Columns
    cols_to_keep = ['cutting-length','cutting-width','material','label',
                    '_top_','_right_','_bot_','_left_']
    
    cols_new_names = ['LARGO','ANCHO','TIPO DE PLACA','OBSERVACION',
                      'LARGO SUPERIOR','ANCHO DERECHO','LARGO INFERIOR','ANCHO IZQUIERDO']
    
    df = df[df.columns.intersection(cols_to_keep)]
    
    # Round
    cols = ['cutting-length','cutting-width']
    df[cols] = df[cols].round(0)
    # Mantener solo numero de canto
    edge_cols = ['_top_', '_right_', '_bot_', '_left_']
    for col in edge_cols:
        df[col] = df[col].apply(tipo_canto)

    # df['_top_'] = df['_top_'].apply(lambda x: tipo_canto(x))
    # df['_right_'] = df['_right_'].apply(lambda x: tipo_canto(x))
    # df['_bot_'] = df['_bot_'].apply(lambda x: tipo_canto(x))
    # df['_left_'] = df['_left_'].apply(lambda x: tipo_canto(x))


    df['label_aux'] = df[['_top_', '_right_', '_bot_', '_left_']].max(axis=1)

    df[edge_cols] = df[edge_cols].map(
        lambda x: "" if pd.isna(x) else 1
    )
    

    df_cantos = pd.read_excel(os.path.normpath(cantos_dir))
    canto_lookup = dict(zip(df_cantos['Canto_ID'], df_cantos['Canto_Nombre']))

    df['label'] = df['label'] + "_" + df['label_aux'].map(canto_lookup)

    # Rename & Reorder
    rename_dict = dict(zip(cols_to_keep, cols_new_names))
    
    df.rename(columns=rename_dict,inplace=True)
    
    cols_order = ['TIPO DE PLACA','LARGO','ANCHO','OBSERVACION',
                  'LARGO SUPERIOR','LARGO INFERIOR','ANCHO DERECHO','ANCHO IZQUIERDO']
    
    df = df[cols_order]
    
    # Insert column CANTIDAD (Value 1 in every row)
    df.insert(1,'CANTIDAD',value=[1]*len(df))
    # Insert column CODIGO EXTERNO (Value 18 in every row)
    df.insert(5,'CODIGO EXTERNO',value=[18]*len(df))
    # Insert column ROTACION (Value "" in every row)
    df.insert(6,'ROTACION',value=""*len(df))

    return df

def convertDfMocona(df_og):
    
    df = df_og.copy()
    
    # Switch LARGO y ANCHO if grain==N
    
    largo_aux = df['cutting-length'].copy().values
    ancho_aux = df['cutting-width'].copy().values
    
    for ptr in df.index:
        if( (df['grain'][ptr] == 'N') or (df['grain'][ptr] == 0)):
            largo_aux[ptr] = df['cutting-width'].copy().values[ptr]
            ancho_aux[ptr] = df['cutting-length'].copy().values[ptr]
            
            
    df['cutting-length'] = largo_aux
    df['cutting-width'] = ancho_aux
    
    # Drop Columns
    cols_to_keep = ['cutting-length','cutting-width','material','label',
                    '_top_','_right_','_bot_','_left_']
    
    
    
    cols_new_names = ['LARGO','ANCHO','TIPO DE PLACA','OBSERVACION',
                      'LARGO SUPERIOR','ANCHO DERECHO','LARGO INFERIOR','ANCHO IZQUIERDO']
    
    df = df[df.columns.intersection(cols_to_keep)]        
    
    
    # Round
    cols = ['cutting-length','cutting-width']
    df[cols] = df[cols].round(0)
    
    # Mantener solo numero de canto

    # df['_top_'] = df['_top_'].apply(lambda x: tipo_canto(x))
    # df['_right_'] = df['_right_'].apply(lambda x: tipo_canto(x))
    # df['_bot_'] = df['_bot_'].apply(lambda x: tipo_canto(x))
    # df['_left_'] = df['_left_'].apply(lambda x: tipo_canto(x))
    # Mantener solo numero de canto
    edge_cols = ['_top_', '_right_', '_bot_', '_left_']
    for col in edge_cols:
        df[col] = df[col].apply(tipo_canto)
    
    # Rename & Reorder
    rename_dict = dict(zip(cols_to_keep, cols_new_names))
    
    df.rename(columns=rename_dict,inplace=True)
    
    cols_order = ['TIPO DE PLACA','LARGO','ANCHO','OBSERVACION',
                  'LARGO SUPERIOR','LARGO INFERIOR','ANCHO DERECHO','ANCHO IZQUIERDO']
    
    df = df[cols_order]
    
    # Insert column CANTIDAD (Value 1 in every row)
    df.insert(1,'CANTIDAD',value=[1]*len(df))
    
    # Insert Column VETA (empty - mocona requirement)
    df.insert(5,'VETA',value=['']*len(df))
    
    return df
    
import traceback  # Add this at the top of your script

def convertFileByFile(filteredFileslist, outDir):
    global logfile
    logfile.write(f'{logdate()} - Procesamiento de archivos individuales iniciado\n') # Added \n
    for file in filteredFileslist:
        try:
            display_filename = file.split(sep='/')[-1]
            dtype_dict = {'cutting-length': float, 'cutting-width': float}
            df = pd.read_excel(os.path.normpath(file), dtype=dtype_dict)
            
            if radio_Empresa.get() == 1: 
                logfile.write(f'{logdate()} - {file} intentando conversión Mocona\n')
                df_converted = convertDfMocona(df)
            elif radio_Empresa.get() == 2: 
                logfile.write(f'{logdate()} - {file} intentando conversión Scheimberg\n')
                df_converted = convertDfScheimberg(df)
            
            # print('____________________________')
            # print(df_converted)
            # print('____________________________')
            df_converted.to_excel(f'{outDir}/out_{display_filename}', index=False)
            
            logfile.write(f'{logdate()} - {file} se ha convertido exitosamente\n')
            logfile.write(f'{logdate()} - El archivo de salida se ha guardado en {outDir}/out_{display_filename} \n')
            log_stringVar.set(f'\U00002705 {logdate()} - Conversión Lista...')
            
        except Exception as e:
            logfile.write(f'{logdate()} - {file} tiene un problema que no permite la conversión\n')
            
            error_traceback = traceback.format_exc()
            logfile.write(f'{logdate()} - Traceback Error:\n{error_traceback}\n')
            # -------------------------------------------
            
            log_stringVar.set(f'\U0000274C {logdate()} - Error: Ver Logs...')
        
        logfile.flush() # Forces python to write to the text file instantly
    

def convertBatch(filteredFileslist, outFile):
    global logfile
    logfile.write(f'{logdate()} - Procesamiento en lotes iniciado\n')
    try:
        df_converted_list = []
        for file in filteredFileslist:
            dtype_dict = {'cutting-length': float, 'cutting-width': float}
            df = pd.read_excel(os.path.normpath(file), dtype=dtype_dict)
            
            
            if radio_Empresa.get() == 1: 
                logfile.write(f'{logdate()} - {file} intentando conversión Mocona\n')
                df_converted = convertDfMocona(df)
            elif radio_Empresa.get() == 2: 
                logfile.write(f'{logdate()} - {file} intentando conversión Scheimberg\n')
                df_converted = convertDfScheimberg(df)            
            
            df_converted_list.append(df_converted)
        
        # print('____________________________')
        # print(df_converted)
        # print('____________________________')

        # Combine everything and save
        df_concat = pd.concat(df_converted_list, axis=0, ignore_index=True)
        df_concat.to_excel(outFile, index=False)

        # Log clean-up based on company type selection
        logfile.write(f'{logdate()} - El procesamiento en lotes se ha completado existosamente\n')
        logfile.write(f'{logdate()} - El archivo de salida se ha guardado en {outFile}\n')
        log_stringVar.set(f'\U00002705 {logdate()} - Conversión Lista...')
        
    except Exception as e:
        logfile.write(f'{logdate()} - Error general en procesamiento por lotes\n')
        
        error_traceback = traceback.format_exc()
        logfile.write(f'{logdate()} - Traceback Error:\n{error_traceback}\n')
        # -------------------------------------------
        
        log_stringVar.set(f'\U0000274C {logdate()} - Error: Ver Logs...')
        
    logfile.flush()

def convert():
    global filesList
    global checkVarList
    
   
    # Filter list of files acconding to checkboxes status
    
    # Bottom Text messege
    if(not filesList):
        log_stringVar.set('\U000026A0 No hay archivos abiertos')
    else:
        filteredFileslist = []
        for ptr in range(len(filesList)):
            if(checkVarList[ptr].get()):
                filteredFileslist.append(filesList[ptr])
        
        if(len(filteredFileslist) == 0):
            log_stringVar.set('\U000026A0 No hay archivos seleccionados')
        else:
            # Convert according to radiobutton status
            if(radio_ConversionType.get() == 1): # Batch
                f = filedialog.asksaveasfile(title='Seleccionar Archivo en Lote',
                                            initialdir=output_dir_StringVar.get(),
                                            filetypes=(("Archivos xlsx","*.xlsx"),),
                                            defaultextension='*.xlsx'
                                            )
                if(f is not None):
                    f.close()
                    outDir = '/'.join(f.name.split(sep='/')[0:-1])
                    output_dir_StringVar.set(outDir)
                    convertBatch(filteredFileslist,f.name)
                
            elif(radio_ConversionType.get() == 2): # Individual
                outDir = filedialog.askdirectory(title='Seleccionar Carpeta para archivos múltiples',
                                                initialdir=output_dir_StringVar.get())
                if(outDir != ''):
                    output_dir_StringVar.set(outDir)
                    convertFileByFile(filteredFileslist,outDir)

def selectionClear():
    global checkVarList
    if(checkVarList):
        [var.set(False) for var in checkVarList]
def selectionAll():
    global checkVarList
    if(checkVarList):
        [var.set(True) for var in checkVarList]

#root = tk.Tk()
root = ttk.Window(themename="cosmo")
root.title('Conversor ABF V2.1.0')
root.geometry('600x700')
# root.iconbitmap("favicon_gc.ico")

# ------- open logfile when launching -------

config_dir = os.path.join(Path.home(), 'Documents', 'conversor_ABF_config')
os.makedirs(config_dir, exist_ok=True)

logfile_path = os.path.join(config_dir, 'conversor_ABF.log')
is_new_log = not os.path.exists(logfile_path) or os.path.getsize(logfile_path) == 0

logfile = open(logfile_path, 'a+')

if is_new_log:
    logfile.write(f'{logdate()} - logfile creado\n')
    logfile.flush() # Asegura que se escriba en el disco de inmediato

config_path = os.path.join(config_dir, 'config')

try:
    with open(config_path, "x") as file:
        print("File created.")
    file.close()
except FileExistsError:
    print("File already exists.")
    pass

with open(config_path, "r") as file:
    cantos_dir = file.readline()

# ------ gui -------

# Vars
filesList = []
checkVarList = []
default_current_dir = os.path.dirname(os.path.abspath(__file__))
#default_output_dir = os.path.join(default_current_dir, 'outputs')

default_output_dir = default_current_dir
title_StringVar = tk.StringVar(value='No Hay Archivos de Entrada Seleccionados')
cantos_dir_StringVar = tk.StringVar(value=f'\U0001F6C8 Archivo de cantos: {cantos_dir}')
output_dir_StringVar = tk.StringVar(value=f'\U0001F6C8 Outputs Path: {default_output_dir}')
current_dir_StringVar = tk.StringVar(value=default_current_dir)
log_stringVar = tk.StringVar(value=f'\U0001F6C8 Información de logs')
radio_ConversionType = tk.IntVar(value=1)
radio_Empresa = tk.IntVar(value=2)



barra_menu = ttk.Menu(root)
menu_archivo = ttk.Menu(barra_menu, tearoff=0)
barra_menu.add_cascade(label="Archivo", menu=menu_archivo)
menu_archivo.add_command(label="Definir archivo de cantos", command=definirCantos)
root.config(menu=barra_menu)

# Define Grid
root.columnconfigure([0], weight=1, uniform='a')
root.columnconfigure([1], weight=2, uniform='a')
root.rowconfigure([0,1,2,3,4,5,6,7,8], weight=1, uniform='a')

# Widgets (All migrated to ttk)
open_btn = ttk.Button(root, text='Abrir Archivos', command=openFiles)
clear_btn = ttk.Button(root, text='Quitar Archivos', command=removeFiles)
convert_btn = ttk.Button(root, text='Convertir', command=convert) # Green button
select_clear_btn = ttk.Button(root, text='Limpiar Seleccion', command=selectionClear, bootstyle="secondary")
select_all_btn = ttk.Button(root, text='Seleccionar Todo', command=selectionAll, bootstyle="secondary")
openlog_btn = ttk.Button(root, text='Abrir Logs', command=openLog, bootstyle="secondary") # Text-link styled button

# Themed Frames for Grouping Radiobuttons
radioButtonConversionTypeFrame = ttk.Frame(root)

radioButton1 = ttk.Radiobutton(radioButtonConversionTypeFrame,
                               text='Convertir en Lote',
                               value=1,
                               variable=radio_ConversionType)
radioButton2 = ttk.Radiobutton(radioButtonConversionTypeFrame,
                               text='Convertir cada archivo',
                               value=2,
                               variable=radio_ConversionType)

radioButtonEmpresaFrame = ttk.Frame(root)

radioButton3 = ttk.Radiobutton(radioButtonEmpresaFrame,
                               text='Mocona',
                               value=1,
                               variable=radio_Empresa)
radioButton4 = ttk.Radiobutton(radioButtonEmpresaFrame,
                               text='Scheimberg',
                               value=2,
                               variable=radio_Empresa)

fileChecklistFrame = ttk.Frame(root)
label_title = ttk.Label(root, textvariable=title_StringVar)

statusLabelsFrame = ttk.Frame(root)
# Bootstyle="info" makes this specific label BLUE
label_odir = ttk.Label(statusLabelsFrame, textvariable=output_dir_StringVar, bootstyle="info")
label_cantosdir = ttk.Label(statusLabelsFrame, textvariable=cantos_dir_StringVar, bootstyle="info")
label_log = ttk.Label(statusLabelsFrame, textvariable=log_stringVar, bootstyle="info")

# Place widgets

open_btn.grid(row=0, column=0, sticky='we', padx=5, pady=5)
clear_btn.grid(row=1, column=0, sticky='we', padx=5, pady=5)
select_clear_btn.grid(row=2, column=0, sticky='we', padx=1, pady=1)
select_all_btn.grid(row=3, column=0, sticky='we', padx=1, pady=1)

radioButtonConversionTypeFrame.grid(row=4, column=0, sticky='we', padx=5, pady=5)
radioButton1.pack(anchor='w', pady=2)
radioButton2.pack(anchor='w', pady=2)

radioButtonEmpresaFrame.grid(row=5, column=0, sticky='we', padx=5, pady=5)
radioButton3.pack(anchor='w', pady=2)
radioButton4.pack(anchor='w', pady=2)

convert_btn.grid(row=6, column=0, sticky='we', padx=5, pady=5)
openlog_btn.grid(row=7, column=0, sticky='we', padx=5, pady=5)

statusLabelsFrame.grid(row=8, column=0, sticky='w', padx=5, pady=5, columnspan=2)
label_odir.pack(anchor='w')
label_cantosdir.pack(anchor='w')
label_log.pack(anchor='w')

label_title.grid(row=0, column=1, sticky='w', padx=10, pady=10)
fileChecklistFrame.grid(row=1, column=1, sticky='nswe', padx=10, pady=10, rowspan=7)

root.mainloop()