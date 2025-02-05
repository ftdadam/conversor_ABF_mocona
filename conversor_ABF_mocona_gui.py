# -*- coding: utf-8 -*-
"""

@author: fdadam
"""

import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog
import pandas as pd
from datetime import datetime
from pathlib import Path
import time
import numpy as np




def logdate():
    return datetime.now().strftime("[%Y/%m/%d, %H:%M:%S]")

def openFiles():
    global filesList
    global checkVarList

    filesFromDiag = filedialog.askopenfilenames(initialdir=current_dir_StringVar,title='Seleccionar Input Files',
                                        filetypes=(("Archivos xlsx","*.xlsx"),))
    filesFromDiag = list(filesFromDiag)

    checkVarFromDiag = []
    for _ in range(len(filesFromDiag)):
        print(_)
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
            Checkbutton(fileChecklistFrame,
                        text=displayFilenamesList[ptr],
                        variable=checkVarList[ptr]).pack(anchor='w')

        print([var.get() for var in checkVarList])

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
        return ''

def convertDf(df_og):
    
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

    df['_top_'] = df['_top_'].apply(lambda x: tipo_canto(x))
    df['_right_'] = df['_right_'].apply(lambda x: tipo_canto(x))
    df['_bot_'] = df['_bot_'].apply(lambda x: tipo_canto(x))
    df['_left_'] = df['_left_'].apply(lambda x: tipo_canto(x))
    
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
    
def convertFileByFile(filteredFileslist,outDir):
    global logfile
    logfile.write(f'{logdate()} - Procesamiento de archivos individuales iniciado')
    for file in filteredFileslist:
        try:
            display_filename = file.split(sep='/')[-1]
            # Read
            dtype_dict = {'cutting-length': float,
                          'cutting-width': float
                          }
            df = pd.read_excel(os.path.normpath(file),dtype=dtype_dict)
            
            df_converted = convertDf(df)
            df_converted.to_excel(f'{outDir}/out_{display_filename}',index=False)
            
            logfile.write(f'{logdate()} - {file} se ha convertido exitosamente\n')
            logfile.write(f'{logdate()} - El archivo de salida se ha guardado en {outDir}/out_{display_filename} \n')
            log_stringVar.set(f'{logdate()} - Conversión Lista...')
        except Exception as e:
            logfile.write(f'{logdate()} - {file} tiene un problema que no permite la conversión\n')
            logfile.write(f'{logdate()} - Debug Error: {type(e).__name__, e}\n')
            log_stringVar.set(f'{logdate()} - Error: Ver Logs...')
    

def convertBatch(filteredFileslist,outFile):
    global logfile
    logfile.write(f'{logdate()} - Procesamiento en lotes iniciado')
    try:
        df_converted_list = []
        for file in filteredFileslist:
            
            display_filename = file.split(sep='/')[-1]
            # Read
            dtype_dict = {'cutting-length': float,
                          'cutting-width': float
                          }
            df = pd.read_excel(os.path.normpath(file),dtype=dtype_dict)
            
            df_converted = convertDf(df)
            
            df_converted_list.append(df_converted)
            
            logfile.write(f'{logdate()} - {file} se ha convertido exitosamente\n')
        
        df_concat = pd.concat(df_converted_list,axis=0,ignore_index=True)
        df_concat.to_excel(outFile,index=False)

        logfile.write(f'{logdate()} - El procesamiento en lotes se ha completado existosamente\n')
        logfile.write(f'{logdate()} - El archivo de salida se ha guardado en {outFile}\n')
        log_stringVar.set(f'{logdate()} - Conversión Lista...')
    except Exception as e:
        logfile.write(f'{logdate()} - {file} tiene un problema que no permite la conversión\n')
        logfile.write(f'{logdate()} - Debug Error: {type(e).__name__, e}\n')
        log_stringVar.set(f'{logdate()} - Error: Ver Logs...')

def convert():
    global filesList
    global checkVarList
    
    # Filter list of files acconding to checkboxes status
    filteredFileslist = []
    for ptr in range(len(filesList)):
        if(checkVarList[ptr].get()):
            filteredFileslist.append(filesList[ptr])

    # Bottom Text messege
    if(not filesList):
        log_stringVar.set('No hay archivos abiertos')
    if(not filesList):
        log_stringVar.set('No hay archivos seleccionados')
    
    
    if(filteredFileslist):
        # Convert according to radiobutton status
        if(radio_IntVar.get() == 1):
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
            
        elif(radio_IntVar.get() == 2):
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

root = tk.Tk()
root.title('Conversor ABF Mocona')
root.geometry('600x450')
#root.iconbitmap("favicon_gc.ico")

# ------- open logfile when launching -------

logfile_path = os.path.join(Path.home(),'Documents','conversor_ABF_mocona.log')
logfile = open(logfile_path,'a+')

# ------- check if logfile is new, make first entry -------
if(os.path.getsize(logfile_path) == 0):
    logfile.write(f'{logdate()} - logfile creado\n')

# ------ gui -------

# Vars
filesList = []
checkVarList = []
default_output_dir = os.path.dirname(os.path.abspath(__file__))
default_current_dir = os.path.dirname(os.path.abspath(__file__))
title_StringVar = tk.StringVar(value='No Hay Archivos de Entrada Seleccionados')
output_dir_StringVar = tk.StringVar(value=f'Output: {default_output_dir}')
current_dir_StringVar = tk.StringVar(value=default_current_dir)
log_stringVar = tk.StringVar(value='')
radio_IntVar = tk.IntVar(value='1')

# Define Grid
root.columnconfigure([0], weight= 1, uniform = 'a')
root.columnconfigure([1], weight= 2, uniform = 'a')
root.rowconfigure([0,1,2,3,4,5,6,7],weight = 1,uniform='a')

# Widgets
open_btn = Button(root,text='Abrir Archivos',command=openFiles)
clear_btn = Button(root,text='Quitar Archivos',command=removeFiles)
convert_btn = Button(root,text='Convertir',command=convert)
select_clear_btn = Button(root,text='Limpiar Seleccion',command=selectionClear)
select_all_btn = Button(root,text='Seleccionar Todo',command=selectionAll)
openlog_btn = Button(root,text='Abrir Logs',command=openLog)

radioButtonFrame = tk.Frame(root)

radioButton1 = Radiobutton(radioButtonFrame,
                               text='Convertir en Lote',
                               value= 1,
                               variable=radio_IntVar,
                               )
radioButton2 = Radiobutton(radioButtonFrame,
                               text='Convertir cada archivo',
                               value= 2,
                               variable=radio_IntVar)

fileChecklistFrame = tk.Frame(root)
label_title = Label(root,textvariable=title_StringVar)

statusLabelsFrame=tk.Frame(root)
label_odir = Label(statusLabelsFrame,textvariable=output_dir_StringVar)
label_log = Label(statusLabelsFrame,textvariable=log_stringVar)

# Place widgets

open_btn.grid(row = 0, column = 0, sticky = 'we', padx=5, pady=5)
clear_btn.grid(row = 1, column = 0, sticky = 'we', padx=5, pady=5)
select_clear_btn.grid(row=2,column = 0, sticky = 'we',padx=1, pady=1)
select_all_btn.grid(row=3,column = 0, sticky = 'we',padx=1, pady=1)
radioButtonFrame.grid(row=4, column = 0, sticky = 'we',padx=5, pady=5)
radioButton1.pack(anchor='w')
radioButton2.pack(anchor='w')
convert_btn.grid(row=5, column = 0,sticky = 'we', padx=5, pady=5)
openlog_btn.grid(row=6, column = 0,sticky = 'we', padx=5, pady=5)
statusLabelsFrame.grid(row=7, column=0,sticky='w', padx=5, pady=5,columnspan=2)
label_odir.pack(anchor='w')
label_log.pack(anchor='w')

label_title.grid(row=0, column = 1, sticky = 'w',padx=10, pady=10)
fileChecklistFrame.grid(row=1,column=1,sticky='nswe',padx=10,pady=10,rowspan=4)

root.mainloop()