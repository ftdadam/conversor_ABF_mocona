# -*- coding: utf-8 -*-
"""

@author: fdadam
conversor_ABF_mocona_v.0.1
"""

import os
import pandas as pd
os.chdir(os.path.dirname(os.path.abspath(__file__))) 

def tipo_canto(canto):
    if(isinstance(canto,str)):
        return int(canto.split('-')[0].strip())
    else:
        return ''

def convert_files(file,iDir,oDir):
    try:
        # Read
        dtype_dict = {'cutting-length': float,
                      'cutting-width': float
                      }
        df = pd.read_excel(f'{iDir}/{file}',dtype=dtype_dict)
        
        # Drop
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
                      'LARGO SUPERIOR','ANCHO DERECHO','LARGO INFERIOR','ANCHO IZQUIERDO']
        
        df = df[cols_order]
        
        df.insert(1,'CANTIDAD',value=[1]*len(df))
        df.insert(5,'VETA',value=['']*len(df))
        
        df.to_excel(f'{oDir}/{file}',index=False)
        
        
        
        print(f'{file} se ha convertido exitosamente')
        return df
    except Exception as e:
        print(f'{file} tiene un problema')
        print(type(e).__name__, e)

inputDir = './inputs/'
outputDir = './outputs'

try:
    os.path.isdir(inputDir)
except:
    print('Carpeta Inputs Creada')
    os.mkdir(inputDir)

try:
    os.path.isdir(outputDir)
except:
    print('Carpeta Outputs Creada')
    os.mkdir(outputDir)


files = os.listdir(inputDir)
files = [file for file in files if file.endswith('.xlsx')]

if(not(files)):
    print('No hay archivo en la carpeta inputs')
else:
    for file in files:
        print(f'Intentando convertir {file}')
        df = convert_files(file,inputDir,outputDir)
        


