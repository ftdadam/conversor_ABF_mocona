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

def convertDf(df_og):
    
    df = df_og.copy()
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
                  'LARGO SUPERIOR','LARGO INFERIOR','ANCHO DERECHO','ANCHO IZQUIERDO']
    
    df = df[cols_order]
    
    df.insert(1,'CANTIDAD',value=[1]*len(df))
    df.insert(5,'VETA',value=['']*len(df))
    
    return df
    
def convertFileByFile(file,iDir,oDir):
    try:
        # Read
        dtype_dict = {'cutting-length': float,
                      'cutting-width': float
                      }
        df = pd.read_excel(f'{inputDir}/{file}',dtype=dtype_dict)
        
        df_converted = convertDf(df)
        df_converted.to_excel(f'{oDir}/out_{file}',index=False)
        
        print(f'{file} se ha convertido exitosamente')
        #return df
    except Exception as e:
        print(f'{file} tiene un problema')
        print(type(e).__name__, e)
    

def convertBatch(files,iDir,oDir):
    try:
        df_converted_list = []
        for file in files:
            # Read
            dtype_dict = {'cutting-length': float,
                          'cutting-width': float
                          }
            df = pd.read_excel(f'{inputDir}/{file}',dtype=dtype_dict)
            
            df_converted = convertDf(df)
            
            df_converted_list.append(df_converted)
            
            print(f'{file} se ha convertido exitosamente')
            #return df
        
        print('El procesamiento en lotes se ha completado existosamente')
        df_concat = pd.concat(df_converted_list,axis=0,ignore_index=True)
        df_concat.to_excel(f'{oDir}/out_lote_{files[0]}',index=False)
        
    except Exception as e:
        print(f'{file} tiene un problema')
        print(type(e).__name__, e)
        
        
if(__name__=='__main__'):

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
        print('No hay archivos en la carpeta inputs')
    else:
        print('\n'.join(('Ingrese una opción',
                         '0: Salir',
                         '1: Procesar cada archivo por separado',
                         '2: Procesar todos los archivos como un único lote'))
              )
        
        valid_options = ['0','1','2']
        
        proc_type = input('Ingrese una opción: ')
        while(proc_type not in valid_options):
            proc_type = input('Ingrese una opción válida: ')
            
            
        if(proc_type == '1'):
            print('Procesando cada archivo por separado')
            for file in files:
                print(f'Intentando convertir {file}')
                convertFileByFile(file,inputDir,outputDir)
        elif(proc_type == '2'):
            print('Procesando todos los archivos en lote')
            convertBatch(files,inputDir,outputDir)
        else:
            print('Saliendo')
            


