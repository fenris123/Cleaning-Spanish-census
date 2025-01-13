# -*- coding: utf-8 -*-
"""
Created on Fri Jan 10 10:09:31 2025

@author: fenris123
"""


### pip install openpyxl
### pip install xlrd
### pip install os
### pip install pandas
import os
import pandas as pd
from openpyxl import load_workbook


# Ruta a la carpeta donde se encuentran los archivos
ruta_carpeta = r"C:\\espaciopython\\padron_siglo_XX"

# A) Convertir los archivos .xls a .xlsx
for archivo in os.listdir(ruta_carpeta):
    if archivo.endswith('.xls'):  # Verificar si es un archivo .xls
        ruta_archivo = os.path.join(ruta_carpeta, archivo)
        
        try:
            # Leer el archivo .xls
            df = pd.read_excel(ruta_archivo, engine='xlrd')
            
            # Crear la ruta para el nuevo archivo .xlsx
            nueva_ruta = os.path.join(ruta_carpeta, archivo.replace('.xls', '.xlsx'))
            
            # Guardar el DataFrame como archivo .xlsx utilizando openpyxl
            df.to_excel(nueva_ruta, index=False, engine='openpyxl')
            print(f'Archivo convertido: {archivo} -> {nueva_ruta}')
            
            # Eliminar el archivo original .xls
            os.remove(ruta_archivo)
        
        except Exception as e:
            print(f'Error al procesar el archivo {archivo}: {e}')





# B) Cambiar el nombre de todos los archivos .xlsx
for archivo in os.listdir(ruta_carpeta):
    if archivo.endswith('.xlsx'):  
        
        # Separar el nombre del archivo y la extensión
        nombre, extension = os.path.splitext(archivo)
        
        # Extraer el año: son los últimos dos caracteres del nombre base
        year = nombre[-2:]  # Los dos últimos dígitos del nombre base
        
        # Crear el nuevo nombre con formato "2000", "2001", etc.
        nuevo_nombre = f"20{year}{extension}" 
        
        # Renombrar el archivo
        os.rename(os.path.join(ruta_carpeta, archivo), os.path.join(ruta_carpeta, nuevo_nombre))
       
# C) limpiar las primeras filas

for archivo in os.listdir(ruta_carpeta):
    if archivo.endswith('.xlsx'):
        ruta_archivo = os.path.join(ruta_carpeta, archivo)
        
        # Cargar el archivo Excel
        wb = load_workbook(ruta_archivo)
        sheet = wb.active  # Usamos la primera hoja activa
        
        # Buscar la primera fila donde la primera columna sea "CPRO" y eliminar filas anteriores
        fila_inicio = 1  # Empezamos desde la primera fila
        while sheet.cell(row=fila_inicio, column=1).value != "CPRO":
            fila_inicio += 1  # Avanzamos fila por fila hasta encontrar "CPRO"
        
        # Eliminar todas las filas antes de la fila donde se encuentra "CPRO"
        if fila_inicio > 1:
            sheet.delete_rows(1, fila_inicio - 1)

        # Obtener el año de la quinta columna, primer fila (los dos últimos caracteres)
        year = str(sheet.cell(row=1, column=5).value)[-2:]  # Accedemos a la quinta columna, primer fila

        # Renombrar las columnas directamente (a mano)
        sheet.cell(row=1, column=1, value="COD_PROV")
        sheet.cell(row=1, column=2, value="PROVINCIA")
        sheet.cell(row=1, column=3, value="COD_MUN")
        sheet.cell(row=1, column=4, value="MUNICIPIO")
        sheet.cell(row=1, column=5, value=f"POB_{year}")
        sheet.cell(row=1, column=6, value=f"HOMBRES_{year}")
        sheet.cell(row=1, column=7, value=f"MUJERES_{year}")

        # Guardar el archivo modificado
        wb.save(ruta_archivo)
        print(f"Archivo procesado: {archivo}")

print("Proceso completado: Archivos renombrados y datos limpios.")



