import pandas as pd
import os
from openpyxl import load_workbook
from copy import copy

# Configuración
archivo_lista = 'archivo-lista.xlsx'
archivo_plantilla = 'archivo-plantilla.xlsx'
carpeta_salida = 'archivos_generados'

# Crear carpeta de salida
os.makedirs(carpeta_salida, exist_ok=True)

# Leer lista de datos
df = pd.read_excel(archivo_lista)

# Cargar plantilla
wb_plantilla = load_workbook(archivo_plantilla)
hoja_plantilla = wb_plantilla.active

for index, row in df.iterrows():
    nombre = row['nombre']      # ajusta a tus nombres de columna
    apellido = row['apellido']  # ajusta a tus nombres de columna
    serie = row['serie']        # ajusta a tus nombres de columna
    
    # Crear copia de la plantilla
    nuevo_wb = load_workbook(archivo_plantilla)
    nueva_hoja = nuevo_wb.active
    
    # Modificar plantilla (ajusta las celdas según necesites)
    nueva_hoja['A1'] = nombre
    nueva_hoja['B1'] = apellido
    nueva_hoja['C1'] = serie
    
    # Guardar archivo
    nombre_archivo = f"{carpeta_salida}/{nombre}_{apellido}_{serie}.xlsx"
    nuevo_wb.save(nombre_archivo)
    print(f"Archivo generado: {nombre_archivo}")

print("Proceso completado!")
