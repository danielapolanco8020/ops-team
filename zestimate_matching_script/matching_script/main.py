"""
SCRIPT DE COMBINACIÓN DE DATOS

Este script combina un archivo de Excel con los resultados del scraping de Zillow.

PASOS PARA USAR:

1. Instalar las librerías necesarias:
   - Abrir la terminal en esta carpeta
   - Ejecutar: pip install -r requirements.txt

2. Guardar tu archivo de Excel en esta misma carpeta

3. Editar las variables abajo (JOB_ID y NOMBRE_INPUT):
   - JOB_ID: El identificador del job que quieres descargar
   - NOMBRE_INPUT: El nombre exacto de tu archivo Excel (con extensión .xlsx)

4. Ejecutar el script:
   - En la terminal, ejecutar: python main.py

El resultado se guardará como un archivo CSV con el nombre "RESULTADO - [tu archivo].csv"
"""

# ==============================================================================
# VARIABLES - EDITAR ESTAS VARIABLES
# ==============================================================================
JOB_ID = '94ac44c813df374a3e3f'
NOMBRE_INPUT = "2026-03-04 CASAONERESIDENTIAL 80K Direct Mail.xlsx"



# ==============================================================================
# PROCESAMIENTO - NO EDITAR
# ==============================================================================
output_files = []
has_signed_urls = False


# ------------------------------------------------------------------------------
# Descarga del output del WSE
# ------------------------------------------------------------------------------
from dotenv import load_dotenv
import pandas as pd
import requests
import json
import sys
import os

load_dotenv()
os.makedirs('data', exist_ok=True)

headers = {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + os.getenv('WSE_API_KEY')
}

params = {
    'id': JOB_ID,
    'output_type': 'csv'
}

response = requests.get('https://wse.icebergdata.io/public-api/downloadOutput', headers=headers, params=params)

if response.status_code == 200:
    response_headers = response.headers
    if response_headers['Content-Type'] == 'text/csv':
        with open(f'data/{JOB_ID}.csv', 'wb') as f:
            f.write(response.content)
        output_files.append(f'data/{JOB_ID}.csv')
    else:
        with open(f'data/{JOB_ID}_signed_urls.json', 'wb') as f:
            f.write(response.content)
        has_signed_urls = True
elif response.status_code == 202:
    print("El archivo no está listo. Por favor, espere a que se complete el job.")
    sys.exit(0)
else:
    print("Error al descargar el archivo.")
    sys.exit(0)


# ------------------------------------------------------------------------------
# Unificación de los outputs
# ------------------------------------------------------------------------------
if has_signed_urls:
    with open(f'data/{JOB_ID}_signed_urls.json', 'r') as f:
        signed_urls = json.load(f)
    for i in range(len(signed_urls)):
        if not os.path.exists(f'data/{JOB_ID}-{i}.csv'):
            response = requests.get(signed_urls[i])
            with open(f'data/{JOB_ID}-{i}.csv', 'wb') as f:
                f.write(response.content)
        output_files.append(f'data/{JOB_ID}-{i}.csv')

dfs = [pd.read_csv(file) for file in output_files]
output_df = pd.concat(dfs)
output_df.to_csv(f'data/{JOB_ID}.csv', index=False)


# ------------------------------------------------------------------------------
# Combinación con el input
# ------------------------------------------------------------------------------
input_df = pd.read_excel(f"./{NOMBRE_INPUT}")
joined_df = pd.merge(input_df, output_df, on=['FOLIO', 'ADDRESS', 'CITY', 'STATE'], how='left')
joined_df.drop(columns=['ZIP_y'], inplace=True)
joined_df.rename(columns={'ZIP_x': 'ZIP'}, inplace=True)
joined_df.to_csv("RESULTADO - " + NOMBRE_INPUT.replace(".xlsx", ".csv"), index=False)
print("RESULTADO - " + NOMBRE_INPUT.replace(".xlsx", ".csv") + " creado. Numero de registros: " + str(len(joined_df)))