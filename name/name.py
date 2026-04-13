import os
import pandas as pd
import re

def actualizar_conteo_filas(ruta_folder):
    folder_path = os.path.join(ruta_folder, "Processed_Files")
    
    if not os.path.exists(folder_path):
        print(f"La carpeta '{folder_path}' no existe.")
        return

    archivos = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls', '.csv'))]

    for nombre_archivo in archivos:
        ruta_completa = os.path.join(folder_path, nombre_archivo)
        
        try:
            # 1. Leer el archivo y contar filas
            # Usamos pandas para obtener el número de filas (excluyendo el encabezado)
            if nombre_archivo.endswith('.csv'):
                df = pd.read_csv(ruta_completa)
            else:
                df = pd.read_excel(ruta_completa)
            
            conteo_real = len(df)
            
            # Convertir el conteo a formato "K" (ejemplo: 5300 -> 5.3K o 5K si es exacto)
            if conteo_real >= 1000:
                valor_k = conteo_real / 1000
                # Si el decimal es .0, lo quitamos (ej: 5.0K -> 5K)
                conteo_str = f"{valor_k:g}K"
            else:
                conteo_str = str(conteo_real)

            # 2. Identificar el patrón de cantidad actual (ej: 6K, 3.5K, 41K)
            # Buscamos un número seguido de una 'K' (mayúscula o minúscula)
            patron = r'(\d+\.?\d*)[kK]'
            
            # 3. Reemplazar el valor viejo por el nuevo
            if re.search(patron, nombre_archivo):
                nuevo_nombre = re.sub(patron, conteo_str, nombre_archivo)
            else:
                # Si por alguna razón no tiene "K", lo insertamos antes de la extensión
                nombre_sin_ext, ext = os.path.splitext(nombre_archivo)
                nuevo_nombre = f"{nombre_sin_ext} {conteo_str}{ext}"

            # 4. Renombrar el archivo físico
            nueva_ruta = os.path.join(folder_path, nuevo_nombre)
            
            if ruta_completa != nueva_ruta:
                # Si el archivo de destino ya existe, lo eliminamos para evitar errores
                if os.path.exists(nueva_ruta):
                    os.remove(nueva_ruta)
                os.rename(ruta_completa, nueva_ruta)
                print(f"Actualizado: {nombre_archivo} -> {nuevo_nombre} ({conteo_real} filas)")
            else:
                print(f"Sin cambios: {nombre_archivo} (Sigue siendo {conteo_str})")

        except Exception as e:
            print(f"Error procesando {nombre_archivo}: {e}")

if __name__ == "__main__":
    # Ejecuta en la carpeta donde está el script
    actualizar_conteo_filas(os.getcwd())