import pandas as pd
import os

def process_excel_files(folder_path='Processed_Files'):
    """
    Procesa archivos Excel, disminuye el DM Count en 1, y organiza los datos 
    en pestañas (0-5 y 6+) con columnas de CASH OFFER (90% y 60/65%).
    """

    if not os.path.exists(folder_path):
        print(f"Error: La carpeta '{folder_path}' no fue encontrada.")
        return

    for filename in os.listdir(folder_path):
        if filename.endswith(('.xlsx', '.xls', '.xlsm')):
            file_path = os.path.join(folder_path, filename)
            
            try:
                print(f"--- Procesando el archivo: {filename} ---")
                
                excel_file = pd.ExcelFile(file_path)
                sheet_names = excel_file.sheet_names
                original_df = pd.read_excel(excel_file, sheet_name=sheet_names[0])
                
                dm_col = 'MARKETING DM COUNT'
                value_col = 'VALUE'
                
                # 1. Disminuir DM COUNT en 1
                if dm_col in original_df.columns:
                    original_df[dm_col] = original_df[dm_col] - 1
                else:
                    print(f"Error: No se encontró la columna '{dm_col}'.")
                    continue

                # 2. Eliminar columna antigua
                column_to_drop = 'ESTIMATED CASH OFFER'
                if column_to_drop in original_df.columns:
                    original_df = original_df.drop(columns=[column_to_drop])

                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    
                    # --- PROCESAR HOJAS 0 A 5 Y LA NUEVA HOJA 6+ ---
                    # Iteramos del 0 al 6. El 6 representará "6 o más"
                    for count in range(7):
                        if count < 6:
                            sheet_name = f'DM Count {count}'
                            filtered_df = original_df.loc[original_df[dm_col] == count].copy()
                            # Definir % según la regla original
                            second_col_name = 'CASH OFFER 60%' if count <= 2 else 'CASH OFFER 65%'
                            rate = 0.60 if count <= 2 else 0.65
                        else:
                            sheet_name = 'DM Count 6 or more'
                            filtered_df = original_df.loc[original_df[dm_col] >= 6].copy()
                            # Regla para 6 o más: 90% y 65%
                            second_col_name = 'CASH OFFER 65%'
                            rate = 0.65

                        if not filtered_df.empty:
                            # Calcular ambas columnas
                            filtered_df['CASH OFFER 90%'] = filtered_df[value_col] * 0.90
                            filtered_df[second_col_name] = filtered_df[value_col] * rate
                            
                            # Reordenar columnas para que las ofertas queden tras el DM Count
                            cols = filtered_df.columns.tolist()
                            dm_idx = cols.index(dm_col)
                            
                            # Construir el nuevo orden
                            new_cols = (
                                cols[:dm_idx + 1] + 
                                ['CASH OFFER 90%', second_col_name] + 
                                [c for c in cols if c not in (['CASH OFFER 90%', second_col_name] + cols[:dm_idx + 1])]
                            )
                            
                            filtered_df[new_cols].to_excel(writer, sheet_name=sheet_name, index=False)
                            print(f"  -> Hoja '{sheet_name}' creada con columnas 90% y {int(rate*100)}%.")

            except Exception as e:
                print(f"Error procesando '{filename}': {e}")
                
    print("\nProceso completado.")

if __name__ == '__main__':
    process_excel_files()