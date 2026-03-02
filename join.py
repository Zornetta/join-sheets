import pandas as pd
import sys
import os
import openpyxl

def prompt_file(prompt_num):
    # Intentar usar tkinter primero
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.attributes('-topmost', True) # Traer al frente
        root.withdraw() # Ocultar la ventana principal
        filename = filedialog.askopenfilename(
            title=f"Seleccionar Hoja {prompt_num} (Excel)",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        root.destroy()
        if filename:
            return filename
    except Exception:
        pass
        
    # Alternativa con zenity si tkinter falla (estándar en Ubuntu/GNOME)
    try:
        import subprocess
        result = subprocess.run(['zenity', '--file-selection', f'--title=Seleccionar Hoja {prompt_num} (Excel)'],
                                capture_output=True, text=True)
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        pass
        
    # Última alternativa por si acaso
    return input(f"Ingrese el nombre de archivo para la Hoja {prompt_num} (ej., 'hoja {prompt_num}.xlsx'): ").strip()

def get_file_and_column(prompt_num):
    print(f"\n--- Detalles de la Hoja {prompt_num} ---")
    print(f"Por favor, seleccione el archivo para la Hoja {prompt_num} desde la ventana de diálogo...")
    
    filename = prompt_file(prompt_num)
    
    if not filename or not os.path.exists(filename):
        print("No se seleccionó un archivo válido. Saliendo.")
        sys.exit(1)
        
    print(f" -> Seleccionado: {filename}")
    
    try:
        print(f"Leyendo '{filename}'...")
        df = pd.read_excel(filename, dtype=str)
        print(f" -> Se cargó con éxito '{filename}' con {len(df)} filas.")
    except Exception as e:
        print(f"ERROR al leer el archivo '{filename}':", e)
        sys.exit(1)
        
    print(f"\nColumnas disponibles para unir en la Hoja {prompt_num}:")
    columns = df.columns.tolist()
    for i, col in enumerate(columns, 1):
        print(f"  [{i}] {col}")
        
    while True:
        choice = input(f"Ingrese el número [1-{len(columns)}] correspondiente a la columna por la cual unir: ").strip()
        try:
            choice_idx = int(choice) - 1
            if 0 <= choice_idx < len(columns):
                col_name = columns[choice_idx]
                print(f" -> Columna seleccionada: '{col_name}'")
                break
            else:
                print(f"Número inválido. Por favor, ingrese un número entre 1 y {len(columns)}.")
        except ValueError:
            print("Por favor, ingrese un número entero válido.")
        
    return df, col_name, filename

def main():
    print("--------------------------------------------------")
    print("Iniciando el proceso interactivo de unión...")
    
    df1, join_col1, file1 = get_file_and_column(1)
    df2, join_col2, file2 = get_file_and_column(2)
    
    print(f"\n -> Usando '{join_col1}' de {file1} y '{join_col2}' de {file2}.")

    print("\nUniendo las hojas...")
    df_merged = pd.merge(df1, df2, left_on=join_col1, right_on=join_col2, how='outer', indicator=True)
    
    # Combinar las columnas de unión en una sola usando el nombre de la columna de la Hoja 1
    if join_col1 != join_col2:
        df_merged[join_col1] = df_merged[join_col1].fillna(df_merged[join_col2])
        df_merged = df_merged.drop(columns=[join_col2])
        
    # Ordenar las IPs no coincidentes hacia el final
    match_mapping = {'both': 0, 'left_only': 1, 'right_only': 2}
    df_merged['match_priority'] = df_merged['_merge'].map(match_mapping)
    df_merged = df_merged.sort_values('match_priority').drop(columns=['match_priority'])
    
    # Agregar una columna de Estado de Coincidencia
    df_merged['_merge'] = df_merged['_merge'].map({
        'both': 'Coincidencia', 
        'left_only': 'Solo en Hoja 1', 
        'right_only': 'Solo en Hoja 2'
    })
    df_merged = df_merged.rename(columns={'_merge': 'Estado de Coincidencia'})
    
    # Analizar el estado de la coincidencia
    matches = len(df_merged[df_merged['Estado de Coincidencia'] == 'Coincidencia'])
    only_1 = len(df_merged[df_merged['Estado de Coincidencia'] == 'Solo en Hoja 1'])
    only_2 = len(df_merged[df_merged['Estado de Coincidencia'] == 'Solo en Hoja 2'])
    
    print(f" -> Unión exitosa. Filas totales: {len(df_merged)}")
    print("ESTADÍSTICAS DEL ESTADO DE COINCIDENCIA:")
    print(f"   - Coincidentes en ambas hojas: {matches}")
    print(f"   - Presentes solo en la hoja 1: {only_1}")
    print(f"   - Presentes solo en la hoja 2: {only_2}")
    
    print("\nEliminando la columna 'Estado de Coincidencia' del resultado final...")
    df_merged = df_merged.drop(columns=['Estado de Coincidencia'])
    
    print("Guardando en 'hojas_unidas.xlsx'...")
    # Rellenar valores NaN con espacios en blanco para evitar que aparezcan como texto "NaN"
    df_merged = df_merged.fillna("")
    # Guardar el dataframe
    output_file = 'hojas_unidas.xlsx'
    df_merged.to_excel(output_file, index=False)
    
    print("Limpiando los encabezados (eliminando 'Unnamed: X' generado por encabezados vacíos)...")
    # Cargar el libro de trabajo guardado para corregir encabezados de forma programática
    wb = openpyxl.load_workbook(output_file)
    ws = wb.active
    
    for cell in ws[1]:
        if cell.value and isinstance(cell.value, str) and cell.value.startswith("Unnamed:"):
            cell.value = ""  # Limpiar el encabezado
            
    wb.save(output_file)
    print("¡Listo! Los datos unidos se guardaron exitosamente en 'hojas_unidas.xlsx'.")
    print("--------------------------------------------------")

if __name__ == '__main__':
    main()
