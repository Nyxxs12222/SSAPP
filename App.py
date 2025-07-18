import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import os

def seleccionar_archivo():
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal
    
    # Configurar la ventana de selección de archivos
    file_path = filedialog.askopenfilename(
        title="Seleccione el archivo CSV",
        filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
    )
    
    return file_path

def procesar_forense(input_csv, output_excel):
    try:
        # Leer el archivo CSV
        df = pd.read_csv(input_csv)
        
        # Verificar columnas requeridas
        columnas_requeridas = ['Telefono', 'Tipo', 'Numero A', 'Numero B', 'Fecha', 'Hora', 
                              'Durac. Seg.', 'IMEI', 'LATITUD', 'LONGITUD', 'Azimuth']
        
        for col in columnas_requeridas:
            if col not in df.columns:
                messagebox.showwarning("Advertencia", f"La columna '{col}' no existe en el archivo CSV")
        
        # Crear el archivo Excel
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            # --- Hoja INTERCALADA ---
            df_intercalada = df.copy()
            
            # Procesamiento numérico
            columnas_numericas = ['Telefono', 'Numero A', 'Numero B', 'IMEI', 'Durac. Seg.', 'Azimuth']
            for col in columnas_numericas:
                if col in df_intercalada.columns:
                    df_intercalada[col] = df_intercalada[col].astype(str).str.replace(r'\D', '', regex=True)
                    df_intercalada[col] = pd.to_numeric(df_intercalada[col], errors='coerce').fillna(0).astype('int64')
            
            # Formatear fecha y hora
            if 'Fecha' in df_intercalada.columns:
                df_intercalada['Fecha'] = pd.to_datetime(df_intercalada['Fecha'], errors='coerce').dt.strftime('%d/%m/%Y')
            
            if 'Hora' in df_intercalada.columns:
                df_intercalada['Hora'] = df_intercalada['Hora'].astype(str).str.replace(r'\D', '', regex=True)
                df_intercalada['Hora'] = df_intercalada['Hora'].str.zfill(6)[:6]
            
            # Filtrar y corregir coordenadas
            if 'LATITUD' in df_intercalada.columns and 'LONGITUD' in df_intercalada.columns:
                df_intercalada = df_intercalada.dropna(subset=['LATITUD', 'LONGITUD'], how='all')
                
                for index, row in df_intercalada.iterrows():
                    try:
                        lat, lon = float(row['LATITUD']), float(row['LONGITUD'])
                        if -180 < lat < 0 and 0 < lon < 90:
                            df_intercalada.at[index, 'LATITUD'], df_intercalada.at[index, 'LONGITUD'] = abs(lon), abs(lat)
                    except (ValueError, TypeError):
                        continue
            
            df_intercalada.to_excel(writer, sheet_name='INTERCALADA', index=False)
            
            # --- Hoja INTEGRADA ---
            df_integrada = df_intercalada.copy()
            df_integrada = df_integrada.drop_duplicates()
            
            def procesar_numero(numero):
                if pd.isna(numero):
                    return numero
                solo_numeros = re.sub(r'\D', '', str(numero))
                return solo_numeros[-10:] if len(solo_numeros) >= 10 else solo_numeros
            
            for col in ['Numero A', 'Numero B']:
                if col in df_integrada.columns:
                    df_integrada[col] = df_integrada[col].apply(procesar_numero)
                    df_integrada[col] = pd.to_numeric(df_integrada[col], errors='coerce')
            
            df_integrada.to_excel(writer, sheet_name='INTEGRADA', index=False)
        
        # Mostrar mensaje de éxito
        messagebox.showinfo(
            "Proceso completado",
            f"Archivo generado con éxito:\n{output_excel}\n\n"
            f"Registros en INTERCALADA: {len(df_intercalada)}\n"
            f"Registros en INTEGRADA: {len(df_integrada)}"
        )
        
        # Abrir el archivo generado automáticamente
        os.startfile(output_excel)
    
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error durante el procesamiento:\n{str(e)}")

def main():
    # Crear ventana principal
    root = tk.Tk()
    root.title("Automatización Forense Digital")
    root.geometry("500x300")
    
    # Estilo
    tk.Label(root, text="Automatización de Procesos Forenses", font=('Arial', 14, 'bold')).pack(pady=20)
    
    # Frame para botones
    frame = tk.Frame(root)
    frame.pack(pady=50)
    
    # Botón para seleccionar archivo
    btn_seleccionar = tk.Button(
        frame,
        text="Seleccionar archivo CSV",
        command=lambda: iniciar_proceso(root),
        height=2,
        width=20
    )
    btn_seleccionar.pack()
    
    root.mainloop()

def iniciar_proceso(root):
    # Seleccionar archivo CSV
    input_csv = seleccionar_archivo()
    
    if not input_csv:
        return  # Usuario canceló
    
    # Crear nombre para el archivo de salida
    output_excel = os.path.splitext(input_csv)[0] + "_procesado.xlsx"
    
    # Mostrar progreso
    progress = tk.Toplevel(root)
    progress.title("Procesando...")
    progress.geometry("300x100")
    tk.Label(progress, text="Procesando archivo, por favor espere...").pack(pady=20)
    progress.update()
    
    # Ejecutar el procesamiento
    procesar_forense(input_csv, output_excel)
    
    # Cerrar ventana de progreso
    progress.destroy()

if __name__ == "__main__":
    main()