import polars as pl
import tkinter as tk
from tkinter import messagebox
import pandas as pd
import re
import os
import threading
from datetime import datetime

def cargar_datos(archivo):
    try:
        # Usar pandas para leer el Excel
        df_pd = pd.read_excel(archivo)

        # Verificar las columnas y sus tipos en pandas
        print(f"Columnas y tipos de datos en pandas DataFrame:\n{df_pd.dtypes}")

        # Convertir las columnas de fecha a formato datetime (si existen)
        if "Inicio_de_Conexión_Dia" in df_pd.columns:
            df_pd["Inicio_de_Conexión_Dia"] = pd.to_datetime(df_pd["Inicio_de_Conexión_Dia"], errors='coerce')

        if "FIN_de_Conexión_Dia" in df_pd.columns:
            df_pd["FIN_de_Conexión_Dia"] = pd.to_datetime(df_pd["FIN_de_Conexión_Dia"], errors='coerce')

        # Convertir las columnas numéricas a tipo adecuado (int o float)
        num_columns = ["Session_Time", "Input_Octects", "Output_Octects"]
        for col in num_columns:
            if col in df_pd.columns:
                df_pd[col] = pd.to_numeric(df_pd[col], errors='coerce')

        # Eliminar columnas no necesarias, como 'Unnamed' o cualquier otra columna vacía o irrelevante
        df_pd = df_pd.dropna(axis=1, how='all')  # Eliminar columnas completamente vacías
        df_pd = df_pd.loc[:, ~df_pd.columns.str.contains('^Unnamed')]  # Eliminar columnas con nombres 'Unnamed'

        # Convertir todas las columnas de texto a string para compatibilidad con Polars
        for col in df_pd.select_dtypes(include='object').columns:
            df_pd[col] = df_pd[col].astype(str)

        # Verifica nuevamente las columnas y sus tipos después de las conversiones
        print(f"Columnas y tipos después de la conversión:\n{df_pd.dtypes}")

        # Convertir el DataFrame de pandas a Polars
        df = pl.from_pandas(df_pd)

        print("Datos cargados correctamente.")
        return df
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        return None
    
def validar_fecha(fecha):
    return bool(re.match(r"^\d{4}-\d{2}-\d{2}$", fecha))

def convertir_fecha(fecha_str):
    try:
        return datetime.strptime(fecha_str, "%Y-%m-%d").date()
    except ValueError:
        return None

def filtrar_por_fecha(df, fecha_inicio, fecha_fin):
    if not (validar_fecha(fecha_inicio) and validar_fecha(fecha_fin)):
        messagebox.showerror("Error", "Formato de fecha incorrecto. Usa YYYY-MM-DD.")
        return df
    
    fecha_inicio = convertir_fecha(fecha_inicio)
    fecha_fin = convertir_fecha(fecha_fin)
    return df.filter((pl.col("Inicio_de_Conexión_Dia") >= fecha_inicio) & (pl.col("Inicio_de_Conexión_Dia") <= fecha_fin))

def detectar_feriados_y_no_laborables(df, feriados):
    feriados_pl = [convertir_fecha(f) for f in feriados if convertir_fecha(f)]
    df = df.with_columns(
        pl.col("Inicio_de_Conexión_Dia").is_in(feriados_pl).alias("Es_feriado"),
        pl.col("Inicio_de_Conexión_Dia").map_elements(lambda d: d.weekday() >= 5 if d is not None else False).alias("Es_no_laborable")
    )
    return df

def exportar_excel(df_final):
    def guardar():
        try:
            carpeta_salida = os.path.dirname(os.path.abspath(__file__))
            archivo_salida = os.path.join(carpeta_salida, "conexiones_feriados.xlsx")
            df_final.to_pandas().to_excel(archivo_salida, index=False)
            messagebox.showinfo("Éxito", f"Datos exportados correctamente en: {archivo_salida}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar el archivo: {e}")
    
    hilo = threading.Thread(target=guardar)
    hilo.start()

def crear_interfaz(df):
    root = tk.Tk()
    root.title("Analizador de Conexiones WiFi")
    root.geometry("500x400")
    
    tk.Label(root, text="Fecha inicio (YYYY-MM-DD):").pack(pady=5)
    fecha_inicio_entry = tk.Entry(root)
    fecha_inicio_entry.pack()
    
    tk.Label(root, text="Fecha fin (YYYY-MM-DD):").pack(pady=5)
    fecha_fin_entry = tk.Entry(root)
    fecha_fin_entry.pack()
    
    resultado_lista = tk.Listbox(root, width=80, height=10)
    resultado_lista.pack(pady=10)
    
    def procesar():
        fecha_inicio = fecha_inicio_entry.get()
        fecha_fin = fecha_fin_entry.get()
        
        df_filtrado = filtrar_por_fecha(df, fecha_inicio, fecha_fin)
        df_final = detectar_feriados_y_no_laborables(df_filtrado, feriados)
        
        resultado_lista.delete(0, tk.END)
        for row in df_final.iter_rows(named=True):
            resultado_lista.insert(tk.END, f"{row['Inicio_de_Conexión_Dia']} - Feriado: {row['Es_feriado']} - No Laborable: {row['Es_no_laborable']}")
        
        messagebox.showinfo("Procesado", f"Se encontraron {len(df_final)} registros.")
        
        
        exportar_excel(df_final)
    
    tk.Button(root, text="Analizar", command=procesar).pack(pady=5)
    tk.Button(root, text="Exportar a Excel", command=lambda: exportar_excel(df)).pack(pady=5)
    root.mainloop()

archivo = r"C:\Users\Facundo\Desktop\Facundo Tareas\protecto_automatas\protecto_automatas\bd_automatas.xlsx"

feriados = ["2024-01-01", "2024-05-01", "2024-12-25"]

df = cargar_datos(archivo)
if df is not None:
    crear_interfaz(df)

