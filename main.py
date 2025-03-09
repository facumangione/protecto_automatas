import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Función para validar el formato de fecha con expresiones regulares
def validar_fecha(fecha):
    patron = r"^\d{4}-\d{2}-\d{2}$"
    return re.match(patron, fecha)

# Función para procesar el archivo y mostrar la vista previa
def procesar_archivo():
    global df_filtrado  # Para usarlo después en la exportación

    # Pedir archivo al usuario
    archivo_excel = filedialog.askopenfilename(title="Selecciona el archivo Excel", filetypes=[("Archivos Excel", "*.xlsx")])
    
    if not archivo_excel:
        messagebox.showerror("Error", "No seleccionaste ningún archivo.")
        return
    
    df = pd.read_excel(archivo_excel)

    # Convertir las fechas a datetime
    df['Inicio_de_Conexión_Dia'] = pd.to_datetime(df['Inicio_de_Conexión_Dia'], errors='coerce')

    # Definir feriados de 2019
    feriados_2019 = [
        "2019-01-01", "2019-03-04", "2019-03-05", "2019-03-24",
        "2019-04-02", "2019-04-19", "2019-05-01", "2019-05-25",
        "2019-06-17", "2019-06-20", "2019-07-09", "2019-08-17",
        "2019-10-12", "2019-11-20", "2019-12-08", "2019-12-25"
    ]
    feriados_2019 = pd.to_datetime(feriados_2019)

    # Obtener fechas ingresadas
    fecha_inicio = entrada_inicio.get()
    fecha_fin = entrada_fin.get()

    # Validar formato de fecha
    if not validar_fecha(fecha_inicio) or not validar_fecha(fecha_fin):
        messagebox.showerror("Error", "Formato de fecha incorrecto. Usa YYYY-MM-DD.")
        return

    fecha_inicio = pd.to_datetime(fecha_inicio)
    fecha_fin = pd.to_datetime(fecha_fin)

    # Filtrar datos
    df_filtrado = df[
        ((df['Inicio_de_Conexión_Dia'].dt.dayofweek >= 5) |  # Sábado (5) y domingo (6)
         (df['Inicio_de_Conexión_Dia'].isin(feriados_2019))) &
        (df['Inicio_de_Conexión_Dia'].between(fecha_inicio, fecha_fin))
    ]

    # Mostrar vista previa
    actualizar_vista_previa(df_filtrado)

# Función para mostrar la tabla con los datos filtrados
def actualizar_vista_previa(df):
    # Limpiar tabla anterior
    for fila in tabla.get_children():
        tabla.delete(fila)

    # Agregar filas a la tabla
    for _, fila in df.iterrows():
        valores = [fila[col] for col in df.columns]
        tabla.insert("", "end", values=valores)

    # Hacer visible el botón de exportar
    boton_exportar["state"] = "normal"

# Función para exportar los datos filtrados a Excel
def exportar_a_excel():
    df_filtrado.to_excel("conexiones_filtradas.xlsx", index=False)
    messagebox.showinfo("Éxito", "Archivo generado: conexiones_filtradas.xlsx")

# Crear la interfaz gráfica
ventana = tk.Tk()
ventana.title("Análisis de Conexiones WiFi")

tk.Label(ventana, text="Fecha de inicio (YYYY-MM-DD):").pack()
entrada_inicio = tk.Entry(ventana)
entrada_inicio.pack()

tk.Label(ventana, text="Fecha de fin (YYYY-MM-DD):").pack()
entrada_fin = tk.Entry(ventana)
entrada_fin.pack()

boton_procesar = tk.Button(ventana, text="Seleccionar Archivo y Procesar", command=procesar_archivo)
boton_procesar.pack()

# Tabla de vista previa
columnas = ["ID", "ID_Sesion", "ID_Conexión_unico", "Usuario", "IP_NAS_AP", "Tipo__conexión", 
            "Inicio_de_Conexión_Dia", "Inicio_de_Conexión_Hora", "FIN_de_Conexión_Dia", 
            "FIN_de_Conexión_Hora", "Session_Time", "Input_Octects", "Output_Octects", 
            "MAC_AP", "MAC_Cliente", "Razon_de_Terminación_de_Sesión"]

tabla = ttk.Treeview(ventana, columns=columnas, show="headings", height=10)
for col in columnas:
    tabla.heading(col, text=col)
    tabla.column(col, width=100)

tabla.pack()

# Botón de exportación (inicialmente deshabilitado)
boton_exportar = tk.Button(ventana, text="Exportar a Excel", command=exportar_a_excel, state="disabled")
boton_exportar.pack()

ventana.mainloop()