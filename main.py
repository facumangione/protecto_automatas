import pandas as pd
import re
from datetime import datetime

# Verifica si una fecha es día no laboral
def es_dia_no_laboral(fecha, feriados):
    if pd.isna(fecha):  # Manejar valores vacíos
        return False
    return fecha.weekday() in [5, 6] or fecha in feriados

# Función para filtrar por rango de fechas
def filtrar_por_rango_fechas(datos, fecha_inicio, fecha_fin):
    return datos[(datos['Inicio_de_Conexión_Dia'] >= fecha_inicio) & (datos['Inicio_de_Conexión_Dia'] <= fecha_fin)]

def procesar_datos_eficientemente(ruta_archivo, feriados, fecha_inicio=None, fecha_fin=None, bloque_size=50000):
    try:
        print("Leyendo el archivo completo...")
        columnas_necesarias = ['Inicio_de_Conexión_Dia', 'Usuario']

        # Leer columnas necesarias y procesar en bloques
        datos = pd.read_excel(ruta_archivo, usecols=columnas_necesarias)
        datos['Inicio_de_Conexión_Dia'] = pd.to_datetime(datos['Inicio_de_Conexión_Dia'], errors='coerce')

        # Filtrar por rango de fechas si es necesario
        if fecha_inicio and fecha_fin:
            datos = filtrar_por_rango_fechas(datos, fecha_inicio, fecha_fin)

        print("Aplicando validaciones y filtrando datos...")
        resultados = []

        for inicio in range(0, len(datos), bloque_size):
            fin = inicio + bloque_size
            bloque = datos.iloc[inicio:fin]

            # Filtrar días no laborales (sábados, domingos, feriados)
            filtro = bloque['Inicio_de_Conexión_Dia'].apply(lambda x: es_dia_no_laboral(x, feriados))
            resultados.append(bloque[filtro])

            print(f"Procesado bloque {inicio} a {fin}...")

        # Unir los resultados procesados
        datos_filtrados = pd.concat(resultados)

        # Exportar a Excel
        archivo_salida = 'conexiones_feriados_y_fines_de_semana.xlsx'
        datos_filtrados.to_excel(archivo_salida, index=False)
        print(f"Resultados exportados a {archivo_salida}")

    except Exception as e:
        print(f"Error al procesar los datos: {e}")

def main():
    # Ruta del archivo
    ruta_archivo = 'bd_automatas.xlsx'

    # Lista de feriados (ejemplo)
    feriados = pd.to_datetime([
        '2019-01-01', '2019-02-24', '2019-03-29', '2019-05-01', 
        '2019-06-20', '2019-07-09', '2019-08-17', '2019-10-12',
        '2019-11-02', '2019-12-25'
    ])

    # Solicitar rango de fechas al usuario
    fecha_inicio_str = input("Ingresa la fecha de inicio (YYYY-MM-DD): ")
    fecha_fin_str = input("Ingresa la fecha de fin (YYYY-MM-DD): ")

    # Convertir las fechas de string a datetime
    fecha_inicio = pd.to_datetime(fecha_inicio_str)
    fecha_fin = pd.to_datetime(fecha_fin_str)

    # Procesar los datos con el rango de fechas
    procesar_datos_eficientemente(ruta_archivo, feriados, fecha_inicio, fecha_fin)

if __name__ == "__main__":
    main()
