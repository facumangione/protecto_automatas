import pandas as pd
import re
from datetime import datetime

# Función para validar formatos con expresiones regulares
def validar_formato(cadena, patron):
    return bool(re.match(patron, str(cadena)))

# Verifica si una fecha es día no laboral
def es_dia_no_laboral(fecha, feriados):
    if pd.isna(fecha):  # Manejar valores vacíos
        return False
    return fecha.weekday() in [5, 6] or fecha in feriados

def procesar_datos_eficientemente(ruta_archivo, feriados, bloque_size=50000):
    try:
        print("Leyendo el archivo completo...")
        columnas_necesarias = ['Inicio_de_Conexión_Dia', 'Usuario']

        # Leer columnas necesarias y procesar en bloques
        datos = pd.read_excel(ruta_archivo, usecols=columnas_necesarias)
        datos['Inicio_de_Conexión_Dia'] = pd.to_datetime(datos['Inicio_de_Conexión_Dia'], errors='coerce')

        print("Aplicando validaciones y filtrando datos...")
        resultados = []

        for inicio in range(0, len(datos), bloque_size):
            fin = inicio + bloque_size
            bloque = datos.iloc[inicio:fin]

            # Validación de usuarios con expresiones regulares (sólo alfanuméricos)
            bloque['Usuario_Valido'] = bloque['Usuario'].apply(lambda x: validar_formato(x, r'^[a-zA-Z0-9_]+$'))

            # Filtrar días no laborales y usuarios válidos
            filtro = (
                bloque['Inicio_de_Conexión_Dia'].apply(lambda x: es_dia_no_laboral(x, feriados)) &
                bloque['Usuario_Valido']
            )
            resultados.append(bloque[filtro])

            print(f"Procesado bloque {inicio} a {fin}...")

        # Unir los resultados procesados
        datos_filtrados = pd.concat(resultados)

        # Exportar a Excel
        archivo_salida = 'conexiones_validadas.xlsx'
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

    procesar_datos_eficientemente(ruta_archivo, feriados)

if __name__ == "__main__":
    main()
