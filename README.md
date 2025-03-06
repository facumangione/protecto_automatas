# Seguimiento de Conexiones en Días Feriados y Fines de Semana 📅

Este proyecto permite analizar las conexiones de usuarios en una base de datos extensa, identificando específicamente las conexiones realizadas durante días feriados y fines de semana en Argentina. Además, ofrece la posibilidad de filtrar los datos por un rango de fechas definido por el usuario y exportar los resultados a un archivo Excel. Tarda aproximadamente 3 minutos en ejecutarse.

---

## Características ✨
- 📂 **Carga eficiente** de grandes volúmenes de datos desde archivos Excel.
- 🔍 **Filtrado de conexiones** realizadas en días feriados, sábados y domingos.
- 📅 **Selección de rango de fechas** para personalizar la búsqueda.
- 📤 **Exportación de resultados** a un archivo Excel.
- ⏱️ Optimización para trabajar con bases de datos extensas (más de un millón de registros).

---

## Instalación 🚀

Sigue estos pasos para instalar y ejecutar el proyecto en tu máquina:

### 1. Clonar el repositorio
Primero, clona este proyecto en tu computadora:
```bash
git clone https://github.com/tu_usuario/seguimiento-conexiones.git
```

### 2. Crear un entorno virtual
Segundo, crea un entorno virtual con el comando:
```bash
python -m venv venv
```
### Linux
```bash
source venv/bin/activate
```
### Windows
```bash
venv\Scripts\activate
```

### 3. Instalacion de dependencias 
Instala las dependencias necesarias ejecutando:
```bash
pip install -r requirements.txt
```
## Si no funciona ejecutar:
```bash
pip install --no-cache-dir pandas==2.0.3
```
```bash
pip install --no-cache-dir openpyxl==3.1.2
pip install --no-cache-dir tk
```

### 4. Uso 
Ejecuta el programa:
```bash
python main.py
```
