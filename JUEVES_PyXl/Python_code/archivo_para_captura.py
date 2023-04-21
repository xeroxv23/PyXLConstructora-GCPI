# LIBRERIAS
import os

# MODULOS
from datos_de_captura import obtener_datos_de_captura

# VARIABLES
ruta_archivo_origen = "C:\\Users\\carlo\\Google Drive\\GCPI - TRABAJO\\PyXLConstruc\\JUEVES_PyXl\\JUEVES_REPORTE.xlsx"

datos_de_captura, lista_de_actividades, trabajador, fecha_domingo, fecha_dia_festivo = obtener_datos_de_captura(ruta_archivo_origen)

def obtener_archivo_para_captura(clave):

        clave_de_obra = datos_de_captura[clave][1]
    
        # obtener la ruta de búsqueda
        ruta_busqueda = 'C:\\Users\\carlo\\Google Drive\\GCPI - TRABAJO\\PyXLConstruc\\JUEVES_PyXl\\SEMANA_12'

        # buscar archivos en la ruta de búsqueda que inicien con el valor de búsqueda
        for archivo in os.listdir(ruta_busqueda):
            if archivo.startswith(clave_de_obra + " ") and os.path.isfile(os.path.join(ruta_busqueda, archivo)):
                # si se encuentra el archivo, regresar la ruta completa
                archivo_para_captura = os.path.join(ruta_busqueda, archivo)

        return archivo_para_captura


