# LIBRERIAS
import xlwings as xw

# MODULOS
from datos_de_captura import trabajador
from archivo_para_captura import obtener_archivo_para_captura
from celda_para_captura import obtener_celda_para_captura
from capturar_trabajador import capturar_trabajador
from destajista import buscar_celda_en_archivo, capturar_destajista
from mandos_intermedios import buscar_celda_mandos_intermedios, valor_total_destajo, capturar_mandos_intermedios

for clave in trabajador:
    archivo_para_captura = obtener_archivo_para_captura(clave)
    celda_para_captura, ultimo_valor = obtener_celda_para_captura(archivo_para_captura)
    capturar_trabajador(clave)

print("Se han capturado los trabajadores del reporte")

for clave in trabajador:

    celda_destajista = buscar_celda_en_archivo(clave)
    capturar_destajista(clave)
    celda_mandos_intermedios = buscar_celda_mandos_intermedios(clave)
    total_destajo = valor_total_destajo(clave)
    capturar_mandos_intermedios(clave)


print("Se capturo destajista y mandos intermedios")