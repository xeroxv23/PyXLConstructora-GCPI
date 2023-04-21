import tkinter as tk
import openpyxl
import os
import xlwings as xw
from celda_para_captura import obtener_celda_para_captura

# Funciones que usan los valores ingresados por el usuario
def funcion1(bodeguero, numero_de_semana):
    # Hacer algo con bodeguero y numero_de_semana
    pass

def funcion2(ruta_carpeta_semana, dias_bodeguero):
    # Hacer algo con ruta_carpeta_semana y dias_bodeguero
    pass

def funcion3(vacaciones_bodeguero):
    # Hacer algo con vacaciones_bodeguero
    pass

# Crear la ventana de la aplicación
root = tk.Tk()

# Crear los campos de entrada de texto
bodeguero_entry = tk.Entry(root)
numero_de_semana_entry = tk.Entry(root)
ruta_carpeta_semana_entry = tk.Entry(root)
dias_bodeguero_entry = tk.Entry(root)
vacaciones_bodeguero_entry = tk.Entry(root)

# Crear las etiquetas para los campos de entrada de texto
bodeguero_label = tk.Label(root, text="¿A qué bodeguero piensas capturar?")
numero_de_semana_label = tk.Label(root, text="Ingrese el número de semana que estará trabajando:")
ruta_carpeta_semana_label = tk.Label(root, text="Ingrese la ruta de la carpeta de la semana:")
dias_bodeguero_label = tk.Label(root, text="¿Cuántos días trabajó el bodeguero?")
vacaciones_bodeguero_label = tk.Label(root, text="¿El bodeguero tomó vacaciones? (Y/N):")

# Crear el botón de ejecución
ejecutar_button = tk.Button(root, text="Ejecutar", command=lambda: ejecutar_codigo(bodeguero_entry.get(), int(numero_de_semana_entry.get()), ruta_carpeta_semana_entry.get(), int(dias_bodeguero_entry.get()), vacaciones_bodeguero_entry.get()))

# Ubicar los widgets en la ventana
bodeguero_label.grid(row=0, column=0)
bodeguero_entry.grid(row=0, column=1)
numero_de_semana_label.grid(row=1, column=0)
numero_de_semana_entry.grid(row=1, column=1)
ruta_carpeta_semana_label.grid(row=2, column=0)
ruta_carpeta_semana_entry.grid(row=2, column=1)
dias_bodeguero_label.grid(row=3, column=0)
dias_bodeguero_entry.grid(row=3, column=1)
vacaciones_bodeguero_label.grid(row=4, column=0)
vacaciones_bodeguero_entry.grid(row=4, column=1)
ejecutar_button.grid(row=5, column=1)

# Función que se ejecuta cuando se presiona el botón de ejecución
def ejecutar_codigo(bodeguero, numero_de_semana, ruta_carpeta_semana, dias_bodeguero, vacaciones_bodeguero):
    if vacaciones_bodeguero.upper()