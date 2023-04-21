# LIBRERIAS
import xlwings as xw

# MODULOS
from datos_de_captura import datos_de_captura, lista_de_actividades, fecha_domingo, fecha_dia_festivo
from archivo_para_captura import obtener_archivo_para_captura
from celda_para_captura import obtener_celda_para_captura

def capturar_trabajador(clave):

    # VARIABLES DE LA FUNCION
    archivo_para_captura = obtener_archivo_para_captura(clave)
    celda_para_captura, ultimo_valor = obtener_celda_para_captura(archivo_para_captura)

    # CODIGOS DE DATOS DE CAPTURA
    codigo_de_nomina = 0
    obra = 1
    dias_trabajados = 2
    horas_extras = 3
    domingo_trabajado = 4
    actividades = 6
    velador = 7
    dia_festivo = 8
    sueldo = 9
    

    # Cargamos el archivo_para_captura de Excel
    wb = xw.Book(archivo_para_captura)
    # Seleccionamos la hoja en la que queremos buscar
    ws = wb.sheets.active

    # Tomamos el valor de la ultima celda de captura
    fila = celda_para_captura.row
    columna = celda_para_captura.column

    # Asignar el codigo de nomina del trabajador
    codigo = ws.range(fila, columna).value = datos_de_captura[clave][codigo_de_nomina]

    # Asignar los valores:

    # CASO NO.1 // SOLO DIAS TRABAJADOS
    if datos_de_captura[clave][horas_extras] is None and datos_de_captura[clave][domingo_trabajado] is None and datos_de_captura[clave][velador] is None and datos_de_captura[clave][dia_festivo] is None:

        celda_orden1 = ws.range(fila, columna +1).value = ultimo_valor
        celda_orden2 = ws.range(fila, columna +15).value = ultimo_valor
        celda_dias = ws.range(fila, columna +11).value = datos_de_captura[clave][dias_trabajados]
        celda_porcentaje = ws.range(fila, columna +18).value = 1
    
    # CASO NO 1.2 VELADOR // SOLO DIAS TRABAJADOS
    elif datos_de_captura[clave][horas_extras] is None and datos_de_captura[clave][domingo_trabajado] is None and datos_de_captura[clave][velador] is not None and datos_de_captura[clave][dia_festivo] is None:

        celda_orden1 = ws.range(fila, columna +1).value = ultimo_valor
        celda_dias = ws.range(fila, columna +11).value = datos_de_captura[clave][dias_trabajados]
        celda_velador = ws.range(fila +1, columna).value = "vel01"
        celda_orden2 = ws.range(fila +1, columna +1).value = ultimo_valor
        celda_velador2 = ws.range(fila +1, columna +11).value = datos_de_captura[clave][velador]
        celda_orden3 = ws.range(fila +1, columna +15).value = ultimo_valor
        celda_porcentaje = ws.range(fila, columna +18).value = 1

    # CASO NO 1.3 VELADOR // MAS DIA FESTIVO TRABAJADO
    elif datos_de_captura[clave][horas_extras] is None and datos_de_captura[clave][domingo_trabajado] is None and datos_de_captura[clave][velador] is not None and datos_de_captura[clave][dia_festivo] is not None:

        celda_orden1 = ws.range(fila, columna +1).value = ultimo_valor
        celda_dias = ws.range(fila, columna +11).value = datos_de_captura[clave][dias_trabajados]
        celda_velador = ws.range(fila +1, columna).value = "vel01"
        celda_orden2 = ws.range(fila +1, columna +1).value = ultimo_valor
        celda_velador2 = ws.range(fila +1, columna +11).value = datos_de_captura[clave][velador]
        celda_dia_festivo = ws.range(fila +2, columna).value = "lote"
        celda_orden3 = ws.range(fila +2, columna +1).value = ultimo_valor
        celda_orden4 = ws.range(fila +2, columna +15).value = ultimo_valor
        celda_dia_festivo2 = ws.range(fila +2, columna +3).value = (f"Dia festivo {fecha_dia_festivo} trabajado")
        celda_dia_festivo3 = ws.range(fila +2, columna +8).value = ((datos_de_captura[clave][sueldo]) / 100)
        celda_dia_festivo4 = ws.range(fila +2, columna +11).value = 4
        celda_porcentaje = ws.range(fila, columna +18).value = 1

    # CASO NO.3 // SI SOLAMENTE HAY HORAS EXTRAS
    elif datos_de_captura[clave][horas_extras] is not None and datos_de_captura[clave][domingo_trabajado] is None and datos_de_captura[clave][dia_festivo] is None:

        celda_orden1 = ws.range(fila, columna +1).value = ultimo_valor
        celda_horas = ws.range(fila +1, columna).value = "lote"
        celda_orden2 = ws.range(fila +1, columna +1).value = ultimo_valor
        celda_horas2 = ws.range(fila +1, columna +8).value = float(datos_de_captura[clave][sueldo]) * 0.0025 
        celda_orden3 = ws.range(fila +1, columna +15).value = ultimo_valor
        celda_horas3 = ws.range(fila +1, columna +11).value = datos_de_captura[clave][horas_extras]
        celda_horas4 = ws.range(fila +1, columna +3).value = (f"Tiempo extra, {datos_de_captura[clave][horas_extras]} horas trabajadas")
        celda_dias = ws.range(fila, columna +11).value = datos_de_captura[clave][dias_trabajados]
        celda_porcentaje = ws.range(fila, columna +18).value = 1

    # CASO NO.4 // SI NO HAY HORAS EXTRAS, PERO SI HAY DOMINGO
    elif datos_de_captura[clave][horas_extras] is None and datos_de_captura[clave][domingo_trabajado] is not None and datos_de_captura[clave][dia_festivo] is None:
            
        celda_orden1 = ws.range(fila, columna +1).value = ultimo_valor
        celda_domingo = ws.range(fila +1, columna).value = "lote"
        celda_orden2 = ws.range(fila +1, columna +1).value = ultimo_valor
        celda_domingo2 = ws.range(fila +1, columna +8).value = ((datos_de_captura[clave][sueldo]) / 100)
        celda_orden3 = ws.range(fila +1, columna +15).value = ultimo_valor
        celda_domingo3 = ws.range(fila +1, columna +11).value = datos_de_captura[clave][domingo_trabajado] + 1
        celda_domingo4 = ws.range(fila +1, columna +3).value = (f"Domingo {fecha_domingo} trabajado")
        celda_dias = ws.range(fila, columna +11).value = datos_de_captura[clave][dias_trabajados]
        celda_porcentaje = ws.range(fila, columna +18).value = 1

    # CASO NO.5 // SI TENEMOS HORAS EXTRAS Y DOMINGO TRABAJADO - SIN DIA FESTIVO
    elif datos_de_captura[clave][horas_extras] is not None and datos_de_captura[clave][domingo_trabajado] is not None and datos_de_captura[clave][dia_festivo] is None:

        celda_orden1 = ws.range(fila, columna +1).value = ultimo_valor
        celda_orden2 = ws.range(fila +1, columna +1).value = ultimo_valor
        celda_orden3 = ws.range(fila +2, columna +1).value = ultimo_valor
        celda_orden4 = ws.range(fila +2, columna +15).value = ultimo_valor
        celda_dias = ws.range(fila, columna +11).value = datos_de_captura[clave][dias_trabajados]
        celda_porcentaje = ws.range(fila, columna +18).value = 1
        celda_horas = ws.range(fila +1, columna).value = "lote"
        celda_horas2 = ws.range(fila +1, columna +8).value = float(datos_de_captura[clave][sueldo]) * 0.0025
        celda_horas3 = ws.range(fila +1, columna +11).value = datos_de_captura[clave][horas_extras]
        celda_horas4 = ws.range(fila +1, columna +3).value = (f"Tiempo extra, {datos_de_captura[clave][horas_extras]} horas trabajadas")
        celda_domingo = ws.range(fila +2, columna).value = "lote"
        celda_domingo2 = ws.range(fila +2, columna +8).value = ((datos_de_captura[clave][sueldo]) / 100)
        celda_domingo3 = ws.range(fila +2, columna +11).value = datos_de_captura[clave][domingo_trabajado] + 1
        celda_domingo4 = ws.range(fila +2, columna +3).value = (f"Domingo {fecha_domingo} trabajado")
    
    # CASO NO.6 // SI TENEMOS DIA FESTIVO
    elif datos_de_captura[clave][dia_festivo] is not None and datos_de_captura[clave][horas_extras] is None and datos_de_captura[clave][domingo_trabajado] is None:

        celda_orden1 = ws.range(fila, columna +1).value = ultimo_valor
        celda_orden2 = ws.range(fila +1, columna +15).value = ultimo_valor
        celda_orden3 = ws.range(fila +1, columna +1).value = ultimo_valor
        celda_dias = ws.range(fila, columna +11).value = datos_de_captura[clave][dias_trabajados]
        celda_porcentaje = ws.range(fila, columna +18).value = 1
        celda_dia_festivo = ws.range(fila +1, columna).value = "lote"
        celda_dia_festivo2 = ws.range(fila +1, columna +3).value = (f"Dia festivo {fecha_dia_festivo} trabajado")
        celda_dia_festivo3 = ws.range(fila +1, columna +8).value = ((datos_de_captura[clave][sueldo]) / 100)
        celda_dia_festivo4 = ws.range(fila +1, columna +11).value = 2
    
    # CASNO NO.7 // SI TENEMOS DIA FESTIVO Y HORAS EXTRAS
    elif datos_de_captura[clave][dia_festivo] is not None and datos_de_captura[clave][horas_extras] is not None and datos_de_captura[clave][domingo_trabajado] is None:

        celda_orden1 = ws.range(fila, columna +1).value = ultimo_valor
        celda_orden2 = ws.range(fila +2, columna +15).value = ultimo_valor
        celda_orden3 = ws.range(fila +1, columna +1).value = ultimo_valor
        celda_orden4 = ws.range(fila +2, columna +1).value = ultimo_valor
        celda_dias = ws.range(fila, columna +11).value = datos_de_captura[clave][dias_trabajados]
        celda_porcentaje = ws.range(fila, columna +18).value = 1
        celda_dia_festivo = ws.range(fila +1, columna).value = "lote"
        celda_dia_festivo2 = ws.range(fila +1, columna +3).value = (f"Dia festivo {fecha_dia_festivo} trabajado")
        celda_dia_festivo3 = ws.range(fila +1, columna +8).value = ((datos_de_captura[clave][sueldo]) / 100)
        celda_dia_festivo4 = ws.range(fila +1, columna +11).value = 2
        celda_horas = ws.range(fila +2, columna).value = "lote"
        celda_horas2 = ws.range(fila +2, columna +8).value = float(datos_de_captura[clave][sueldo]) * 0.0025
        celda_horas3 = ws.range(fila +2, columna +11).value = datos_de_captura[clave][horas_extras]
        celda_horas4 = ws.range(fila +2, columna +3).value = (f"Tiempo extra, {datos_de_captura[clave][horas_extras]} horas trabajadas")

    # CASNO NO.8 // SI TENEMOS DIA FESTIVO, HORAS EXTRAS Y DIA FESTIVO TRABAJADO
    elif datos_de_captura[clave][dia_festivo] is not None and datos_de_captura[clave][horas_extras] is not None and datos_de_captura[clave][domingo_trabajado] is not None:

        celda_orden1 = ws.range(fila, columna +1).value = ultimo_valor
        celda_orden2 = ws.range(fila +3, columna +15).value = ultimo_valor
        celda_orden3 = ws.range(fila +1, columna +1).value = ultimo_valor
        celda_orden4 = ws.range(fila +2, columna +1).value = ultimo_valor
        celda_orden5 = ws.range(fila +3, columna +1).value = ultimo_valor
        celda_dias = ws.range(fila, columna +11).value = datos_de_captura[clave][dias_trabajados]
        celda_porcentaje = ws.range(fila, columna +18).value = 1
        celda_dia_festivo = ws.range(fila +1, columna).value = "lote"
        celda_dia_festivo2 = ws.range(fila +1, columna +3).value = (f"Dia festivo {fecha_dia_festivo} trabajado")
        celda_dia_festivo3 = ws.range(fila +1, columna +8).value = ((datos_de_captura[clave][sueldo]) / 100)
        celda_dia_festivo4 = ws.range(fila +1, columna +11).value = 2
        celda_horas = ws.range(fila +2, columna).value = "lote"
        celda_horas2 = ws.range(fila +2, columna +8).value = float(datos_de_captura[clave][sueldo]) * 0.0025
        celda_horas3 = ws.range(fila +2, columna +11).value = datos_de_captura[clave][horas_extras]
        celda_horas4 = ws.range(fila +2, columna +3).value = (f"Tiempo extra, {datos_de_captura[clave][horas_extras]} horas trabajadas")
        celda_domingo = ws.range(fila +3, columna).value = "lote"
        celda_domingo2 = ws.range(fila +3, columna +8).value = ((datos_de_captura[clave][sueldo]) / 100)
        celda_domingo3 = ws.range(fila +3, columna +11).value = datos_de_captura[clave][domingo_trabajado] + 1
        celda_domingo4 = ws.range(fila +3, columna +3).value = (f"Domingo {fecha_domingo} trabajado")
    
    # Asignar actividades

    # CASO NO.1 // SI TENEMOS ACTIVIDADES, HORAS EXTRAS Y DOMINGO TRABAJADO
    if datos_de_captura[clave][actividades] is not None and datos_de_captura[clave][horas_extras] is not None and datos_de_captura[clave][domingo_trabajado] is not None and datos_de_captura[clave][dia_festivo] is None:

        for valor in lista_de_actividades[clave]:
            celda_actividades = ws.range(fila +3, columna +3)
            celda_actividades.value = valor
            fila += 1
    
    # CASO NO.2 // SI TENEMOS ACTIVIDADES Y TENEMOS VELADOR
    elif datos_de_captura[clave][actividades] is not None and datos_de_captura[clave][velador] is not None and datos_de_captura[clave][dia_festivo] is None:

        for valor in lista_de_actividades[clave]:
            celda_actividades = ws.range(fila +2, columna +3)
            celda_actividades.value = valor
            fila += 1

    # CASO NO.3 // SI TENEMOS ACTIVIDADES, TENEMOS VELADOR Y DIA FESTIVO
    elif datos_de_captura[clave][actividades] is not None and datos_de_captura[clave][velador] is not None and datos_de_captura[clave][dia_festivo] is not None:

        for valor in lista_de_actividades[clave]:
            celda_actividades = ws.range(fila +3, columna +3)
            celda_actividades.value = valor
            fila += 1

    # CASO NO.4 SI TENEMOS ACTIVIDADES Y SOLAMENTE HORAS EXTRAS
    elif datos_de_captura[clave][actividades] is not None and datos_de_captura[clave][horas_extras] is not None and datos_de_captura[clave][domingo_trabajado] is None and datos_de_captura[clave][dia_festivo] is None:

        for valor in lista_de_actividades[clave]:
            celda_actividades = ws.range(fila +2, columna +3)
            celda_actividades.value = valor
            fila += 1

    # CASO NO.5 SI TENEMOS ACTIVIDADES Y SOLAMENTE DOMINGO TRABAJADO
    elif datos_de_captura[clave][actividades] is not None and datos_de_captura[clave][horas_extras] is None and datos_de_captura[clave][domingo_trabajado] is not None and datos_de_captura[clave][dia_festivo] is None:

        for valor in lista_de_actividades[clave]:
            celda_actividades = ws.range(fila +2, columna +3)
            celda_actividades.value = valor
            fila += 1

    # CASO NO.6 SI TENEMOS SOLAMENTE ACTIVIDADES
    elif datos_de_captura[clave][actividades] is not None and datos_de_captura[clave][horas_extras] is None and datos_de_captura[clave][domingo_trabajado] is None and datos_de_captura[clave][dia_festivo] is None:

        for valor in lista_de_actividades[clave]:
            celda_actividades = ws.range(fila +1, columna +3)
            celda_actividades.value = valor
            fila += 1

    # CASO NO.7 SI TENEMOS ACTIVIDADES Y DIA FESTIVO
    elif datos_de_captura[clave][actividades] is not None and datos_de_captura[clave][horas_extras] is None and datos_de_captura[clave][domingo_trabajado] is None and datos_de_captura[clave][dia_festivo] is not None:

        for valor in lista_de_actividades[clave]:
            celda_actividades = ws.range(fila +2, columna +3)
            celda_actividades.value = valor
            fila += 1

    # CASO NO.8 SI TENEMOS ACTIVIDADES, DIA FESTIVO Y HORAS EXTRAS
    elif datos_de_captura[clave][actividades] is not None and datos_de_captura[clave][horas_extras] is not None and datos_de_captura[clave][domingo_trabajado] is None and datos_de_captura[clave][dia_festivo] is not None:

        for valor in lista_de_actividades[clave]:
            celda_actividades = ws.range(fila +3, columna +3)
            celda_actividades.value = valor
            fila += 1

    # CASO NO.8 SI TENEMOS ACTIVIDADES, DIA FESTIVO, HORAS EXTRAS Y DOMINGO TRABAJADO
    elif datos_de_captura[clave][actividades] is not None and datos_de_captura[clave][horas_extras] is not None and datos_de_captura[clave][domingo_trabajado] is not None and datos_de_captura[clave][dia_festivo] is not None:

        for valor in lista_de_actividades[clave]:
            celda_actividades = ws.range(fila +4, columna +3)
            celda_actividades.value = valor
            fila += 1
    
    # ? SI NO TENEMOS NI ACTIVIDADES
    else:
        pass

    # Guardar y cerrar el libro de Excel
    wb.save(archivo_para_captura)
    
    return print("Se ha capturado en la obra", datos_de_captura[clave][obra], "el trabajador numero : ", datos_de_captura[clave][codigo_de_nomina], "En la celda:", celda_para_captura.coordinate)

