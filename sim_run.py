import xlwings as xw
import pandas as pd

# Abrir el archivo Excel con xlwings
wb = xw.Book('sim.xlsx')  # Asegúrate que el archivo 'sim.xlsx' esté en el directorio correcto
sheet = wb.sheets[0]  # Selecciona la primera hoja

# Establecer los parámetros de simulación en el Excel
def establecer_parametros(sheet, parametros):
    for celda, valor in parametros.items():
        sheet.range(celda).value = valor
    wb.save()

# Configuración de los parámetros de la simulación
parametros = {
    'B2': 300,  # Plan en Mbps
    'B3': 0.2,   # Ocupación de la casa
    'B4': 0.8,   # Porcentaje de uso de la red de áreas privadas
    'B7': 0,    # Invitados en la casa
    'B8': 'OFF'  # Evento especial
}
establecer_parametros(sheet, parametros)

# Función para recalcular una celda y capturar valores de otras celdas
def recalcular_y_capturar(sheet, celda_a_recalcular, celdas_a_capturar):
    # Recalcular la celda, forzando el recálculo de Excel
    sheet.range(celda_a_recalcular).value = sheet.range(celda_a_recalcular).value
    wb.app.calculate()
    
    # Capturar los valores de las celdas especificadas
    valores = {nombre: int(sheet.range(celda).value) if sheet.range(celda).value is not None else None for celda, nombre in celdas_a_capturar.items()}
    
    return valores

# Mapeo de celdas a capturar con nombres descriptivos
celdas_a_capturar = {
    'T24': 'A',
    'P24': 'B',
    'L24': 'C',
    'T3': 'D',
    'P3': 'E',
    'L3': 'F',
    'G16': 'G',
    'G10': 'H',
    'G22': 'JARDIN',
    'G28': 'COCINA',
    'Q14': 'R1',
    'Q20': 'R2',
    'J14': 'R3'
}

# Realizar múltiples recalculaciones y capturas
resultados = []
n_recalculos = 30  # Define cuántas veces quieres repetir el proceso

for _ in range(n_recalculos):
    valores = recalcular_y_capturar(sheet, 'E10', celdas_a_capturar)
    resultados.append(valores)

# Crear un DataFrame con los resultados
df = pd.DataFrame(resultados)

# Guardar los resultados en un nuevo archivo Excel
df.to_excel('vacaciones.xlsx', index=False)

print("Recalculos completados y resultados guardados en 'resultados_capturados.xlsx'")

# Cerrar el libro de Excel
wb.close()
