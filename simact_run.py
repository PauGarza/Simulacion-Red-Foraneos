import xlwings as xw
import pandas as pd

# Abrir el archivo Excel con xlwings
wb = xw.Book('simact.xlsx')  # Asegúrate que el archivo 'sim.xlsx' esté en el directorio correcto
sheet = wb.sheets[0]  # Selecciona la primera hoja

# Establecer los parámetros de simulación en el Excel
def establecer_parametros(sheet, parametros):
    for celda, valor in parametros.items():
        sheet.range(celda).value = valor
    wb.save()

# Configuración de los parámetros de la simulación
parametros = {
    'B3': 75,  # Plan en Mbps
    'B4': 0.2,   # Ocupación de la casa
    'B5': 0.8,   # Porcentaje de uso de la red de áreas privadas
    'B8': 0,    # Invitados en la casa
    'B9': 'OFF'  # Evento especial
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
    'U25': 'A',
    'Q25': 'B',
    'M25': 'C',
    'U4': 'D',
    'Q4': 'E',
    'M4': 'F',
    'G11': 'G',
    'G5': 'H',
    'G17': 'JARDIN',
    'G23': 'COCINA',
    'R15': 'R1',
    'R21': 'R2',
}

# Realizar múltiples recalculaciones y capturas
resultados = []
n_recalculos = 30  # Define cuántas veces quieres repetir el proceso

for _ in range(n_recalculos):
    valores = recalcular_y_capturar(sheet, 'E5', celdas_a_capturar)
    resultados.append(valores)

# Crear un DataFrame con los resultados
df = pd.DataFrame(resultados)

# Guardar los resultados en un nuevo archivo Excel
df.to_excel('vacaciones_act.xlsx', index=False)

print("Recalculos completados y resultados guardados en 'resultados_capturados.xlsx'")

# Cerrar el libro de Excel
wb.close()
