import xlwings as xw
import pandas as pd

# Abrir el archivo Excel con xlwings
wb = xw.Book('sim.xlsx')  # Cambia esto por la ruta de tu archivo Excel
sheet = wb.sheets[0]  # Selecciona la primera hoja

# Función para recalcular una celda y capturar valores de otras celdas
def recalcular_y_capturar(sheet, celda_a_recalcular, celdas_a_capturar):
    # Recalcular la celda
    sheet.range(celda_a_recalcular).value = sheet.range(celda_a_recalcular).value
    
    # Capturar los valores de las celdas especificadas
    valores = {celda: sheet.range(celda).value for celda in celdas_a_capturar}
    
    return valores

# Celdas a recalcular y capturar
celda_recalcular = 'E10'  # Cambia esto por la celda con RAND() u otra fórmula
celdas_a_capturar = ['K3', 'P3', 'T3']  # Cambia esto por las celdas que quieres capturar

# Realizar múltiples recalculaciones y capturas
resultados = []
n_recalculos = 10  # Cambia esto por el número de recalculaciones que deseas realizar

for _ in range(n_recalculos):
    valores = recalcular_y_capturar(sheet, celda_recalcular, celdas_a_capturar)
    resultados.append(valores)

# Crear un DataFrame con los resultados
df = pd.DataFrame(resultados)

# Guardar los resultados en un nuevo archivo Excel
df.to_excel('resultados_capturados.xlsx', index=False)

print("Recalculos completados y resultados guardados en 'resultados_capturados.xlsx'")

# Cerrar el libro de Excel
wb.close()
