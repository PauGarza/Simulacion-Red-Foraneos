import random
import pandas as pd

# Parámetros fijos para cada dispositivo
SMART_TV_USO = 5  # Mbps
PHONE_USO = 3  # Mbps

def uso_cuarto(porcentaje_privado_total):
    # Determinar el estado ON/OFF de cada dispositivo basado en porcentaje_privado_total
    tv_on = random.random() < porcentaje_privado_total
    laptop_on = random.random() < porcentaje_privado_total
    phone_on = random.random() < porcentaje_privado_total
    tablet_on = random.random() < porcentaje_privado_total

    # Calcular el uso de Mbps para cada dispositivo
    tv_uso = SMART_TV_USO if tv_on else 0
    laptop_uso = (random.randint(1, 30)) if laptop_on else 0
    phone_uso = PHONE_USO if phone_on else 0
    tablet_uso = (random.randint(1, 8)) if tablet_on else 0

    # Sumar el uso total de Mbps en áreas privadas
    uso_total = tv_uso + laptop_uso + phone_uso + tablet_uso
    
    return uso_total

def uso_comun(porcentaje_comun_total, evento_esp):
    uso_mbps = 0
    
    if evento_esp:
        tv_uso = 10  # TV siempre ON si hay evento especial
        laptop_uso = (random.randint(1, 10)) if random.random() < porcentaje_comun_total else 0
        phone_uso = (random.randint(18, 47)) if random.random() < porcentaje_comun_total else 0
        tablet_uso = (random.randint(1, 10)) if random.random() < porcentaje_comun_total else 0
    else:
        tv_uso = 0  # TV siempre OFF si no hay evento especial
        laptop_uso = (random.randint(1, 10) * random.randint(1, 8)) if random.random() < porcentaje_comun_total else 0
        phone_uso = (random.randint(1, 10) * random.randint(1, 8)) if random.random() < porcentaje_comun_total else 0
        tablet_uso = (random.randint(1, 10) * random.randint(1, 8)) if random.random() < porcentaje_comun_total else 0

    # Sumar el uso total de Mbps en áreas comunes
    uso_total = tv_uso + laptop_uso + phone_uso + tablet_uso
    
    return uso_total

def sim_uso_red(porcentaje_ocupacion, uso_privado, uso_com, invitados, evento_esp):
    # Probabilidad de uso de dispositivos en áreas privadas y comunes
    private_USO_prob = uso_privado * porcentaje_ocupacion / 100
    comun_USO_prob = uso_com * porcentaje_ocupacion / 100

    # Simulación del uso en áreas privadas
    uso_cuarto_mbps = {key: 0 for key in 'ABCDEFGH'}

    for key in uso_cuarto_mbps:
        uso_cuarto_mbps[key] = uso_cuarto(private_USO_prob)
    
    uso_comun_mbps = {
        'COCINA': 0,
        'JARDIN': 0
    }

    for key in uso_comun_mbps:
        uso_comun_mbps[key] = uso_comun(comun_USO_prob, evento_esp)

    uso_router_mbps = {
        'R1': 0,
        'R2': 0,
        'R3': 0
    }

    uso_router_mbps['R3'] = uso_cuarto_mbps['G'] + uso_cuarto_mbps['H'] + uso_comun_mbps['COCINA'] + uso_comun_mbps['JARDIN']
    uso_router_mbps['R2'] = uso_cuarto_mbps['A'] + uso_cuarto_mbps['B'] + uso_cuarto_mbps['C']
    uso_router_mbps['R1'] = uso_cuarto_mbps['D'] + uso_cuarto_mbps['E'] + uso_cuarto_mbps['F'] + uso_router_mbps['R3'] + uso_router_mbps['R2']
    
    result = {**uso_cuarto_mbps, **uso_comun_mbps, **uso_router_mbps}
    return result

def realizar_simulaciones(n, porcentaje_ocupacion, uso_privado, uso_com, invitados, evento_esp, plan_mbps):
    resultados = []
    for _ in range(n):
        resultado = sim_uso_red(porcentaje_ocupacion, uso_privado, uso_com, invitados, evento_esp)
        resultado['Plan_Mbps'] = plan_mbps
        resultado['Uso_Total_Mbps'] = sum([v for k, v in resultado.items() if k not in ['Plan_Mbps']])
        resultados.append(resultado)
    return resultados

# Parámetros de la simulación
n_simulaciones = 30
porcentaje_ocupacion = 75
uso_privado = 60
uso_com = 40
invitados = 0
evento_esp = False
plan_mbps = 300

# Realizar simulaciones
resultados = realizar_simulaciones(n_simulaciones, porcentaje_ocupacion, uso_privado, uso_com, invitados, evento_esp, plan_mbps)

# Crear un DataFrame con los resultados
df = pd.DataFrame(resultados)

# Guardar los resultados en un archivo Excel
df.to_excel('resultados_simulacion_ajustados.xlsx', index=False)

print("Simulaciones completadas y resultados guardados en 'resultados_simulacion_ajustados.xlsx'")
