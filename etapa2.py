import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

def run_etapa2():
    # Rutas de archivos
    base_dir = 'optimizacion'
    data_dir = os.path.join(base_dir, 'xlsx')
    graficas_dir = os.path.join(base_dir, 'graficas')
    
    if not os.path.exists(graficas_dir):
        os.makedirs(graficas_dir)

    demanda_path = os.path.join(data_dir, 'Dataton 2023 Etapa 2.xlsx')
    workers_path = os.path.join(data_dir, 'Dataton 2023 Etapa 2.xlsx')

    # Cargar los Datos
    demanda = pd.read_excel(demanda_path, sheet_name='demand')
    workers = pd.read_excel(workers_path, sheet_name='workers')

    # Mostrar las primeras filas de los datos para revisión
    print(demanda.head())
    print(workers.head())

    # Parámetros
    minimo_trabajo_continuo = 4  # franjas de 15 minutos (1 hora)
    maximo_trabajo_continuo = 8  # franjas de 15 minutos (2 horas)
    tiempo_almuerzo = 6  # franjas de 15 minutos (1.5 horas)
    franja_almuerzo_min = 18  # 11:30 AM
    franja_almuerzo_max = 26  # 1:30 PM
    jornada_completa = 28  # franjas de 15 minutos (7 horas)
    jornada_medio_tiempo = 16  # franjas de 15 minutos (4 horas)
    jornada_sabado_completa = 20  # franjas de 15 minutos (5 horas)
    empleados = workers['documento'].unique()
    horarios = {empleado: [] for empleado in empleados}

    # Algoritmo de asignación
    def asignar_horarios(demanda, empleados, workers):
        capacidad_actual = np.zeros(len(demanda))
        
        for i, row in demanda.iterrows():
            demanda_actual = row['demanda']
            for empleado in empleados:
                tipo_contrato = workers[workers['documento'] == empleado]['contrato'].values[0]
                dia_semana = (i // 96) % 6  # 96 franjas de 15 minutos por día y 6 días (lunes a sábado)
                jornada = jornada_completa if tipo_contrato == 'TC' else jornada_medio_tiempo
                if dia_semana == 5:  # Sábado
                    jornada = jornada_sabado_completa if tipo_contrato == 'TC' else jornada_medio_tiempo
                
                if len(horarios[empleado]) < jornada:
                    if capacidad_actual[i] < demanda_actual:
                        horarios[empleado].append('Trabaja')
                        capacidad_actual[i] += 1
                    elif capacidad_actual[i] == demanda_actual:
                        if len(horarios[empleado]) >= minimo_trabajo_continuo:
                            if tipo_contrato == 'TC' and (franja_almuerzo_min <= i % 96 <= franja_almuerzo_max and 'Almuerza' not in horarios[empleado]):
                                horarios[empleado].append('Almuerza')
                            else:
                                horarios[empleado].append('Pausa Activa')
                        else:
                            horarios[empleado].append('Trabaja')
                    else:
                        horarios[empleado].append('Trabaja')
                else:
                    horarios[empleado].append('Nada')

    asignar_horarios(demanda, empleados, workers)

    # Visualización y Validación
    def graficar_demanda_vs_capacidad(demanda, capacidad, titulo, filename):
        plt.figure(figsize=(12, 6))
        plt.bar(demanda['fecha_hora'], demanda['demanda'], label='Demanda', color='gray')
        plt.plot(demanda['fecha_hora'], capacidad, label='Capacidad', color='blue')
        plt.xlabel('Hora del día')
        plt.ylabel('Demanda / Capacidad')
        plt.title(titulo)
        plt.legend()
        plt.xticks(rotation=45)
        plt.savefig(os.path.join(graficas_dir, filename))
        plt.close()

    def graficar_diferencia(demanda, capacidad, titulo, filename):
        diferencia = demanda['demanda'] - capacidad
        plt.figure(figsize=(12, 6))
        plt.bar(demanda['fecha_hora'], diferencia, label='Demanda - Capacidad', color=['red' if diff > 0 else 'green' for diff in diferencia])
        plt.xlabel('Hora del día')
        plt.ylabel('Diferencia')
        plt.title(titulo)
        plt.axhline(0, color='black', linewidth=0.5)
        plt.legend()
        plt.xticks(rotation=45)
        plt.savefig(os.path.join(graficas_dir, filename))
        plt.close()

    # Graficar resultados
    capacidad = np.zeros(len(demanda))
    for i in range(len(demanda)):
        capacidad[i] = sum(1 for horario in horarios.values() if horario[i] == 'Trabaja')

    graficar_demanda_vs_capacidad(demanda, capacidad, 'Capacidad vs. Demanda (Etapa 2)', 'Gpy1_etapa2.png')
    graficar_diferencia(demanda, capacidad, 'Demanda - Capacidad (Etapa 2)', 'Gpy2_etapa2.png')

    # Guardar Resultados
    horarios_df = pd.DataFrame.from_dict(horarios, orient='index').transpose()
    horarios_df.index = demanda['fecha_hora']
    horarios_df.reset_index(inplace=True)
    horarios_df.rename(columns={'index': 'fecha_hora'}, inplace=True)
    horarios_df.to_excel(os.path.join(data_dir, 'horarios_optimizados_etapa2.xlsx'), index=False)

if __name__ == "__main__":
    run_etapa2()
