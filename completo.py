import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

def run_completo():
    # Rutas de archivos
    base_dir = 'optimizacion'
    data_dir = os.path.join(base_dir, 'xlsx')
    graficas_dir = os.path.join(base_dir, 'graficas')
    
    if not os.path.exists(graficas_dir):
        os.makedirs(graficas_dir)

    demanda_etapa1_path = os.path.join(data_dir, 'Dataton 2023 Etapa 1.xlsx')
    workers_etapa1_path = os.path.join(data_dir, 'Dataton 2023 Etapa 1.xlsx')
    demanda_etapa2_path = os.path.join(data_dir, 'Dataton 2023 Etapa 2.xlsx')
    workers_etapa2_path = os.path.join(data_dir, 'Dataton 2023 Etapa 2.xlsx')

    # Cargar los Datos
    demanda_etapa1 = pd.read_excel(demanda_etapa1_path, sheet_name='demand')
    workers_etapa1 = pd.read_excel(workers_etapa1_path, sheet_name='workers')
    demanda_etapa2 = pd.read_excel(demanda_etapa2_path, sheet_name='demand')
    workers_etapa2 = pd.read_excel(workers_etapa2_path, sheet_name='workers')

    # Mostrar las primeras filas de los datos para revisión
    print(demanda_etapa1.head())
    print(workers_etapa1.head())
    print(demanda_etapa2.head())
    print(workers_etapa2.head())

    # Parámetros comunes
    minimo_trabajo_continuo = 4  # franjas de 15 minutos (1 hora)
    maximo_trabajo_continuo = 8  # franjas de 15 minutos (2 horas)
    tiempo_almuerzo = 6  # franjas de 15 minutos (1.5 horas)
    franja_almuerzo_min = 18  # 11:30 AM
    franja_almuerzo_max = 26  # 1:30 PM

    # Parámetros específicos para Etapa 1
    jornada_laboral_etapa1 = 32  # franjas de 15 minutos (8 horas)
    empleados_etapa1 = workers_etapa1['documento'].unique()
    horarios_etapa1 = {empleado: [] for empleado in empleados_etapa1}

    # Parámetros específicos para Etapa 2
    jornada_completa = 28  # franjas de 15 minutos (7 horas)
    jornada_medio_tiempo = 16  # franjas de 15 minutos (4 horas)
    jornada_sabado_completa = 20  # franjas de 15 minutos (5 horas)
    empleados_etapa2 = workers_etapa2['documento'].unique()
    horarios_etapa2 = {empleado: [] for empleado in empleados_etapa2}

    # Algoritmo de asignación para Etapa 1
    def asignar_horarios_etapa1(demanda, empleados):
        capacidad_actual = np.zeros(len(demanda))
        
        for i, row in demanda.iterrows():
            demanda_actual = row['demanda']
            for empleado in empleados:
                if len(horarios_etapa1[empleado]) < jornada_laboral_etapa1:
                    if capacidad_actual[i] < demanda_actual:
                        horarios_etapa1[empleado].append('Trabaja')
                        capacidad_actual[i] += 1
                    elif capacidad_actual[i] == demanda_actual:
                        if len(horarios_etapa1[empleado]) >= minimo_trabajo_continuo:
                            if len(horarios_etapa1[empleado]) >= maximo_trabajo_continuo or (franja_almuerzo_min <= i <= franja_almuerzo_max and 'Almuerza' not in horarios_etapa1[empleado]):
                                horarios_etapa1[empleado].append('Almuerza' if franja_almuerzo_min <= i <= franja_almuerzo_max else 'Pausa Activa')
                            else:
                                horarios_etapa1[empleado].append('Trabaja')
                        else:
                            horarios_etapa1[empleado].append('Trabaja')
                    else:
                        horarios_etapa1[empleado].append('Trabaja')
                else:
                    horarios_etapa1[empleado].append('Nada')

    asignar_horarios_etapa1(demanda_etapa1, empleados_etapa1)

    # Algoritmo de asignación para Etapa 2
    def asignar_horarios_etapa2(demanda, empleados, workers):
        capacidad_actual = np.zeros(len(demanda))
        
        for i, row in demanda.iterrows():
            demanda_actual = row['demanda']
            for empleado in empleados:
                tipo_contrato = workers[workers['documento'] == empleado]['contrato'].values[0]
                dia_semana = (i // 96) % 6  # 96 franjas de 15 minutos por día y 6 días (lunes a sábado)
                jornada = jornada_completa if tipo_contrato == 'TC' else jornada_medio_tiempo
                if dia_semana == 5:  # Sábado
                    jornada = jornada_sabado_completa if tipo_contrato == 'TC' else jornada_medio_tiempo
                
                if len(horarios_etapa2[empleado]) < jornada:
                    if capacidad_actual[i] < demanda_actual:
                        horarios_etapa2[empleado].append('Trabaja')
                        capacidad_actual[i] += 1
                    elif capacidad_actual[i] == demanda_actual:
                        if len(horarios_etapa2[empleado]) >= minimo_trabajo_continuo:
                            if tipo_contrato == 'TC' and (franja_almuerzo_min <= i % 96 <= franja_almuerzo_max and 'Almuerza' not in horarios_etapa2[empleado]):
                                horarios_etapa2[empleado].append('Almuerza')
                            else:
                                horarios_etapa2[empleado].append('Pausa Activa')
                        else:
                            horarios_etapa2[empleado].append('Trabaja')
                    else:
                        horarios_etapa2[empleado].append('Trabaja')
                else:
                    horarios_etapa2[empleado].append('Nada')

    asignar_horarios_etapa2(demanda_etapa2, empleados_etapa2, workers_etapa2)

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

    # Graficar resultados para Etapa 1
    capacidad_etapa1 = np.zeros(len(demanda_etapa1))
    for i in range(len(demanda_etapa1)):
        capacidad_etapa1[i] = sum(1 for horario in horarios_etapa1.values() if horario[i] == 'Trabaja')

    graficar_demanda_vs_capacidad(demanda_etapa1, capacidad_etapa1, 'Capacidad vs. Demanda (Etapa 1)', 'capacidad_vs_demanda_etapa1.png')
    graficar_diferencia(demanda_etapa1, capacidad_etapa1, 'Demanda - Capacidad (Etapa 1)', 'demanda_capacidad_etapa1.png')

    # Graficar resultados para Etapa 2
    capacidad_etapa2 = np.zeros(len(demanda_etapa2))
    for i in range(len(demanda_etapa2)):
        capacidad_etapa2[i] = sum(1 for horario in horarios_etapa2.values() if horario[i] == 'Trabaja')

    graficar_demanda_vs_capacidad(demanda_etapa2, capacidad_etapa2, 'Capacidad vs. Demanda (Etapa 2)', 'capacidad_vs_demanda_etapa2.png')
    graficar_diferencia(demanda_etapa2, capacidad_etapa2, 'Demanda - Capacidad (Etapa 2)', 'demanda_capacidad_etapa2.png')

    # Guardar Resultados
    horarios_df_etapa1 = pd.DataFrame.from_dict(horarios_etapa1, orient='index').transpose()
    horarios_df_etapa1.index = demanda_etapa1['fecha_hora']
    horarios_df_etapa1.reset_index(inplace=True)
    horarios_df_etapa1.rename(columns={'index': 'fecha_hora'}, inplace=True)

    horarios_df_etapa2 = pd.DataFrame.from_dict(horarios_etapa2, orient='index').transpose()
    horarios_df_etapa2.index = demanda_etapa2['fecha_hora']
    horarios_df_etapa2.reset_index(inplace=True)
    horarios_df_etapa2.rename(columns={'index': 'fecha_hora'}, inplace=True)

    horarios_df_etapa1.to_excel(os.path.join(data_dir, 'horarios_optimizados_etapa1.xlsx'), index=False)
    horarios_df_etapa2.to_excel(os.path.join(data_dir, 'horarios_optimizados_etapa2.xlsx'), index=False)

if __name__ == "__main__":
    run_completo()