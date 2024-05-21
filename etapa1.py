import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

def run_etapa1():
    # Rutas de archivos
    base_dir = 'optimizacion'
    data_dir = 'xlsx'
    graficas_dir = os.path.join(base_dir, 'graficas')
    
    if not os.path.exists(graficas_dir):
        os.makedirs(graficas_dir)

    demanda_path = os.path.join(data_dir, 'Dataton 2023 Etapa 1.xlsx')
    workers_path = os.path.join(data_dir, 'Dataton 2023 Etapa 1.xlsx')

    # Cargar los Datos
    demanda = pd.read_excel(demanda_path, sheet_name='demand')
    workers = pd.read_excel(workers_path, sheet_name='workers')

    ## Mostrar las primeras filas de los datos para revisión
    #print(demanda.head())
    #print(workers.head())

    # Extraer la hora de la columna 'fecha_hora'
    demanda['hora'] = pd.to_datetime(demanda['fecha_hora']).dt.strftime('%H:%M')

    # Parámetros
    minimo_trabajo_continuo = 4  # franjas de 15 minutos (1 hora)
    maximo_trabajo_continuo = 8  # franjas de 15 minutos (2 horas)
    tiempo_almuerzo = 6  # franjas de 15 minutos (1.5 horas)
    franja_almuerzo_min = 18  # 11:30 AM
    franja_almuerzo_max = 26  # 1:30 PM
    jornada_laboral = 32  # franjas de 15 minutos (8 horas)
    empleados = workers['documento'].unique()
    horarios = {empleado: [] for empleado in empleados}

    # Algoritmo de asignación
    def asignar_horarios(demanda, empleados):
        capacidad_actual = np.zeros(len(demanda))
        
        for i, row in demanda.iterrows():
            demanda_actual = row['demanda']
            for empleado in empleados:
                if len(horarios[empleado]) < jornada_laboral:
                    if capacidad_actual[i] < demanda_actual:
                        horarios[empleado].append('Trabaja')
                        capacidad_actual[i] += 1
                    elif capacidad_actual[i] == demanda_actual:
                        if len(horarios[empleado]) >= minimo_trabajo_continuo:
                            if len(horarios[empleado]) >= maximo_trabajo_continuo or (franja_almuerzo_min <= i <= franja_almuerzo_max and 'Almuerza' not in horarios[empleado]):
                                horarios[empleado].append('Almuerza' if franja_almuerzo_min <= i <= franja_almuerzo_max else 'Pausa Activa')
                            else:
                                horarios[empleado].append('Trabaja')
                        else:
                            horarios[empleado].append('Trabaja')
                    else:
                        horarios[empleado].append('Trabaja')
                else:
                    horarios[empleado].append('Nada')

    asignar_horarios(demanda, empleados)

    def graficar_demanda_vs_capacidad(demanda, capacidad, titulo, filename):
        plt.figure(figsize=(10, 6))
    
        # Ancho de las barras
        bar_width = 0.80
        separador_width = 0.001
    
        # Posiciones de las barras
        positions = np.arange(len(demanda))
    
        # Graficar barras de demanda con separadores blancos
        for pos in positions:
            plt.bar(pos, demanda['demanda'][pos], color='gray', width=bar_width)
            plt.bar(pos + bar_width, 0, color='white', width=separador_width)  # Separador blanco
    
        # Graficar la capacidad
        plt.plot(positions, capacidad, label='Demanda', color='none', linewidth=8)
        plt.plot(positions, capacidad, label='Capacidad', color='blue', linewidth=3)
    
    
        plt.xlabel('Hora del día')
        plt.ylabel('Demanda / Capacidad')
        plt.title(titulo)
        plt.legend()
        plt.xticks(ticks=positions, labels=demanda['hora'], rotation=90, ha='right', fontsize=8)
        plt.tight_layout()
        plt.savefig(os.path.join(graficas_dir, filename))
        plt.close()

    def graficar_diferencia(demanda, capacidad, titulo, filename):
        diferencia = demanda['demanda'] - capacidad
        plt.figure(figsize=(24, 12))
        plt.bar(demanda['fecha_hora'], diferencia, label='Demanda - Capacidad', color=['red' if diff > 0 else 'green' for diff in diferencia], width=0.02)
        plt.xlabel('Hora del día')
        plt.ylabel('Diferencia')
        plt.title(titulo)
        plt.axhline(0, color='black', linewidth=0.5)
        plt.legend()
        plt.xticks(rotation=90, ha='right', fontsize=4)
        plt.tight_layout()
        plt.savefig(os.path.join(graficas_dir, filename))
        plt.close()

    # Graficar resultados
    capacidad = np.zeros(len(demanda))
    for i in range(len(demanda)):
        capacidad[i] = sum(1 for horario in horarios.values() if horario[i] == 'Trabaja')

    graficar_demanda_vs_capacidad(demanda, capacidad, 'Capacidad vs. Demanda (Etapa 1)', 'Gpy1_etapa1.png')
    graficar_diferencia(demanda, capacidad, 'Demanda - Capacidad (Etapa 1)', 'Gpy2_etapa1.png')

    # Guardar Resultados
    horarios_df = pd.DataFrame.from_dict(horarios, orient='index').transpose()
    horarios_df.index = demanda['fecha_hora']
    horarios_df.reset_index(inplace=True)
    horarios_df.rename(columns={'index': 'fecha_hora'}, inplace=True)
    horarios_df.to_excel(os.path.join(base_dir, 'xlsx', 'horarios_optimizados_etapa1.xlsx'), index=False)

if __name__ == "__main__":
    run_etapa1()
