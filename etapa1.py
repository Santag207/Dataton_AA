import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

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
            empleados_disponibles = [empleado for empleado in empleados if len(horarios[empleado]) < jornada_laboral]

            # Primero asignamos empleados hasta cubrir la demanda
            while capacidad_actual[i] < demanda_actual and empleados_disponibles:
                empleado = empleados_disponibles.pop(0)
                horarios[empleado].append('Trabaja')
                capacidad_actual[i] += 1

            # Luego asignamos empleados para mantener mínimo uno trabajando
            if capacidad_actual[i] == 0 and empleados_disponibles:
                empleado = empleados_disponibles.pop(0)
                horarios[empleado].append('Trabaja')
                capacidad_actual[i] += 1

            # Asignamos las pausas activas y almuerzos
            for empleado in empleados:
                if len(horarios[empleado]) < jornada_laboral:
                    if len(horarios[empleado]) >= minimo_trabajo_continuo:
                        if len(horarios[empleado]) >= maximo_trabajo_continuo or (franja_almuerzo_min <= i <= franja_almuerzo_max and 'Almuerza' not in horarios[empleado]):
                            if 'Almuerza' not in horarios[empleado] and franja_almuerzo_min <= i <= franja_almuerzo_max:
                                for _ in range(tiempo_almuerzo):
                                    if len(horarios[empleado]) < jornada_laboral:
                                        horarios[empleado].append('Almuerza')
                            else:
                                horarios[empleado].append('Pausa Activa')
                        else:
                            horarios[empleado].append('Trabaja')
                    else:
                        horarios[empleado].append('Trabaja')
                else:
                    horarios[empleado].append('Nada')

        # Verificación y corrección final
        for i in range(len(demanda)):
            if sum(1 for horario in horarios.values() if horario[i] == 'Trabaja') == 0:
                for empleado in empleados:
                    if horarios[empleado][i] == 'Nada':
                        horarios[empleado][i] = 'Trabaja'
                        break

    asignar_horarios(demanda, empleados)

    def graficar_barras_demanda(demanda, titulo, filename):
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
    
        plt.xlabel('Hora del día')
        plt.ylabel('Demanda')
        plt.title(titulo)
        plt.legend(['Demanda'], loc='upper right')
        plt.xticks(ticks=positions, labels=demanda['hora'], rotation=90, ha='right', fontsize=8)
        plt.tight_layout()
        
        # Eliminar archivo existente si existe
        full_filename = os.path.join(graficas_dir, filename)
        if os.path.exists(full_filename):
            os.remove(full_filename)
        
        plt.savefig(full_filename)
        plt.close()

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

        # Graficar la capacidad con un gráfico de escalera
        plt.step(positions, capacidad, label='Capacidad', color='blue', linewidth=3, where='mid')

        plt.xlabel('Hora del día')
        plt.ylabel('Demanda / Capacidad')
        plt.title(titulo)
        plt.legend()
        plt.xticks(ticks=positions, labels=demanda['hora'], rotation=90, ha='right', fontsize=8)
        plt.tight_layout()
        
        # Eliminar archivo existente si existe
        full_filename = os.path.join(graficas_dir, filename)
        if os.path.exists(full_filename):
            os.remove(full_filename)
        
        plt.savefig(full_filename)
        plt.close()

    def graficar_diferencia(demanda, capacidad, titulo, filename):
        diferencia = demanda['demanda'] - capacidad
        plt.figure(figsize=(10, 6))
        bar_width = 0.80
        positions = np.arange(len(demanda))

        plt.bar(positions, diferencia, label='Demanda - Capacidad', color=['red' if diff > 0 else 'green' for diff in diferencia], width=bar_width)
        plt.xlabel('Hora del día')
        plt.ylabel('Diferencia')
        plt.title(titulo)
        plt.axhline(0, color='black', linewidth=0.5)
        plt.legend(['Baja Capacidad', 'Sobre Capacidad'], loc='upper right')
        plt.xticks(ticks=positions, labels=demanda['hora'], rotation=90, ha='right', fontsize=8)
        plt.tight_layout()
        
        # Eliminar archivo existente si existe
        full_filename = os.path.join(graficas_dir, filename)
        if os.path.exists(full_filename):
            os.remove(full_filename)

        plt.savefig(os.path.join(graficas_dir, filename))
        plt.close()

    # Graficar resultados
    capacidad = np.zeros(len(demanda))
    for i in range(len(demanda)):
        capacidad[i] = sum(1 for horario in horarios.values() if horario[i] == 'Trabaja')

    graficar_barras_demanda(demanda, 'Demanda (Etapa 1)', 'Gpy1_etapa1.png')
    graficar_demanda_vs_capacidad(demanda, capacidad, 'Capacidad vs. Demanda (Etapa 1)', 'Gpy2_etapa1.png')
    graficar_diferencia(demanda, capacidad, 'Demanda - Capacidad (Etapa 1)', 'Gpy3_etapa1.png')

    # Guardar Resultados
    horarios_df = pd.DataFrame.from_dict(horarios, orient='index').transpose()
    horarios_df.index = demanda['fecha_hora']
    horarios_df.reset_index(inplace=True)
    horarios_df.rename(columns={'index': 'fecha_hora'}, inplace=True)
    
    # Ruta del archivo Excel
    excel_filename = os.path.join(base_dir, 'xlsx', 'horarios_optimizados_etapa1.xlsx')
    
    # Eliminar archivo existente si existe
    if os.path.exists(excel_filename):
        os.remove(excel_filename)
    

    horarios_df.to_excel(excel_filename, index=False)

    def agregar_imagenes_excel(excel_filename, graficas_dir):
        # Cargar el libro de Excel existente
        wb = load_workbook(excel_filename)
    
        # Definir las rutas de las imágenes y los títulos de las hojas
        img_paths = [
            ('Gpy1_etapa1.png', 'Demanda (Etapa 1)'),
            ('Gpy2_etapa1.png', 'Capacidad vs. Demanda (Etapa 1)'),
            ('Gpy3_etapa1.png', 'Demanda - Capacidad (Etapa 1)')
        ]
    
        # Agregar las imágenes a hojas nuevas
        for img_path, title in img_paths:
            ws = wb.create_sheet(title)
            img = Image(os.path.join(graficas_dir, img_path))
            ws.add_image(img, 'A1')
    
        # Guardar el libro de Excel con las imágenes agregadas
        wb.save(excel_filename)

    agregar_imagenes_excel(excel_filename, graficas_dir)

    if __name__ == "__main__":
        run_etapa1()