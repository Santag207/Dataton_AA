# Datatón Bancolombia 2023 - Solución de Programación Horaria

## Integrantes
- Santiago Castro Zuluaga

## Descripción
Este proyecto proporciona una solución para el reto de Datatón Bancolombia 2023. La tarea principal es optimizar la programación horaria de los empleados de caja en diferentes sucursales, tanto diaria como semanalmente, cumpliendo con diversas restricciones y tipos de contrato.

## Funcionalidades

### 1. Carga de Datos
El script carga los datos de demanda y los detalles de los empleados desde archivos Excel. Estos datos incluyen las franjas horarias, la demanda de empleados y la información sobre los trabajadores.

### 2. Análisis de Datos
Se realiza una descripción básica de los datos cargados para entender su estructura y valores.

### 3. Optimización de Horarios
El algoritmo de optimización asigna horarios a los empleados teniendo en cuenta las siguientes restricciones:
- Los empleados deben trabajar un mínimo de 1 hora y un máximo de 2 horas continuas antes de tomar una pausa o almuerzo.
- El almuerzo tiene una duración fija de 1.5 horas y solo es válido para empleados de tiempo completo de lunes a viernes.
- Las franjas horarias de almuerzo están restringidas entre las 11:30 AM y las 1:30 PM.
- Los empleados de tiempo completo trabajan 7 horas diarias de lunes a viernes y 5 horas los sábados. Los empleados de medio tiempo trabajan 4 horas diarias.
- La jornada laboral debe ser continua y debe finalizar con un estado de "Trabaja".

### 4. Visualización
El script genera gráficos para visualizar la demanda horaria, la capacidad de empleados asignados y la diferencia entre demanda y capacidad. Estos gráficos ayudan a validar que la asignación de empleados cumple con las restricciones y minimiza las diferencias.

### 5. Generación de Resultados
Los resultados de la programación horaria optimizada se guardan en un archivo Excel para su revisión y análisis adicional.

## Uso de Bibliotecas
El proyecto utiliza las siguientes bibliotecas:
- `pandas`: Para la manipulación y análisis de datos.
- `numpy`: Para operaciones numéricas y manejo de matrices.
- `matplotlib`: Para la generación de gráficos y visualización de datos.
- `openpyxl`: Para la manipulación de archivos Excel y la adición de gráficos en el Excel.

## Estructura del Código

### Etapa 1
- **Archivo:** `etapa1.py`
- **Función Principal:** `asignar_horarios_etapa1`
- **Descripción:**
  La función `asignar_horarios_etapa1` asigna horarios a los empleados teniendo en cuenta las restricciones de tiempo de trabajo continuo y pausas. 
  1. Inicializa un diccionario `horarios` para almacenar el estado de cada empleado en cada franja horaria.
  2. Itera a través de cada empleado y franja horaria, asignando `'Trabaja'`, `'Almuerza'` o `'Pausa Activa'` según las restricciones.
  3. Asegura que cada franja horaria tenga al menos un empleado trabajando.
- **Salida:** Un diccionario con los horarios asignados a cada empleado.

### Etapa 2
- **Archivo:** `etapa2.py`
- **Función Principal:** `asignar_horarios_etapa2`
- **Descripción:**
  La función `asignar_horarios_etapa2` asigna horarios a los empleados considerando diferentes tipos de jornadas laborales (completa, medio tiempo, sábado).
  1. Inicializa un diccionario `horarios` para almacenar el estado de cada empleado en cada franja horaria.
  2. Itera a través de cada día y franja horaria, asignando `'Trabaja'`, `'Almuerza'` o `'Pausa Activa'` según las restricciones.
  3. Asegura que cada franja horaria tenga al menos un empleado trabajando.
- **Salida:** Un diccionario con los horarios asignados a cada empleado.

### Ejecución Completa
- **Archivo:** `completo.py`
- **Función Principal:** `run_completo`
- **Descripción:**
  La función `run_completo` ejecuta las etapas 1 y 2, cargando los datos necesarios, llamando a las funciones de asignación de horarios, y generando los resultados.
  1. Carga los datos de demanda y empleados desde los archivos Excel.
  2. Ejecuta la función `asignar_horarios_etapa1` para la Etapa 1.
  3. Ejecuta la función `asignar_horarios_etapa2` para la Etapa 2.
  4. Genera gráficos de visualización y guarda los resultados en archivos Excel.

## Cómo Ejecutar el Proyecto

1. **Instalar Dependencias:**
   ```bash
   pip install pandas numpy matplotlib openpyxl
2. **Compilarlo desde CMD**
    ```bash
   python completo.py
3. Verificar resultados y solucion realizada para el reto
