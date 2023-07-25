# README

## Descripción general

Este proyecto es una herramienta de procesamiento de hojas de cálculo de Excel escrita en Python que permite procesar archivos de Excel de Busconido, realizar varias validaciones y transformaciones de datos, y luego guardar los resultados en un nuevo archivo de Excel. El proyecto se compone de dos scripts principales, `main.py` y `functions.py`.

Busconido es un portal de arriendo en Chile que aporta datos actuales del mercado sobre propiedades en todo el país. La idea fue sacar los datos de esta fuente y procesarlos para obtener información relevante para el negocio.

El script `main.py` es el punto de entrada del programa y utiliza funciones definidas en el script `functions.py` para realizar el procesamiento. `functions.py` contiene una serie de funciones auxiliares que se utilizan para tareas específicas en el proceso de transformación de datos.

## `main.py`

`main.py` consta de una función principal `main()` que se ejecuta cuando el script se ejecuta como un programa independiente. Aquí está la descripción detallada del flujo del programa:

1. **Configuración del proceso y selección del archivo a procesar:** El programa primero llama a la función `setup_process()` para solicitar al usuario que seleccione el archivo de Excel a procesar. Esta función devuelve la ruta del archivo seleccionado. También se instala el registro del código durante esta parte.

2. **Verificación de la validez del archivo seleccionado:** A continuación, el programa llama a la función `check_file_validity(selected_file)` para verificar si el archivo seleccionado es válido y se puede abrir para leer.

3. **Lectura y preprocesamiento del archivo:** El archivo de Excel seleccionado se lee en un diccionario de DataFrames de pandas, donde cada clave del diccionario es el nombre de una hoja de cálculo en el archivo de Excel, y el valor correspondiente es un DataFrame de pandas que contiene los datos de esa hoja de cálculo.

4. **Creación de un nuevo libro de Excel:** Se crea un nuevo libro de trabajo de Excel utilizando la clase `Workbook` del paquete `openpyxl`. También se elimina la hoja de cálculo predeterminada llamada 'Sheet' si existe.

5. **Creación y aplicación de estilos para el libro de trabajo:** Se crean y aplican estilos personalizados para las filas pares e impares de las hojas de cálculo utilizando la función `create_and_apply_styles(wb)`.

6. **Procesamiento de cada hoja en el DataFrame y modificación del libro de trabajo:** La función `process_sheet(dfs, wb, odd_row_style, even_row_style)` se llama para procesar cada hoja en el DataFrame y modificar el libro de trabajo.

7. **Guardar el libro de trabajo:** Finalmente, el libro de trabajo modificado se guarda como un nuevo archivo de Excel utilizando la función `save_workbook(wb, selected_file)`.

## `functions.py`

El script `functions.py` consta de una serie de funciones auxiliares que se utilizan para realizar varias tareas en el proceso de transformación de datos. Aquí está la descripción detallada de cada función:

1. `setup_process()`: Pide al usuario que seleccione el archivo de Excel a procesar, crea una carpeta para guardar los resultados y configura el registro de eventos.

2. `check_file_validity(selected_file)`: Verifica si el archivo seleccionado es válido y se puede abrir para leer.

3. `create_and_apply_styles(wb)`: Crea y aplica estilos personalizados para las filas pares e impares de las hojas de cálculo.

4. `calc_m2_totales(df)`: Calcula los metros cuadrados totales según varias columnas en el DataFrame.
 
5. `calc_stats(sheet, group)`: Calcula estadísticas para cada grupo de datos. Las estadísticas calculadas incluyen el promedio, la moda, la mediana, el mínimo, el máximo y los percentiles de 80, 85, 90 y 95. También se crea un cuadro de resumen que contiene las estadísticas calculadas en la parte inferior de la hoja de cálculo. 

6. `process_sheet(dfs, wb, odd_row_style, even_row_style)`: Procesa cada hoja en el DataFrame, realiza varias transformaciones de datos y modifica el libro de trabajo.

7. `save_workbook(wb, selected_file)`: Guarda el libro de trabajo modificado como un nuevo archivo de Excel.

## Uso

Para ejecutar el script, asegúrese de tener Python instalado en su sistema, así como las bibliotecas necesarias (`pandas`, `numpy`, `openpyxl`, `tkinter`, `logging`, `os`, `math`, `datetime`). Luego, puede ejecutar el script `main.py` utilizando el intérprete de Python:

```
python main.py
```

Se le pedirá que seleccione el archivo de Excel a procesar. El programa luego procesará el archivo, realizará varias transformaciones de datos y guardará los resultados en un nuevo archivo de Excel en la misma carpeta que el archivo original. El archivo de Excel tiene que venir de Busconido, tener una sola hoja y contar con los siguientes columnas:

- `Estudio`
- `Habitaciones`
- `Baños`
- `Latitud`
- `Superficie`
- `Mts Total`
- `Mts Útil`
- `Mts Total Imp`
- `Mts Útil Imp`
- `Url Busconido`
- `Descripción`
- `F. Desactivación`
- `Precio ($)`


También hay la posibilidad de correr el archivo ejecutable, que se encuentra en la carpeta `executable_files`. Para ejecutar el archivo ejecutable, simplemente haga doble clic en él. Se le pedirá que seleccione el archivo de Excel a procesar. El programa luego procesará el archivo, realizará varias transformaciones de datos y guardará los resultados en un nuevo archivo de Excel en una nueva carpeta en la misma carpeta que el archivo ejecutable.

## Dependencias

Este proyecto requiere las siguientes bibliotecas de Python:

- `pandas`
- `numpy`
- `openpyxl`
- `tkinter`
- `logging`
- `os`
- `math`
- `datetime`

```
pip install pandas numpy openpyxl tkinter logging os math datetime
```

## Proceso de transformación de datos

El proceso de transformación de datos en `process_sheet()` realiza varias operaciones en el DataFrame para cada hoja de cálculo en el archivo de Excel. Aquí hay una descripción detallada de lo que sucede en esta función:

1. **Reemplazo de valores NaN:** Los valores NaN en ciertas columnas se reemplazan con valores por defecto.

2. **Generación de nuevas columnas:** Se generan nuevas columnas basadas en ciertos cálculos y lógicas.

3. **Aplicación de filtros:** Se aplican filtros a ciertas columnas basados en ciertos criterios.

4. **Ordenamiento de columnas:** Las columnas se ordenan en un orden específico.

5. **Formato de celdas:** Se aplica el formato adecuado a las celdas de ciertas columnas.

6. **Aplicación de estilos:** Se aplican estilos personalizados a las filas y columnas.

Al final de estas transformaciones, los datos de cada hoja de cálculo se escriben en una nueva hoja en el nuevo libro de trabajo.

## Registro de eventos

Durante el proceso de transformación de datos, se registra una serie de eventos para ayudar a rastrear el progreso del programa y diagnosticar cualquier problema que pueda surgir. El registro de eventos se configura en la función `setup_process()`, y los mensajes de registro se generan en varias partes del código utilizando la biblioteca `logging` de Python.

Los mensajes de registro incluyen información sobre:

- Inicio y fin del proceso de transformación de datos.
- Cualquier error o excepción que ocurra durante el proceso.
- Información detallada sobre el progreso del programa, como la cantidad de datos procesados, las operaciones realizadas, etc.

Los mensajes de registro se escriben en la consola y también se guardan en un archivo de registro para referencia futura.

## Personalización

Puede personalizar este script para adaptarlo a sus necesidades específicas. Algunas formas en las que puede personalizar este script incluyen:

- Modificar las funciones de transformación de datos en `process_sheet()` para realizar diferentes operaciones en los datos.
- Ajustar los estilos de las filas y columnas en `create_and_apply_styles()`.
- Cambiar la lógica de validación del archivo en `check_file_validity()`.

Recuerde que cualquier cambio en el script debe ser probado adecuadamente para asegurarse de que el programa sigue funcionando correctamente.

## Contribuciones

Las contribuciones a este proyecto son bienvenidas. Si encuentra un error o tiene una sugerencia para una nueva característica, por favor abra un nuevo issue. Si desea contribuir con código, por favor abra una solicitud de pull.

## Licencia

Este proyecto está licenciado bajo los términos de la licencia de CCLA.

## Contacto

Este programa fue creado por Josh Mysore (joshmysore@college.harvard.edu), un estudiante de Havard, durante una pasantía con Compass en Julio 2023. Si tiene alguna pregunta o comentario sobre este proyecto, por favor póngase en contacto con el mantenedor actual del proyecto. 