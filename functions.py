import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
import os
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, NamedStyle, Side, Border
import math as math
import logging
from datetime import datetime

def setup_process():
    # Ask user for filename once and store it in a variable
    root = tk.Tk()
    root.withdraw()
    selected_file = filedialog.askopenfilename()
    selected_file_name = os.path.basename(selected_file)
    # Get current date
    now = datetime.now()
    # Format date as string
    current_date = now.strftime("%Y_%d_%m")
    folder_name = 'resultados_{0}_{1}'.format(selected_file_name, current_date)

    # make folder
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
        logging.info('Carpeta creada en {0}'.format(os.getcwd()))

    # config para logging
    logging.basicConfig(
        level=logging.INFO, 
        format='%(asctime)s - %(levelname)s - %(message)s', 
        handlers=[
            # name log after the file selected as such debug_{filename of selected file}_time.log and put it in folder_name
            logging.FileHandler("{0}/debug_{1}_{2}.log".format(folder_name, selected_file_name, current_date)),
            logging.StreamHandler()
        ]
    )
    # statement saying log file has been created and saved whereever it is found
    logging.info('Archivo de log creado en {0}'.format(os.getcwd()))

    # change the working directory to the new folder
    os.chdir(folder_name)

    logging.info('Empezando el programa')

    return selected_file

import math
import os
import logging
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, PatternFill, Border, Side, Font
from openpyxl.utils.dataframe import dataframe_to_rows


def create_workbook():
    # Define estilo para las celdas de cabecera
    odd_row_fill = PatternFill(
        start_color="D3D3D3", end_color="D3D3D3", fill_type="solid"
    )
    odd_row_style = NamedStyle(name="odd_row_style", fill=odd_row_fill)
    even_row_fill = PatternFill(
        start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
    )
    even_row_style = NamedStyle(name="even_row_style", fill=even_row_fill)

    # Crea un nuevo archivo de Excel
    wb = Workbook()

    # Añade el estilo a los estilos del archivo
    wb.add_named_style(odd_row_style)
    wb.add_named_style(even_row_style)

    return wb, odd_row_style, even_row_style


def validate_excel_file(selected_file):
    logging.info(f'Procesando el archivo {selected_file}')
    # Compueba que el archivo seleccionado es un archivo de Excel
    if not selected_file.endswith((".xls", ".xlsx")):
        logging.info(f'"{selected_file}" no es un archivo de Excel.')
        return False

    try:
        # Comprueba que el archivo no está abierto en otro programa
        with open(selected_file, "r") as f:
            pass
    except IOError:
        logging.info(f'"{selected_file}" está abierto en otro programa.')
        return False

    try:
        # Comprueba que el archivo no está vacío
        dfs = pd.read_excel(selected_file, sheet_name=None)
        first_sheet = list(dfs.keys())[0]
        if dfs[first_sheet].empty:
            logging.info(f'"{selected_file}" está vacío. Saltando este archivo.')
            return False
    except Exception as e:
        logging.info(
            f'No se puede leer "{selected_file}". Error: {str(e)}.'
        )
        return False

    logging.info(f'Archivo {selected_file} ha pasado los checks de ser un archivo de Excel con contenido.')
    return True


def validate_required_columns(dfs, selected_file):
    # Comprueba que el archivo contiene todas las columnas necesarias
    required_columns = [
        "Estudio",
        "Habitaciones",
        "Baños",
        "Latitud",
        "Superficie",
        "Mts Total",
        "Mts Útil",
        "Mts Total Imp",
        "Mts Útil Imp",
        "Url Busconido",
        "Descripción",
        "F. Desactivación",
        "Precio ($)",
    ]

    for sheet_name, df in dfs.items():
        if not set(required_columns).issubset(df.columns):
            print(
                f'Hoja "{sheet_name}" en "{selected_file}" no contiene las columnas requeridas. Saltando este archivo.'
            )
            dfs.pop(sheet_name)


def process_dataframes(dfs, selected_file):
    # Continue with the dataframe processing based on the validated sheets
    for sheet_name, df in dfs.items():
        # Compueba que las columnas "Habitaciones" y "Baños" contienen valores numéricos
        if df["Habitaciones"].dtype not in ["int64", "float64"] or df[
            "Baños"
        ].dtype not in ["int64", "float64"]:
            print(
                f'Hoja "{sheet_name}" en "{selected_file}" contiene los datos incorrectos para  "Habitaciones" y/o "Baños". Saltando este archivo.'
            )
            dfs.pop(sheet_name)
            continue
        df["Habitaciones"] = (df["Habitaciones"] // 1).fillna("NaN")
        df["Baños"] = (df["Baños"] // 1).fillna("NaN")
        if df["Habitaciones"].min() < 1 or df["Baños"].min() < 1:
            print(
                f'Hoja "{sheet_name}" en "{selected_file}" contiene los datos incorrectos para  "Habitaciones" y/o "Baños". Saltando este archivo.'
            )
            dfs.pop(sheet_name)
            continue
    return dfs


def calculate_ranges(df):
    df["m2 totales"] = calc_m2_totales(df)

    # Group the dataframe by 'Tipología'
    grouped = df.groupby('Tipología')

    # For each typology, calculate statistics
    for typology, group in grouped:
        # Calculate the min and max of m2 totals
        min_m2 = min(group["m2 totales"])
        max_m2 = max(group["m2 totales"])

        # If the min and max are the same, there is only one range
        if min_m2 == max_m2:
            num_ranges = 1
        else:
            # Calculate the number of ranges
            num_ranges = math.ceil((max_m2 - min_m2) / 10)

        # Create a list of tuples with the ranges
        filter_ranges = [(min_m2 + i * 10, min_m2 + (i + 1) * 10) for i in range(num_ranges)]

        # Add the range to the Rangos column
        df.loc[group.index, "Rangos"] = pd.cut(
            group["m2 totales"],
            bins=[range[0] for range in filter_ranges] + [max_m2 + 1],
            labels=[f"{range[0]}-{range[1]}" for range in filter_ranges],
            include_lowest=True,
        )
        
    return df

def save_to_excel(wb, df, odd_row_style, even_row_style):
    # Add DataFrame to Worksheet
    ws = wb.active
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # Add Style to Rows
    for row in ws.iter_rows(min_row=2):
        if row[0].row % 2 == 0:
            for cell in row:
                cell.style = even_row_style
        else:
            for cell in row:
                cell.style = odd_row_style

    # Save Workbook
    wb.save(filename="output.xlsx")
    



# Define m2 totales función
def calc_m2_totales(df):
    columns = ["Superficie",
               "Mts Total",
               "Mts Útil",
               "Mts Total Imp",
               "Mts Útil Imp"]
    df["m2 totales"] = df[columns].max(axis=1)
    return df["m2 totales"]

def calc_stats(sheet, group):

    # Define el número de filas de estadísticas
    row_num_stats = len(group) + 4

    # Define el estilo de las celdas de estadísticas
    bold_font = Font(bold=True)
    blue_fill = PatternFill(
        start_color="9ab7e6", end_color="9ab7e6", fill_type="solid"
    )

    # Define las etiquetas de las estadísticas
    labels = [
        "Promedio:",
        "Moda:",
        "Mediana:",
        "Rango Mínimo:",
        "Rango Máximo:",
        "Percentil 80:",
        "Percentil 85:",
        "Percentil 90:",
        "Percentil 95:",
    ]

    # Define las fórmulas de las estadísticas
    for i in range(len(labels) + 1):
        if i == 0:  # For the first row
            sheet.cell(
                row=row_num_stats, column=3, value="N. Precio ($)"
            ).font = bold_font
            sheet.cell(
                row=row_num_stats, column=4, value="N. m2 totales"
            ).font = bold_font
            sheet.cell(row=row_num_stats, column=3).fill = blue_fill
            sheet.cell(row=row_num_stats, column=4).fill = blue_fill
        else:  # Para el resto de filas
            sheet.cell(
                row=row_num_stats + i, column=2, value=labels[i - 1]
            ).font = bold_font
            sheet.cell(row=row_num_stats + i, column=2).fill = blue_fill

    # Busca las columnas de Precio ($) y m2 totales
    precio_column = None
    m2_column = None
    for col, cell in enumerate(sheet[1], start=1):
        if cell.value == "Precio ($)":
            precio_column = col
        elif cell.value == "m2 totales":
            m2_column = col

    if precio_column is None or m2_column is None:
        print(f"Columns not found for 'Precio ($)' and 'm2 totales'")
        return

    # Define las funciones de estadísticas
    functions = {
        1: [None],  # PROMEDIO
        13: [None],  # MODA
        12: [None],  # MEDIANA
        5: [None],  # MIN
        4: [None],  # MAX
        18: ["0.8", "0.85", "0.9", "0.95"],  # PERCENTIL
    }

    # Bucle para calcular las estadísticas
    i = 0
    for function, args in functions.items():
        for arg in args:
            # Hazlo con argumentos y sin argumentos
            if arg is None:
                formula_precio = f"=AGREGAR({function}, 5, {sheet.cell(row=2, column=precio_column).column_letter}2:{sheet.cell(row=row_num_stats - 1, column=precio_column).column_letter}{row_num_stats - 1})"
                formula_m2 = f"=AGREGAR({function}, 5, {sheet.cell(row=2, column=m2_column).column_letter}2:{sheet.cell(row=row_num_stats - 1, column=m2_column).column_letter}{row_num_stats - 1})"
            else:
                formula_precio = f"=AGREGAR({function}, 5, {sheet.cell(row=2, column=precio_column).column_letter}2:{sheet.cell(row=row_num_stats - 1, column=precio_column).column_letter}{row_num_stats - 1}, {arg})"
                formula_m2 = f"=AGREGAR({function}, 5, {sheet.cell(row=2, column=m2_column).column_letter}2:{sheet.cell(row=row_num_stats - 1, column=m2_column).column_letter}{row_num_stats - 1}, {arg})"

            sheet.cell(
                row=row_num_stats + 1 + i, column=3, value=formula_precio
            ).number_format = "$#,##0.00"
            sheet.cell(
                row=row_num_stats + 1 + i, column=4, value=formula_m2
            ).number_format = "#,##0.00"
            i += 1

if __name__ == "__main__":
    selected_file = setup_process()
    process_spreadsheet(selected_file)