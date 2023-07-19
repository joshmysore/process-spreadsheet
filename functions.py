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

# Define función para seleccionar el archivo
def setup_process():
    # Ask user for filename once and store it in a variable
    input('Presione Enter para seleccionar el archivo de Excel a procesar...')
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


    # Elimina la hoja por defecto
    del wb["Sheet"]

    # Guarda el archivo y dice la dirección
    processed_file_name = "procesado_" + os.path.basename(selected_file)
    wb.save(processed_file_name)
    #make this print statemnet a logging one
    logging.info(f'Procesamiento de datos finalizado. Resultados guardados como" {processed_file_name}" en {os.getcwd()}')

def validate_file_extension(selected_file):
    """Ensure the provided file has a valid Excel extension (.xls or .xlsx)."""
    if not selected_file.endswith((".xls", ".xlsx")):
        logging.info(f'"{selected_file}" no es un archivo de Excel.')
        return False
    return True

def file_is_open(selected_file):
    """Check if the provided file is open in another program."""
    try:
        with open(selected_file, "r") as f:
            pass
    except IOError:
        logging.info(f'"{selected_file}" está abierto en otro programa.')
        return True
    return False

def file_is_empty(selected_file):
    """Check if the provided Excel file is empty."""
    try:
        dfs = pd.read_excel(selected_file, sheet_name=None)
        first_sheet = list(dfs.keys())[0]
        if dfs[first_sheet].empty:
            logging.info(f'"{selected_file}" está vacío. Saltando este archivo.')
            return True
    except Exception as e:
        logging.info(
            f'No se puede leer "{selected_file}". Error: {str(e)}.'
        )
    return False

def validate_required_columns(df, required_columns):
    """Ensure the provided dataframe contains all the required columns."""
    missing_columns = [column for column in required_columns if column not in df.columns]
    if missing_columns:
        logging.info(f"Columns missing from dataframe: {', '.join(missing_columns)}")
        return False
    return True

def validate_data_values(df):
    """Verify that certain columns contain numeric values and the values fall within the correct range."""
    numeric_columns = ['Habitaciones', 'Baños']
    for column in numeric_columns:
        if not pd.api.types.is_numeric_dtype(df[column]):
            logging.info(f"Column {column} is not numeric.")
            return False
        if df[column].min() < 0:
            logging.info(f"Column {column} contains negative values.")
            return False
    return True

def create_typology(df):
    """Create the 'Tipología' column in the dataframe."""
    try:
        df['Tipología'] = df['Habitaciones'].astype(str) + ' habitaciones y ' + df['Baños'].astype(str) + ' baños'
    except Exception as e:
        logging.info(f"Failed to create 'Tipología' column. Error: {str(e)}")
        return False
    return True

def remove_duplicates(df):
    """Remove duplicate rows from the dataframe based on 'Latitud' column."""
    try:
        df.drop_duplicates(subset=['Latitud'], keep='first', inplace=True)
    except Exception as e:
        logging.info(f"Failed to remove duplicates. Error: {str(e)}")
        return False
    return True

def create_m2_totales_and_ranges(df):
    """Create 'm2 totales' and 'Rangos' columns in the dataframe."""
    try:
        df['m2 totales'] = df['Superficie cubierta (m2)'] + df['Superficie descubierta (m2)']
        df['Rangos'] = pd.cut(df['m2 totales'], bins=[0,50,100,200, np.inf], labels=['0-50','50-100','100-200', '200+'])
    except Exception as e:
        logging.info(f"Failed to create 'm2 totales' and 'Rangos' columns. Error: {str(e)}")
        return False
    return True

def remove_unnecessary_columns(df):
    """Remove unnecessary columns from the dataframe."""
    columns_to_remove = ['Column1', 'Column2', 'Column3']  # specify columns to remove here
    try:
        df.drop(columns_to_remove, axis=1, inplace=True)
    except Exception as e:
        logging.info(f"Failed to remove unnecessary columns. Error: {str(e)}")
        return False
    return True

def reorder_columns(df, price_index):
    """Reorder columns in the dataframe."""
    try:
        columns = list(df.columns)
        columns.insert(price_index, columns.pop(columns.index('Precio')))
        df = df[columns]
    except Exception as e:
        logging.info(f"Failed to reorder columns. Error: {str(e)}")
        return False
    return True

def calculate_and_add_ranges(df, grouped):
    """Calculate ranges and add it to the 'Rangos' column."""
    try:
        df['Rangos'] = df.groupby(grouped)['m2 totales'].apply(
            lambda x: pd.cut(x, bins=[0,50,100,200, np.inf], labels=['0-50','50-100','100-200', '200+'])
        )
    except Exception as e:
        logging.info(f"Failed to calculate and add ranges. Error: {str(e)}")
        return False
    return True

def create_sheets_for_typologies(wb, df, grouped):
    """Create new sheets for each typology."""
    try:
        for name, group in df.groupby(grouped):
            group.to_excel(wb, sheet_name=name, index=False)
    except Exception as e:
        logging.info(f"Failed to create sheets for each typology. Error: {str(e)}")
        return False
    return True

def style_and_adjust_sheet(wb, sheet, group):
    """Define styles for rows and adjust column width."""
    try:
        # Get the sheet from workbook
        ws = wb[sheet]

        # Define a font style
        font = Font(bold=True)
        for cell in ws[1]:  # assuming the first row needs to be bold
            cell.font = font

        # Adjust column widths
        for column in ws.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[(column[0].column)].width = adjusted_width

    except Exception as e:
        logging.info(f"Failed to style and adjust sheet. Error: {str(e)}")
        return False
    return True

def filter_rangos_column(sheet):
    """Apply a filter to the 'Rangos' column."""
    try:
        # Assuming 'Rangos' column is column C (3rd column), adjust as per your sheet
        sheet.auto_filter.ref = sheet.dimensions
        sheet.auto_filter.add_filter_column(2, ["0-50", "50-100", "100-200", "200+"])
    except Exception as e:
        logging.info(f"Failed to filter 'Rangos' column. Error: {str(e)}")
        return False
    return True

def save_workbook(wb, selected_file):
    """Save the workbook."""
    try:
        wb.save(filename = selected_file)
    except Exception as e:
        logging.info(f"Failed to save workbook. Error: {str(e)}")
        return False
    return True

# Define m2 totales función
def calc_m2_totales(df):
    columns = ["Superficie",
               "Mts Total",
               "Mts Útil",
               "Mts Total Imp",
               "Mts Útil Imp"]
    df["m2 totales"] = df[columns].max(axis=1)
    logging.info(f'Columnas {columns} procesadas para calcular m2 totales')
    return df["m2 totales"]

# Define función para calcular estadísticas
def calc_stats(sheet, group):
    # Define el número de filas de estadísticas
    row_num_stats = len(group) + 4
    logging.info(f'Calculando estadísticas para {len(group)} filas')

    # Define el estilo de las celdas de estadísticas
    bold_font = Font(bold=True)
    blue_fill = PatternFill(
        start_color="9ab7e6", end_color="9ab7e6", fill_type="solid"
    )
    logging.info(f'Estilo de celdas de estadísticas definido como bold_font y blue_fill')

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
    logging.info(f'Etiquetas de estadísticas definidas: {labels}')

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
