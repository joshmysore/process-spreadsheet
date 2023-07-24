from openpyxl import Workbook
from functions import *

def main():
    # Configurar el proceso y seleccionar el archivo a procesar
    selected_file = setup_process()

    # Verificar la validez del archivo seleccionado
    check_file_validity(selected_file)

    # Leer y preprocesar el archivo en un DataFrame
    dfs = pd.read_excel(selected_file, sheet_name=None)

    # Crear un nuevo libro de Excel
    wb = Workbook()

    # Eliminar la hoja predeterminada 'Sheet'
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # Crear y aplicar los estilos para el libro de trabajo
    odd_row_style, even_row_style = create_and_apply_styles(wb)

    # Procesar cada hoja en el DataFrame y modificar el libro de trabajo
    wb = process_sheet(dfs, wb, odd_row_style, even_row_style)

    # Guardar el libro de trabajo
    save_workbook(wb, selected_file)

if __name__ == "__main__":
    main()