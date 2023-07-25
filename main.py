from openpyxl import Workbook
from functions import *

def main():
    # Configurar el proceso y seleccionar el archivo a procesar
    folders, loggers, overall_loggers = setup_process()

    # Crear bucle para seleccionar el archivo a procesar
    for selected_file, folder_name in folders.items():
        # Seleccionar el logger correspondiente al archivo seleccionado
        logger = loggers[selected_file]
        # Seleccionar el logger general correspondiente al archivo seleccionado
        overall_logger = overall_loggers[selected_file]

        # Verificar la validez del archivo seleccionado
        if not check_file_validity(selected_file, logger):
            print(f'El archivo {selected_file} no es válido.')
            logging.info(f'El archivo {selected_file} no es válido.')
            continue

        # Leer y preprocesar el archivo en un DataFrame
        dfs = read_and_preprocess_file(selected_file, logger)

        # Crear un nuevo libro de Excel
        wb = Workbook()

        # Eliminar la hoja predeterminada 'Sheet'
        if 'Sheet' in wb.sheetnames:
            del wb['Sheet']

        # Crear y aplicar los estilos para el libro de trabajo
        odd_row_style, even_row_style = create_and_apply_styles(wb)

        # Cambiar el directorio de trabajo al directorio de la carpeta
        os.chdir(folder_name)

        # Procesar cada hoja en el DataFrame y modificar el libro de trabajo
        wb = process_sheet(dfs, wb, odd_row_style, even_row_style, logger)

        # Guardar el libro de trabajo
        save_workbook(wb, selected_file, logger)

        # Cambiar el directorio de trabajo al directorio principal
        os.chdir("..")

        # Cerrar los handlers del logger para el archivo procesado
        for handler in logger.handlers:
            handler.close()
            
        # Cerrar los handlers del logger general para el archivo procesado
        for handler in overall_logger.handlers:
            handler.close()

    # Cerrar los handlers del logger general para todos los archivos
    for selected_file, overall_logger in overall_loggers.items():
        overall_logger.info(f'Archivo {selected_file} procesado.')

if __name__ == "__main__":
    main()