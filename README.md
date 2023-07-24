# README

## Overview

This Python script is designed to perform a series of transformations on Excel files, specifically those containing real estate data. The script provides functionalities to:

- Load an Excel file selected by the user.
- Create a new folder named after the selected file and current date, where the log file and results will be stored.
- Check whether the Excel file has the required structure and content.
- Process the spreadsheet by performing operations like removing duplicates, calculating total square meterage, and calculating typology.
- Perform data validation checks.
- Log events and actions for debugging and tracking.
- Create a new processed Excel file with the results of the transformations, including the calculation of some statistical measures.

## Dependencies

This script requires the following Python libraries to be installed and available:

- pandas
- numpy
- openpyxl
- tkinter
- os
- math
- logging
- datetime

You can install any missing dependencies with pip:

```bash
pip install pandas numpy openpyxl tkinter os math logging datetime
```

## Usage

You can execute the script directly from a terminal:

```bash
python script_name.py
```

Replace "script_name.py" with the actual name of the script file.

Upon execution, the script will prompt the user to select an Excel file to process. The selected file is then validated and processed according to the designed transformations. The results are saved into a new Excel file, which is stored in a new directory named after the selected file and the current date. A log file is also generated and stored in the same directory.

## Code Structure

The script is organized around several functions, each designed to perform a specific task:

- `setup_process()`: Sets up the process by asking the user to select the Excel file to process, creates a new directory for results and logs, and sets up the logging configuration.
- `process_spreadsheet(selected_file)`: Main function that manages the processing of the Excel file.
- `calc_m2_totales(df)`: Calculates the total square meterage for a given DataFrame.
- `calc_stats(sheet, group)`: Calculates various statistics for a given group of data.

The script executes from the `__main__` clause, which triggers the setup process and then processes the selected file.

## Input Excel File Structure

The input Excel file should be structured in a specific way for the script to process it correctly. Specifically, it should contain the following columns:

- Estudio
- Habitaciones
- Baños
- Latitud
- Superficie
- Mts Total
- Mts Útil
- Mts Total Imp
- Mts Útil Imp
- Url Busconido
- Descripción
- F. Desactivación
- Precio ($)

## Output

The output of the script is an Excel file, named "procesado_" followed by the name of the original Excel file. It is saved in a new directory along with the log file. The processed file contains the result of the transformations performed on the original data, along with calculated statistics.

## Logging

The script includes detailed logging. The log messages are printed to the console and saved in a log file named "debug_" followed by the name of the original Excel file and the current date. The log file is stored in the same directory as the result Excel file.

## Note

This script has been designed with the assumption that the data in the Excel files adhere to a specific structure and format. It may not perform as expected if this assumption is not met. Please ensure the Excel files follow the required structure before executing the script.