# main.py
import functions

def main():
    # Prompt the user to select files
    selected_files = functions.select_files()
    for file_path in selected_files:
        # Validate each file
        is_valid = functions.read_and_validate_file(file_path)
        if is_valid:
            # Process each valid spreadsheet
            df = functions.process_spreadsheet(file_path)
            # Calculate m2 totales
            df = functions.calc_m2_totales(df)
            # Create a new workbook and styles
            wb, odd_row_style, even_row_style = functions.create_workbook_and_styles()
            # Process each sheet within a workbook
            functions.process_sheet(df, wb, odd_row_style, even_row_style)
            # Adjust the width of columns in the Excel sheet
            functions.adjust_column_widths(wb)
        else:
            print(f"File at {file_path} is not a valid Excel file.")

if __name__ == "__main__":
    main()
