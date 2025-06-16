import pandas as pd
import os
from dynamic_console import DynamicConsole # Impor kelas DynamicConsole
from xlsb_converter import xlsb_converter

def export_to_excel(df: pd.DataFrame, folder_path: str, output_file_name: str): # type: ignore
    """
    Export DataFrame to Excel file (.xlsx).
    """
    if df.empty:
        DynamicConsole.print_message("No data is processed for export to Excel.", "warning")
        return
    
    output_path = os.path.join(folder_path, output_file_name)

    try:
        # Save DataFrame to Excel file without indexes
        df.to_excel(output_path, index=False)
        
        # convert xlsx to xlsb
        xlsb_converter(output_path)

        DynamicConsole.print_message(f"The XLSX file was created successfully: {os.path.splitext(output_path)[0] + '.xlsb'}", "info")

        # Delete xlsx file after convert to xlsb
        try:
            os.remove(output_path)
        except FileNotFoundError:
            DynamicConsole.print_message(f"Error: File '{output_path}' was not found.", "error")
        except PermissionError:
            DynamicConsole.print_message(f"Error: No permission to delete file '{output_path}'.", "error")
        except Exception as e:
            DynamicConsole.print_message(f"An error occurred while deleting the file '{output_path}': {e}", "error")
    except Exception as e:
        DynamicConsole.print_message(f"Error saving Excel file: {e}", "error")
