import pandas as pd
import os
from excel_processor.xlsb_converter import xlsb_converter

def export_to_excel(df: pd.DataFrame, folder_path: str, output_file_name: str, console_instance): # type: ignore
    """
    Export DataFrame to Excel file (.xlsx).
    """
    if df.empty:
        console_instance.print_message("No data is processed for export to Excel.", "warning")
        return
    
    output_path = os.path.join(folder_path, output_file_name)

    try:
        # Save DataFrame to Excel file without indexes
        df.to_excel(output_path, index=False)
        console_instance.print_message(f"The XLSX file was created successfully: {output_path}", "success")
        
        # convert xlsx to xlsb
        if xlsb_converter(output_path, console_instance):
             # Delete xlsx file after convert to xlsb
            try:
                os.remove(output_path)
            except FileNotFoundError:
                console_instance.print_message(f"Error: File '{output_path}' was not found.", "error")
            except PermissionError:
                console_instance.print_message(f"Error: No permission to delete file '{output_path}'.", "error")
            except Exception as e:
                console_instance.print_message(f"An error occurred while deleting the file '{output_path}': {e}", "error")

    except Exception as e:
        console_instance.print_message(f"Error saving Excel file: {e}", "error")
