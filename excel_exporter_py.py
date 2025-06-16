import pandas as pd
import os
from dynamic_console import DynamicConsole # Impor kelas DynamicConsole

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
        DynamicConsole.print_message(f"The XLSX file was created successfully: {output_path}", "info")
    except Exception as e:
        DynamicConsole.print_message(f"Error saving Excel file: {e}", "error")
