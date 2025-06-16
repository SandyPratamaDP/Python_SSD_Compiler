# Credit to : gibz104
# https://github.com/gibz104/xlsb-converter.git

import win32com.client as win32
import os
from dynamic_console import DynamicConsole # Import DynamicConsole Class

def xlsb_converter(file_path):
    xlApp = None
    try:
        xlApp = win32.Dispatch('Excel.Application')  # Create one Excel Application object
        xlApp.Visible = False  # Hide Excel window
        xlApp.ScreenUpdating = False  # Do not update the Excel window
        xlApp.DisplayAlerts = False  # Don't show alerts

        if os.path.splitext(file_path)[1].lower() in ['.xls', '.xlsx', '.xlsm', '.csv']:
            target_path = os.path.splitext(file_path)[0] + '.xlsb'  # Sets target path as .xlsb file extension
            
            wb = None # Workbook initialization per iteration
            try:
                wb = xlApp.Workbooks.Open(Filename=file_path, ReadOnly=True)  # Open files in read-only mode

                for ws in wb.Sheets: # Iterate through each worksheet in the workbook
                    ws.Columns.AutoFit()

                wb.SaveAs(Filename=target_path, FileFormat=50)  # Save file as xlsb format (FileFormat=50)
                DynamicConsole.print_message(f'Saved {target_path} from {file_path}')
            except Exception as e:
                DynamicConsole.print_message(f'Error processing {file_path}: {e}', 'error')  # Print an error if it cannot be processed
            finally:
                if wb is not None:
                    wb.Close(False)  # Tutup workbook tanpa menyimpan perubahan yang tidak diinginkan
        else:
            DynamicConsole.print_message(f"Skipping {file_path}: Not a supported Excel/CSV file type for conversion.", 'warning')
            
    except Exception as e:
        DynamicConsole.print_message(f"An error occurred with Excel application: {e}", 'error')
    finally:
        if xlApp is not None:
            xlApp.Quit()  # Shut down the Excel process only once after all is done or there is a fatal error
            del xlApp