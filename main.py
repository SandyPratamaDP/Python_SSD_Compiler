import os
import argparse
from datetime import datetime

# Import functions and classes from other modules
from dynamic_console import DynamicConsole
from csv_parser import parse_csv_sections
from data_processor import extract_coil_no_from_filename, process_record
from excel_exporter_py import export_to_excel
import pandas as pd # Required for pd.DataFrame

def main():
    parser = argparse.ArgumentParser(description="Console application to convert export results from the SDD application from CSV files to XLSB.")
    parser.add_argument("--path", required=True, help="Path to the folder containing the CSV file.")
    parser.add_argument("--startDate", required=True, help="Start date (YYYY-MM-DD or HH:mm) for filtering CSV files.")
    parser.add_argument("--endDate", required=True, help="The end date (YYYY-MM-DD or HH:mm) to filter the CSV file.")
    # Optional argument for the name of the Excel output file
    parser.add_argument("--outputFileName", required=False, help="Custom Excel output file name (without extension). Default: CompiledData_{timestamp}")

    args = parser.parse_args()

    folder_path = args.path
    start_date_str = args.startDate
    end_date_str = args.endDate
    custom_output_filename_base = args.outputFileName

    date_formats = ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"]

    # Parsing start_date
    start_date = None
    for fmt in date_formats:
        try:
            start_date = datetime.strptime(start_date_str, fmt)
            break
        except ValueError:
            continue
    
    # Parsing end_date
    end_date = None
    for fmt in date_formats:
        try:
            end_date = datetime.strptime(end_date_str, fmt)
            # If only a date is given for end_date, set it to the end of the day
            if ' ' not in end_date_str or len(end_date_str.split(' '))[0] == 10: # type: ignore
                end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)
            break
        except ValueError:
            continue

    if start_date is None or end_date is None:
        DynamicConsole.print_message("Error: Invalid date format. UseYYYY-MM-DD or YYYY-MM-DD HH:mm.", "error")
        return

    DynamicConsole.print_message(
        f"Search for CSV files in '{folder_path}' from '{start_date}' to '{end_date}'.", "info"
    )

    if not os.path.isdir(folder_path):
        DynamicConsole.print_message(f"Error: The folder '{folder_path}' was not found.", "error")
        return

    if start_date > end_date:
        DynamicConsole.print_message("Warning: The start date is greater than the end date. Reverse the order.", "warning")
        start_date, end_date = end_date, start_date

    csv_files = []
    for f_name in os.listdir(folder_path):
        if f_name.lower().endswith(".csv"):
            f_path = os.path.join(folder_path, f_name)
            try:
                file_mtime = datetime.fromtimestamp(os.path.getmtime(f_path))
                if start_date <= file_mtime <= end_date:
                    csv_files.append(f_path)
            except Exception as e:
                DynamicConsole.print_message(f"Warning: Failed to get the modification time of file '{f_name}': {e}", "warning")

    if not csv_files:
        DynamicConsole.print_message(f"No CSV files were found in '{folder_path}' in that date range.", "warning")
        return

    all_processed_records_for_excel = []

    for file_path in csv_files:
        DynamicConsole.print_message(f"Reading file: {os.path.basename(file_path)}", "info")
        
        coil_no = extract_coil_no_from_filename(os.path.basename(file_path))
        top_df, bottom_df = parse_csv_sections(file_path)

        for index, row_series in top_df.iterrows():
            processed_rec = process_record(row_series, coil_no, "Top")
            if processed_rec:
                all_processed_records_for_excel.append(processed_rec)
        
        for index, row_series in bottom_df.iterrows():
            processed_rec = process_record(row_series, coil_no, "Bottom")
            if processed_rec:
                all_processed_records_for_excel.append(processed_rec)

    # Convert list of dictionaries to DataFrame for easier Excel exports
    final_df = pd.DataFrame(all_processed_records_for_excel)

    # Define the proper column order for Excel output
    excel_output_columns = [
        "Coil No", "Class Name", "Defect Name", "Grade Defect", "Top/Bottom",
        "Distance from HE CGL (m)", "Distance Left (mm)", "Distance Right (mm)",
        "Distance Center (mm)", "Height", "Width", "Segment Width Ratio", "Orientation"
    ]
    
    # Rearrange columns and fill missing values with None (NaN in pandas)
    final_df = final_df.reindex(columns=excel_output_columns, fill_value=None)

    # Excel output file name determination logic
    if custom_output_filename_base:
        output_file_name_base = custom_output_filename_base
    else:
        output_file_name_base = "CompiledData"
    
    current_timestamp_for_filename = datetime.now().strftime('%Y%m%d%H%M%S')
    # Forcefully add .xlsx extension
    output_full_filename = f"{output_file_name_base}_{current_timestamp_for_filename}.xlsx"
    
    export_to_excel(final_df, folder_path, output_full_filename) # type: ignore

    DynamicConsole.print_message(f"\nFinish processing the CSV files.", "info")

if __name__ == "__main__":
    main()
