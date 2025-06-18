import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkcalendar import DateEntry

# Import functions and classes from other modules
# Assuming these modules are in the same directory or accessible via PYTHONPATH
from console.dynamic_console_gui import DynamicConsoleGUI # Modified to work with GUI
from csv_parser import parse_csv_sections
from data_processor import extract_coil_no_from_filename, process_record
from excel_processor.excel_exporter import export_to_excel
import pandas as pd # Required for pd.DataFrame

class SDDConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("SDD Compiler")
        master.geometry("800x600") # Set initial window size

        icon_path = os.path.join(os.path.dirname(__file__), 'favicon.ico')
        try:
            if os.path.exists(icon_path):
                master.iconbitmap(icon_path) # Ini mengatur ikon jendela dan taskbar
            else:
                print(f"Warning: Icon file not found at {icon_path}")
        except tk.TclError as e:
            print(f"Error setting icon: {e}. Make sure the .ico file is valid.")

        self.console_output = DynamicConsoleGUI(master)

        # --- Input Frame ---
        input_frame = tk.LabelFrame(master, text="Input Parameters", padx=10, pady=10)
        input_frame.pack(pady=10, padx=10, fill="x")

        # Folder Path
        tk.Label(input_frame, text="CSV Folder Path:").grid(row=0, column=0, sticky="w", pady=2)
        self.path_entry = tk.Entry(input_frame, width=60)
        self.path_entry.grid(row=0, column=1, sticky="ew", pady=2)
        tk.Button(input_frame, text="Browse", command=self.browse_folder).grid(row=0, column=2, padx=5, pady=2)

        # Start Date
        tk.Label(input_frame, text="Start Date:").grid(row=1, column=0, sticky="w", pady=2)
        self.start_date_entry = DateEntry(input_frame, selectmode='day', date_pattern='yyyy-mm-dd')
        self.start_date_entry.grid(row=1, column=1, sticky="ew", pady=2)

        # End Date
        tk.Label(input_frame, text="End Date:").grid(row=2, column=0, sticky="w", pady=2)
        self.end_date_entry = DateEntry(input_frame, selectmode='day', date_pattern='yyyy-mm-dd')
        self.end_date_entry.grid(row=2, column=1, sticky="ew", pady=2)
        # Set end date to current date by default for convenience
        self.end_date_entry.set_date(datetime.now())

        # Output File Name
        tk.Label(input_frame, text="Output File Name (Optional):").grid(row=3, column=0, sticky="w", pady=2)
        self.output_filename_entry = tk.Entry(input_frame, width=60)
        self.output_filename_entry.grid(row=3, column=1, sticky="ew", pady=2)
        self.output_filename_entry.insert(0, "CompiledData") # Default value

        # Configure column weights for resizing
        input_frame.grid_columnconfigure(1, weight=1)

        # --- Run Button ---
        self.run_button = tk.Button(master, text="Run Conversion", command=self.run_conversion, height=2)
        self.run_button.pack(pady=10, padx=10, fill="x")

        # --- Console Output Area ---
        console_frame = tk.LabelFrame(master, text="Process Log", padx=10, pady=5)
        console_frame.pack(pady=5, padx=10, fill="both", expand=True)

        self.log_text = scrolledtext.ScrolledText(console_frame, width=80, height=15, state='disabled', wrap=tk.WORD)
        self.log_text.pack(fill="both", expand=True)
        self.console_output.set_text_widget(self.log_text) # Link DynamicConsole to this widget

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, folder_selected)

    def run_conversion(self):
        folder_path = self.path_entry.get()
        start_date_str = self.start_date_entry.get_date().strftime('%Y-%m-%d')
        end_date_str = self.end_date_entry.get_date().strftime('%Y-%m-%d')
        custom_output_filename_base = self.output_filename_entry.get()

        # Clear previous log messages
        self.console_output.clear_log()

        # Basic validation
        if not folder_path:
            self.console_output.print_message("Error: Please select a folder path.", "error")
            return

        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            # If only a date is given for end_date, set it to the end of the day
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d").replace(hour=23, minute=59, second=59, microsecond=999999)
        except ValueError:
            self.console_output.print_message("Error: Invalid date format. Use YYYY-MM-DD.", "error")
            return

        if not os.path.isdir(folder_path):
            self.console_output.print_message(f"Error: The folder '{folder_path}' was not found.", "error")
            return

        if start_date > end_date:
            self.console_output.print_message("Warning: The start date is greater than the end date. Reversing order for processing.", "warning")
            start_date, end_date = end_date, start_date

        self.console_output.print_message(
            f"Search for CSV files in '{folder_path}' from '{start_date.strftime('%Y-%m-%d')}' to '{end_date.strftime('%Y-%m-%d')}'.", "info"
        )
        self.master.update_idletasks() # Update GUI to show messages

        csv_files = []
        for f_name in os.listdir(folder_path):
            if f_name.lower().endswith(".csv"):
                f_path = os.path.join(folder_path, f_name)
                try:
                    file_mtime = datetime.fromtimestamp(os.path.getmtime(f_path))
                    if start_date <= file_mtime <= end_date:
                        csv_files.append(f_path)
                except Exception as e:
                    self.console_output.print_message(f"Warning: Failed to get the modification time of file '{f_name}': {e}", "warning")

        if not csv_files:
            self.console_output.print_message(f"No CSV files were found in '{folder_path}' in that date range.", "warning")
            self.console_output.print_message(f"Please check the folder path and date range.", "info")
            return

        all_processed_records_for_excel = []

        for file_path in csv_files:
            self.console_output.print_message(f"Reading file: {os.path.basename(file_path)}", "info")
            self.master.update_idletasks() # Update GUI

            coil_no = extract_coil_no_from_filename(os.path.basename(file_path), self.console_output)
            try:
                top_df, bottom_df = parse_csv_sections(file_path, self.console_output)

                for index, row_series in top_df.iterrows():
                    processed_rec = process_record(row_series, coil_no, "Top")
                    if processed_rec:
                        all_processed_records_for_excel.append(processed_rec)

                for index, row_series in bottom_df.iterrows():
                    processed_rec = process_record(row_series, coil_no, "Bottom")
                    if processed_rec:
                        all_processed_records_for_excel.append(processed_rec)
            except Exception as e:
                self.console_output.print_message(f"Error processing file '{os.path.basename(file_path)}': {e}", "error")
                continue


        if not all_processed_records_for_excel:
            self.console_output.print_message("No data was processed from the selected CSV files. Nothing to export.", "warning")
            return

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
        
        export_to_excel(final_df, folder_path, output_full_filename, self.console_output)

        self.console_output.print_message(f"Finish processing the CSV files.", "info")


def main():
    root = tk.Tk()
    app = SDDConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()