import pandas as pd
import io
import os # Required for os.path.basename
from dynamic_console import DynamicConsole # Impor kelas DynamicConsole

def parse_csv_sections(file_path):
    """
    Reads a CSV file, separating the 'Top' and 'Bottom' parts based on the 'Bottom' delimiter.
    Returns two DataFrame handles.
    """
    file_content = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            file_content = f.readlines()
    except Exception as e:
        DynamicConsole.print_message(f"Error: Failed to read file '{os.path.basename(file_path)}': {e}", "error")
        return pd.DataFrame(), pd.DataFrame()

    bottom_section_start_index = -1
    for i, line in enumerate(file_content):
        if line.strip().lower() == "bottom":
            bottom_section_start_index = i
            break

    # Common header that will be used to read both parts
    common_header_str = "Defect No.,Class Name,Top m,Distance from Left Edge mm,Distance from Right Edge mm,Distance from Center mm,Height mm,Width mm,Segment Width Ratio,Orientation"
    # common_header_list = [h.strip() for h in common_header_str.split(',')] # No need to be explicit as pandas handles it

    top_df = pd.DataFrame()
    bottom_df = pd.DataFrame()

    # Processing the 'Top' section
    top_data_lines = []
    if bottom_section_start_index != -1:
        # Take a row of data, skip the first 3 header rows, stop before "Bottom" and the blank row above it
        top_data_lines = file_content[3 : bottom_section_start_index - 1]
    else:
        # If there is no "Bottom", the whole file is "Top" after the header
        top_data_lines = file_content[3:]
    
    # Filter empty rows and create CSV string for pandas
    top_csv_string = common_header_str + "\n" + "".join([line for line in top_data_lines if line.strip()])
    
    try:
        # Using StringIO to read strings as files
        top_df = pd.read_csv(io.StringIO(top_csv_string), sep=',', na_values=['', 'NULL'])
    except pd.errors.EmptyDataError:
        DynamicConsole.print_message(f"    Warning: TOP section in '{os.path.basename(file_path)}' is empty.", "warning")
    except Exception as e:
        DynamicConsole.print_message(f"    Error: Failed to read TOP data from '{os.path.basename(file_path)}': {e}", "error")

    # Processing the 'Bottom' part
    if bottom_section_start_index != -1:
        # Skip "Bottom" and the next 2 header lines
        bottom_data_lines = file_content[bottom_section_start_index + 3:]
        
        # Filter empty rows and create CSV string for pandas
        bottom_csv_string = common_header_str + "\n" + "".join([line for line in bottom_data_lines if line.strip()])
        
        try:
            bottom_df = pd.read_csv(io.StringIO(bottom_csv_string), sep=',', na_values=['', 'NULL'])
        except pd.errors.EmptyDataError:
            DynamicConsole.print_message(f"    Warning: The BOTTOM section in '{os.path.basename(file_path)}' is empty.", "warning")
        except Exception as e:
            DynamicConsole.print_message(f"    Error: Failed to read BOTTOM data from '{os.path.basename(file_path)}': {e}", "error")
            
    return top_df, bottom_df
