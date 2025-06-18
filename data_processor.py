import re
import pandas as pd # Required for pandas Series

def extract_coil_no_from_filename(filename, console_instance):
    """
    Extracts the 'Coil No' (for example, KE5538) from the file name.
    Pattern: 2 uppercase letters followed by 4 numbers.
    """
    pattern = r"\.\d{2}-\d{2}-\d{2}\.([A-Z]{2}\d{4})\s+\d{2}\.Defects\.csv"
    match = re.search(pattern, filename)
    if match and len(match.groups()) > 0:
        return match.group(1)
    console_instance.print_message(
        f"    Warning: Coil No is not found in file name '{filename}' or does not match the new pattern. Using an empty string.", "warning"
    )
    return ""

def process_record(record_series: pd.Series, coil_no: str, top_bottom_status: str):
    """
    Process a single record (pandas Series) and convert it to the desired Excel format.
    Returns the dictionary of the processed record or None if it should be skipped.
    """
    new_record = {}

    new_record["Coil No"] = coil_no

    # Take Class Name, make sure it is not NaN (Not a Number)
    class_name_str = str(record_series.get("Class Name", "")).strip()
    if class_name_str == "nan": # pandas converts empty strings to NaN
        class_name_str = ""

    class_name_parts = class_name_str.split('-')

    new_record["Class Name"] = class_name_parts[0].strip() if len(class_name_parts) > 0 else ""
    new_record["Defect Name"] = class_name_parts[3].strip() if len(class_name_parts) > 3 else ""
    new_record["Grade Defect"] = class_name_parts[1].strip()[0] if len(class_name_parts) > 1 and len(class_name_parts[1]) > 0 else ""

    # IF THE CLASS NAME IS EMPTY, DO NOT ADD THIS LINE
    if not new_record["Class Name"]: # Check if the Class Name is empty after processing
        return None # Indicates this record should be skipped

    new_record["Top/Bottom"] = top_bottom_status

    # Other column mappings
    new_record["Distance from HE CGL (m)"] = record_series.get("Top m", None)
    new_record["Distance Left (mm)"] = record_series.get("Distance from Left Edge mm", None)
    new_record["Distance Right (mm)"] = record_series.get("Distance from Right Edge mm", None)
    new_record["Distance Center (mm)"] = record_series.get("Distance from Center mm", None)
    new_record["Height"] = record_series.get("Height mm", None)
    new_record["Width"] = record_series.get("Width mm", None)
    new_record["Segment Width Ratio"] = record_series.get("Segment Width Ratio", None)
    new_record["Orientation"] = record_series.get("Orientation", None)

    return new_record
