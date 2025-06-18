# Python SDD Compiler

This application is designed to convert export results from the SDD application (CSV files) into a compiled XLSB (Excel Binary Workbook) format. You can filter CSV files based on a specified date range.

---

## Key Features

* **CSV to XLSB Conversion:** Transforms structured CSV data into a single, organized XLSB file.
* **Date Range Filtering:** Processes only CSV files modified within a specified start and end date/time range.
* **Automatic Coil Number Extraction:** Automatically extracts coil numbers directly from the CSV filenames.
* **Flexible Output Naming:** Allows for custom output filenames or uses a timestamped default.
* **Dual Interface:** Supports both a user-friendly Graphical User Interface (GUI) and a scriptable Command-Line Interface (CLI).

---

## Prerequisites

Before running the application, ensure you have Python 3.x installed and the necessary libraries.

1. **Python 3.x:** Download from [python.org](https://www.python.org/downloads/).
2. **Required Python Libraries:** Install them using `pip`:

    ```bash
    pip install pandas openpyxl tkcalendar pywin32
    ```

    * `pandas`: For data manipulation.
    * `openpyxl`: Required by pandas for `.xlsx` operations.
    * `tkcalendar`: For the date picker widget in the GUI.
    * `pywin32`: For interacting with Microsoft Excel (converting to XLSB).

    **Important Note for `pywin32`:** After installing `pywin32`, you often need to run a post-installation script to register the COM objects correctly, especially if you encounter issues with Excel automation. Open your Command Prompt/Terminal **as an administrator** and run:

    ```bash
    python -m win32com.client.makepy
    ```

    Follow any on-screen prompts.

---

## How to Use

You can run this application using either its Graphical User Interface (GUI) or via the Command-Line Interface (CLI).

### 1. Using the Graphical User Interface (GUI)

The GUI provides an intuitive way to select folders, dates, and customize settings without typing commands.

#### Running the GUI

1. Navigate to the project directory in your file explorer.
2. Locate the `main_gui.py` file.
3. **Double-click `main_gui.py` to run the application.**
    * **Troubleshooting (Permissions):** If you encounter errors, especially those related to Excel saving (e.g., "SaveAs method of Workbook class failed"), it's often due to insufficient permissions. Try **right-clicking `main_gui.py` and selecting "Run as administrator"**.

#### GUI Interface Explanation

* **CSV Folder Path:** Click the **"Browse"** button to select the folder containing your SDD export CSV files.
* **Start Date / End Date:** Use the **date pickers** to select the start and end dates for filtering CSV files based on their modification date.
* **Output File Name (Optional):** Enter a custom name for your output Excel file (e.g., `MyReport`). If left blank, it will default to `CompiledData_YYYYMMDDHHMMSS.xlsx`.
* **Run Conversion:** Click this button to start the processing and conversion.
* **Process Log:** This area will display real-time messages about the application's progress, warnings, and errors.

### 2. Using the Command-Line Interface (CLI)

The CLI is useful for scripting, automation, or users who prefer command-line operations.

#### Running the CLI

1. Open your **Command Prompt (Windows)** or **Terminal (Linux/macOS)**.
2. Navigate to the directory where your application files (`main.py`, `csv_parser.py`, etc.) are located using the `cd` command.

    ```bash
    cd D:\Other Project\Python\SDD Data Compiler # Example path
    ```

3. Execute the `main.py` script with the required arguments.

#### CLI Arguments

* `--path <folder_path>` **(Required):** The path to the folder containing your CSV files. Enclose paths with spaces in double quotes.
* `--startDate <date_time>` **(Required):** The start date and optionally time for filtering CSV files.
  * **Supported formats:**
    * `YYYY-MM-DD` (e.g., `2025-01-01`)
    * `YYYY-MM-DD HH:mm` (e.g., `2025-01-01 09:00`)
    * `YYYY-MM-DD HH:mm:SS` (e.g., `2025-01-01 09:00:00`)
  * **Note:** If only the date is provided for `endDate`, it will default to the end of that day (23:59:59).
* `--endDate <date_time>` **(Required):** The end date and optionally time for filtering CSV files. Uses the same formats as `--startDate`.
* `--outputFileName <name>` **(Optional):** A custom base name for your output Excel file (without the `.xlsx` or `.xlsb` extension). If omitted, the default will be `CompiledData_YYYYMMDDHHMMSS`.

#### CLI Examples

* **Process CSVs from a specific folder for a full day:**

    ```bash
    python main.py --path "Z:\06. Planning\IT System\05. Member\03. Sandy\SDD" --startDate "2025-06-16" --endDate "2025-06-16"
    ```

* **Process CSVs within a specific time range, with a custom output name:**

    ```bash
    python main.py --path "C:\Users\YourUser\Documents\SDD_Exports" --startDate "2025-05-10 10:30" --endDate "2025-05-10 14:45" --outputFileName "May10_MorningData"
    ```

* **Process CSVs from a network drive (using double quotes for the path):**

    ```bash
    python main.py --path "\\NetworkShare\SDD_Reports\Daily" --startDate "2024-12-01" --endDate "2024-12-05"
    ```

---

## Important Notes

* **Excel Installation:** This application requires **Microsoft Excel to be installed** on your system to perform the XLSB conversion, as it uses COM automation (`pywin32`).
* **Error Handling:** The application provides console/GUI logs for progress, warnings, and errors. Pay attention to these messages if you encounter issues.
* **File Deletion:** After successful conversion to XLSB, the temporary `.xlsx` file is automatically deleted. Ensure no other applications are holding a lock on this file during the process.

Feel free to open an `issue` or contact the developer if you encounter any problems!
