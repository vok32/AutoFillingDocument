# Document Auto-filling

This script is designed to automatically fill Word document templates based on data from an Excel spreadsheet.

## Description

The `auto_fill_documents.py` script is a Python program that provides automatic filling of Word document templates based on data provided in Excel format. It is created using various Python libraries and modules such as `openpyxl` for Excel manipulation, `docxtpl` for Word template filling, and `Tkinter` for creating a user interface.

## Usage

1. Run the `auto_fill_documents.py` script.
2. Select the Word template file, the Excel file with data, and the folder to save the results.
3. Click the "Continue" button to start filling the documents.
4. After the document creation process is complete, you will see a corresponding notification.

## Functionality

- **File and Folder Selection**: Users can select Word template and Excel files using the respective buttons.
- **Document Filling**: The program populates Word templates with data from Excel, creating new documents.
- **Result Saving**: Generated documents are saved in the user-selected folder.
- **User Interface**: The program offers a simple and intuitive user interface for easy management of the document filling process.

## Requirements

To run this script, you need to have the following Python libraries installed:

- `openpyxl`
- `docxtpl`
- `docx2pdf`
- `docx2txt`
- `Tkinter`

## Built-in Functions

The script includes the following built-in functions:

- **close_program()**: Closes the program.
- **restart_program()**: Restarts the program.
- **delete_files_with_pattern()**: Deletes files with the header name.
- **generate_unique_suffix(length)**: Generates a unique suffix of the specified length.
- **parse_template(template_path)**: Parses variables from the Word template.
- **select_template_file()**: Opens a dialog box to select a Word template file.
- **select_excel_file()**: Opens a dialog box to select an Excel file with data.
- **select_save_folder()**: Opens a dialog box to select a folder to save files.
- **open_folder()**: Opens the folder with the compiled files.
- **compare_headers_and_variables(header_row, template_variables)**: Compares Excel headers and template variables.
- **show_header_and_variable_selection_ui(root, header_row, template_variables)**: Displays the interface for selecting headers and variables.
- **show_differences_ui(root, header_row, template_variables)**: Displays the interface for differences between headers and variables.
- **set_label_width(label, max_width)**: Sets the label width.
- **select_column(root, header_row, template_variables)**: Displays the interface for selecting an Excel column.
- **show_success_or_report_window(root)**: Displays a window for successful completion or error report.
- **clear_window(root)**: Clears the main window.
- **create_doc(root, akt_list, header_row, column_index, column_name, convert_to_pdf=True, delete_docx=True)**: Creates a document.
- **excel_read(root, path_file)**: Reads data from an Excel file.
- **on_closing(root)**: Handles window closure.
- **show_developer_info()**: Displays developer information.

## Example of Variable Usage in Word Template

```bash
Dear {{ Name }},

We are pleased to inform you that your order No.{{ Order Number }} has been successfully processed.
```
In this example, {{ Name }} and {{ Order Number }} are variables that will be replaced with specific values from the Excel file.

## Important Note

The program assumes that the data in the Excel table has a specific structure, with corresponding column headers that match the variables in the Word template.

## Executable File Setup

### Method 1: Using PyInstaller

1. **PyInstaller Installation**:  If you haven't installed PyInstaller yet, do so using pip:
   pip install pyinstalle

2. **Creating the Executable File:**: Navigate to the directory containing your Python script (auto_fill_documents.py) using the command line, then execute PyInstaller:
    ```bash
    pyinstaller --onefile auto_fill_documents.py

3. **Locating the Executable File**: After the PyInstaller build process is complete, you will find the executable file in the dist directory inside the directory containing your script.

4. **Running the Executable File**: Double-click the created executable file (auto_fill_documents.exe) to run the program.

### Method 2: Using Auto-py-to-exe

1. **Auto-py-to-exe Installation**: If you haven't installed Auto-py-to-exe yet, do so using pip:
    ```bash
    pip install auto-py-to-exe

2. **Running Auto-py-to-exe**: Launch Auto-py-to-exe, enter the command auto-py-to-exe in the command line or terminal.

3. **Setting Parameters**: In the Auto-py-to-exe window, specify the path to your Python script (auto_fill_documents.py). Choose the desired build parameters, such as single file or multiple files, and click the "Convert .py to .exe" button.

4. **Locating the Executable File**: After the conversion process is complete, Auto-py-to-exe will create the executable file in the location you specified.
5. **Running the Executable File**: Double-click the created executable file to run the program.

## Versions
v1.0 - release

v1.1 - unique suffix "file_name_new_*" and universal replacement/cancellation of the replacement of the collected files

## Author

This script was developed by  [R V](https://github.com/vok32).