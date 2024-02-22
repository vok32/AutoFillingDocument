# Document Autofill

This script is designed to automatically fill Word document templates based on data from an Excel spreadsheet.

## Description

The script `auto_fill_documents.py` is a Python program that provides automatic filling of Word document templates based on data provided in Excel format. It is created using various Python libraries and modules, such as `openpyxl` for working with Excel, `docxtpl` for filling Word templates, and `Tkinter` for creating the user interface.

## Usage

1. Run the `auto_fill_documents.py` script.
2. Select the Word template file, Excel file with data, and folder to save the results.
3. Click the "Continue" button to start filling the documents.
4. After the document creation process is complete, you will see a corresponding notification.

## Functionality

- **File and Folder Selection**: The user can select Word template and Excel files using the corresponding buttons.
- **Document Filling**: The program fills Word templates with data from Excel, creating new documents.
- **Results Saving**: Generated documents are saved in the user-selected folder.
- **User Interface**: The program provides a simple and intuitive user interface, allowing easy control over the document filling process.

## Requirements

To run this script, the following Python libraries are required:

- `openpyxl`
- `docxtpl`
- `docx2pdf`
- `docx2txt`
- `Tkinter`

## Built-in Functions

The script provides the following built-in functions:

- **close_program()**: Closes the program.
- **restart_program()**: Restarts the program.
- **generate_unique_suffix(length)**: Generates a unique suffix of a specified length.
- **parse_template(template_path)**: Parses variables from the Word template.
- **select_template_file()**: Opens a dialog window to select a Word template file.
- **select_excel_file()**: Opens a dialog window to select an Excel file with data.
- **select_save_folder()**: Opens a dialog window to select a folder to save files.
- **open_folder()**: Opens the folder with assembled files.
- **compare_headers_and_variables(header_row, template_variables)**: Compares Excel headers and template variables.
- **show_header_and_variable_selection_ui(root, header_row, template_variables)**: Displays the interface for selecting headers and variables.
- **show_differences_ui(root, header_row, template_variables)**: Displays the interface for differences between headers and variables.
- **set_label_width(label, max_width)**: Sets the width of the label.
- **select_column(root, header_row, template_variables)**: Displays the interface for selecting an Excel column.
- **show_success_or_report_window(root)**: Displays a window for successful execution or error report.
- **clear_window(root)**: Clears the main window.
- **create_doc(root, akt_list, header_row, column_index, column_name, convert_to_pdf=True, delete_docx=True)**: Creates a document.
- **excel_read(root, path_file)**: Reads data from an Excel file.
- **on_closing(root)**: Handles window closing.
- **show_developer_info()**: Displays developer information.

## Example of Using Variables in Word Template

\```markdown
Dear {{ Name }},

We are pleased to inform you that your order No.{{ Order Number }} has been successfully processed.
\```

In this example, `{{ Name }}` and `{{ Order Number }}` are variables that will be replaced with specific values from the Excel file.

## Important Note

The program assumes that the data in the Excel table has a specific structure, with corresponding column headers that match the variables in the Word template.

## Author

This script was developed by [Roman](https://github.com/vok32).