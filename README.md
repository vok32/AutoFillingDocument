# Autofill documents

This script is designed to automatically fill in Word document templates based on data from an Excel spreadsheet.

## Files

- `auto_fill_documents.py `: The main script for running the program.
- `requirements.txt `: List of dependencies.

## Dependencies

- `openpyxl': For working with Excel files.
- `tkinter`: To create a graphical user interface.
- `docx2txt`: To extract text from Word files.
- `docxtpl`: To fill in Word document templates.
- `docx2pdf': For converting Word documents to PDF.

## Creating an executable file (exe)

You can also create an executable file (exe) from this script. To do this, follow these steps:

1. Install PyInstaller if you don't have it yet:
    ```
    pip install pyinstaller
    ```

2. Run PyInstaller for your script:
    ```
    pyinstaller --onefile auto_fill_documents.py
    ```

3. The executable file will be created in the `dist` folder.

For more information about creating an executable file, see the instructions in the file README.md .

## Developer

- [GitHub](https://github.com/vok32)