import re
import os
import openpyxl as op
from docx2txt import process
from tkinter import Tk, Label, Button, Listbox, Scrollbar, messagebox, simpledialog
from docxtpl import DocxTemplate
from docx2pdf import convert

TEMPLATE_PATH = "C:\\Users\\Roman\\Desktop\\Макрос\\Шаблон\\ШаблонПовышка.docx"
EXCEL_PATH = "C:\\Users\\Roman\\Desktop\\Макрос\\ДанныеДляПовышки.xlsx"
SAVE_PATH = "C:\\Users\\Roman\\Desktop\\Макрос\\Собранные файлы"

def close_program():
    exit()

def parse_template(template_path):
    text = process(template_path)
    variables = set(re.findall(r'\{\{(.+?)\}\}', text))  # Используем множество для уникальных переменных
    return variables

def compare_headers_and_variables(header_row, template_variables):
    excel_headers_set = set(header_row.keys())
    template_variables_set = set(template_variables)
    
    if excel_headers_set != template_variables_set:
        missing_in_excel = excel_headers_set - template_variables_set
        missing_in_template = template_variables_set - excel_headers_set
        message = "Отличия между заголовками в файле Excel и переменными в шаблоне:\n\n"
        if missing_in_excel:
            message += "Заголовки в файле Excel, но отсутствующие в шаблоне:\n"
            for header in missing_in_excel:
                message += f"- {header}\n"
            message += "\n"
        if missing_in_template:
            message += "Переменные в шаблоне, но отсутствующие в заголовках файла Excel:\n"
            for var in missing_in_template:
                message += f"- {var}\n"
        return message
    else:
        return "Заголовки в файле Excel совпадают с переменными в шаблоне."

def show_header_and_variable_selection_ui(header_row, template_variables):
    root = Tk()
    root.title("Выбор заголовков и переменных")

    def on_closing():
        if messagebox.askokcancel("Выход", "Вы уверены, что хотите выйти?"):
            exit()

    def go_back():
        root.destroy()
        excel_read(EXCEL_PATH)  # Перезапуск процесса с передачей пути к файлу Excel

    root.protocol("WM_DELETE_WINDOW", on_closing)

    excel_headers_message = "Список заголовков в файле Excel:\n"
    for index, column_name in enumerate(header_row, start=1):
        excel_headers_message += f"{index}. {column_name}\n"

    template_variables_message = ""
    if template_variables:
        template_variables_message = "Переменные в шаблоне:\n"
        for index, var in enumerate(template_variables, start=1):
            template_variables_message += f"{index}. {var}\n"

    excel_headers_label = Label(root, text=excel_headers_message, justify="left")
    excel_headers_label.grid(row=0, column=0, sticky="w")

    template_variables_label = Label(root, text=template_variables_message, justify="left")
    template_variables_label.grid(row=1, column=0, sticky="w")

    continue_button = Button(root, text="Продолжить", command=lambda: root.destroy())
    continue_button.grid(row=2, column=0, pady=10)

    close_button = Button(root, text="Закрыть программу", command=close_program)
    close_button.grid(row=4, column=0, pady=5)

    root.mainloop()

def show_differences_ui(header_row, template_variables):
    root = Tk()
    root.title("Отличия между заголовками и переменными")

    def on_closing():
        if messagebox.askokcancel("Выход", "Вы уверены, что хотите выйти?"):
            exit()

    def go_back():
        root.destroy()
        show_header_and_variable_selection_ui(header_row, template_variables)

    root.protocol("WM_DELETE_WINDOW", on_closing)

    differences_message = compare_headers_and_variables(header_row, template_variables)
    differences_label = Label(root, text=differences_message, justify="left")
    differences_label.grid(row=0, column=0, sticky="w")

    replace_button = Button(root, text="Продолжить и заменить", command=lambda: root.destroy())
    replace_button.grid(row=1, column=0, pady=5)

    close_button = Button(root, text="Закрыть программу", command=close_program)
    close_button.grid(row=3, column=0, pady=5)

    root.mainloop()

def show_success_or_report_window(success=True):
    root = Tk()
    root.title("Успех или отчет об ошибке")

    def on_closing():
        if messagebox.askokcancel("Выход", "Вы уверены, что хотите выйти?"):
            exit()

    def go_back():
        root.destroy()
        show_differences_ui(header_row, template_variables)

    def on_yes_clicked():
        messagebox.showinfo("Ну и ябеда!", "Ну и ябеда!")
        exit()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    if success:
        success_label = Label(root, text="Замена прошла успешно", justify="left", fg="green", font=("Arial", 12))
        success_label.grid(row=0, column=0, pady=5)

    report_label = Label(root, text="Вы хотите отправить разработчику отчёт об ошибке в приложении?", justify="left")
    report_label.grid(row=1, column=0, pady=5)

    yes_button = Button(root, text="Да, отправить", command=on_yes_clicked)
    yes_button.grid(row=2, column=0, pady=5)

    close_button = Button(root, text="Закрыть программу", command=close_program)
    close_button.grid(row=4, column=0, pady=5)

    root.mainloop()

def select_column(header_row, template_variables, akt_list):
    root = Tk()
    root.title("Выбор столбца")

    def on_closing():
        if messagebox.askokcancel("Выход", "Вы уверены, что хотите выйти?"):
            exit()

    def go_back():
        root.destroy()
        show_differences_ui(header_row, template_variables)

    root.protocol("WM_DELETE_WINDOW", on_closing)

    def select_column_callback():
        selected_index = listbox.curselection()
        if selected_index:
            selected_column_name = listbox.get(selected_index[0])  # Получаем название столбца из списка
            selected_column = header_row[selected_column_name]  # Получаем индекс столбца по его названию
            # Здесь мы добавляем новое окно для выбора формата файла
            choice_root = Tk()
            choice_root.title("Выбор формата файла")

            def pdf_convert():
                choice_root.destroy()
                root.destroy()
                create_doc(akt_list, header_row, selected_column, selected_column_name, convert_to_pdf=True)
                show_success_or_report_window()

            def docx_only():
                choice_root.destroy()
                root.destroy()
                create_doc(akt_list, header_row, selected_column, selected_column_name, convert_to_pdf=False)
                show_success_or_report_window()

            pdf_button = Button(choice_root, text="PDF", command=pdf_convert)
            pdf_button.grid(row=0, column=0, padx=10, pady=5)

            docx_button = Button(choice_root, text="DOCX Only", command=docx_only)
            docx_button.grid(row=0, column=1, padx=10, pady=5)

            choice_root.mainloop()
        else:
            messagebox.showinfo("Ошибка", "Не выбрано название для файла. Пожалуйста, выберите название.")

    selection_label = Label(root, text="Выберите название для файлов:")
    selection_label.grid(row=0, column=0, sticky="w")

    listbox = Listbox(root, selectmode="SINGLE")
    listbox.grid(row=1, column=0, sticky="w")

    scrollbar = Scrollbar(root, orient="vertical")
    scrollbar.config(command=listbox.yview)
    scrollbar.grid(row=1, column=1, sticky="ns")

    listbox.config(yscrollcommand=scrollbar.set)
    for column_name in header_row:
        listbox.insert("end", column_name)

    select_button = Button(root, text="Выбрать", command=select_column_callback)
    select_button.grid(row=2, column=0, pady=10)

    close_button = Button(root, text="Закрыть программу", command=close_program)
    close_button.grid(row=4, column=0, pady=5)

    root.mainloop()

def create_doc(akt_list, header_row, column_index, column_name, convert_to_pdf=True):
    # Преобразуем заголовок в словарь для более быстрого доступа к индексам столбцов
    header_dict = {name: index for index, name in enumerate(header_row)}
    
    # Заполняем шаблон данными и создаем документы для каждой строки
    for row_data in akt_list:
        context = {}
        for variable, column in header_row.items():
            index = header_dict.get(variable)  # Получаем индекс столбца по имени переменной
            if index is not None:
                context[variable] = row_data[index]  # Получаем данные из соответствующего столбца
            else:
                context[variable] = ''  # Если столбец отсутствует, используем пустую строку
        
        # Получаем заголовок из первой строки Excel
        first_row_header = list(header_row.keys())[0]
        
        # Заменяем недопустимые символы в названии файла
        file_name = re.sub(r'[\\/*?:"<>|]', '_', str(context[column_name]))
        
        # Сохраняем документ в формате DOCX
        doc = DocxTemplate(TEMPLATE_PATH)
        doc.render(context)
        doc_path = os.path.join(SAVE_PATH, f"шаблон-{file_name}.docx")
        doc.save(doc_path)
        
        if convert_to_pdf:
            # Конвертируем документ в формат PDF
            pdf_path = os.path.join(SAVE_PATH, f"шаблон-{file_name}.pdf")
            convert(doc_path, pdf_path)
            print(f"Созданы файлы: {doc_path} (DOCX) и {pdf_path} (PDF)")
        else:
            print(f"Создан файл: {doc_path} (DOCX)")

    # Удаляем файлы с названием "шаблон-None", если они были созданы
    none_file_path = os.path.join(SAVE_PATH, "шаблон-None.docx")
    if os.path.exists(none_file_path):
        os.remove(none_file_path)
        print(f"Файл {none_file_path} удален успешно.")

    none_file_pdf_path = os.path.join(SAVE_PATH, "шаблон-None.pdf")
    if os.path.exists(none_file_pdf_path):
        os.remove(none_file_pdf_path)
        print(f"Файл {none_file_pdf_path} удален успешно.")

    # Удаляем файлы с названием "шаблон-выбранное_название_для_файлов", если они были созданы
    selected_file_path = os.path.join(SAVE_PATH, f"шаблон-{column_name}.docx")
    if os.path.exists(selected_file_path):
        os.remove(selected_file_path)
        print(f"Файл {selected_file_path} удален успешно.")

    selected_file_pdf_path = os.path.join(SAVE_PATH, f"шаблон-{column_name}.pdf")
    if os.path.exists(selected_file_pdf_path):
        os.remove(selected_file_pdf_path)
        print(f"Файл {selected_file_pdf_path} удален успешно.")
            
def excel_read(path_file):
    wb_read = op.load_workbook(filename=path_file, data_only=True)
    sheet_read = wb_read.active
    
    header_row = {cell.value: cell.column_letter for cell in sheet_read[1]}
    template_variables = parse_template(TEMPLATE_PATH)
    akt_list = [list(row) for row in sheet_read.iter_rows(values_only=True)]

    differences = compare_headers_and_variables(header_row, template_variables)
    if differences != "Заголовки в файле Ecel совпадают с переменными в шаблоне.":
        show_header_and_variable_selection_ui(header_row, template_variables)
        show_differences_ui(header_row, template_variables)
        select_column(header_row, template_variables, akt_list)
    else:
        print("Заголовки в файле Excel совпадают с переменными в шаблоне.")
        # Выбираем произвольный заголовок Excel для создания файла
        default_column_name = next(iter(header_row.keys()))
        create_doc(akt_list, header_row, 1, default_column_name)

if __name__ == '__main__':
    path_file = EXCEL_PATH
    excel_read(path_file)
