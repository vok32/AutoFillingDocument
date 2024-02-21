import re
import os
import openpyxl as op
from docx2txt import process
from tkinter import Tk, Label, Button, Listbox, Scrollbar, messagebox
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


def show_header_and_variable_selection_ui(root, header_row, template_variables):
    # Очищаем содержимое основного окна
    clear_window(root)

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

    continue_button = Button(root, text="Продолжить", command=lambda: show_differences_ui(root, header_row, template_variables))
    continue_button.grid(row=2, column=0, pady=10)

    close_button = Button(root, text="Закрыть программу", command=close_program)
    close_button.grid(row=4, column=0, pady=5)


def show_differences_ui(root, header_row, template_variables):
    # Очищаем содержимое основного окна
    clear_window(root)

    differences_message = compare_headers_and_variables(header_row, template_variables)
    differences_label = Label(root, text=differences_message, justify="left")
    differences_label.grid(row=0, column=0, sticky="w")

    max_width = len(differences_message)
    set_label_width(differences_label, max_width)

    replace_button = Button(root, text="Продолжить и заменить", command=lambda: select_column(root, header_row, template_variables))
    replace_button.grid(row=1, column=0, pady=5)

    close_button = Button(root, text="Закрыть программу", command=close_program)
    close_button.grid(row=3, column=0, pady=5)


def select_column(root, header_row, template_variables):
    # Очищаем содержимое основного окна
    clear_window(root)

    wb_read = op.load_workbook(filename=EXCEL_PATH, data_only=True)
    sheet_read = wb_read.active

    akt_list = [list(row) for row in sheet_read.iter_rows(values_only=True)]

    def select_column_callback():
        selected_index = listbox.curselection()
        if selected_index:
            selected_column_name = listbox.get(selected_index[0])  # Получаем название столбца из списка
            selected_column = header_row[selected_column_name]  # Получаем индекс столбца по его названию
            # Здесь мы добавляем новое окно для выбора формата файла
            choice_root = Tk()
            choice_root.title("Выбор формата файла")

            def docx_pdf():
                choice_root.destroy()
                create_doc(root, akt_list, header_row, selected_column, selected_column_name, convert_to_pdf=True, delete_docx=False)
                show_success_or_report_window(root)

            def docx_only():
                choice_root.destroy()
                create_doc(root, akt_list, header_row, selected_column, selected_column_name, convert_to_pdf=False)
                show_success_or_report_window(root)

            def pdf_only():
                choice_root.destroy()
                create_doc(root, akt_list, header_row, selected_column, selected_column_name, convert_to_pdf=True, delete_docx=True)
                show_success_or_report_window(root)

            pdf_button = Button(choice_root, text="DOCX and PDF", command=docx_pdf)
            pdf_button.grid(row=0, column=0, padx=10, pady=5)

            docx_button = Button(choice_root, text="DOCX Only", command=docx_only)
            docx_button.grid(row=0, column=1, padx=10, pady=5)

            pdf_only_button = Button(choice_root, text="PDF Only", command=pdf_only)
            pdf_only_button.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

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

def show_success_or_report_window(root):
    # Очищаем содержимое основного окна
    clear_window(root)

    success_label = Label(root, text="Замена прошла успешно", justify="left", fg="green", font=("Arial", 12))
    success_label.grid(row=0, column=0, pady=5)

    report_label = Label(root, text="Вы хотите отправить разработчику отчёт об ошибке в приложении?", justify="left")
    report_label.grid(row=1, column=0, pady=5)

    yes_button = Button(root, text="Да, отправить", command=lambda: messagebox.showinfo("Ну и ябеда!", "Ну и ябеда!"))
    yes_button.grid(row=2, column=0, pady=5)

    close_button = Button(root, text="Закрыть программу", command=close_program)
    close_button.grid(row=3, column=0, pady=5)


def clear_window(root):
    # Очищаем содержимое основного окна
    for widget in root.winfo_children():
        widget.destroy()


def set_label_width(label, max_width):
    label.config(width=max_width)


def create_doc(root, akt_list, header_row, column_index, column_name, convert_to_pdf=True, delete_docx=True):
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
            if delete_docx:
                os.remove(doc_path)  # Удаляем DOCX файл после конвертации в PDF
                print(f"Файл {doc_path} удален успешно.")
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

def excel_read(root, path_file):
    wb_read = op.load_workbook(filename=path_file, data_only=True)
    sheet_read = wb_read.active

    header_row = {cell.value: cell.column_letter for cell in sheet_read[1]}
    template_variables = parse_template(TEMPLATE_PATH)

    differences = compare_headers_and_variables(header_row, template_variables)
    if differences != "Заголовки в файле Ecel совпадают с переменными в шаблоне.":
        show_header_and_variable_selection_ui(root, header_row, template_variables)
    else:
        print("Заголовки в файле Excel совпадают с переменными в шаблоне.")
        # Выбираем произвольный заголовок Excel для создания файла
        default_column_name = next(iter(header_row.keys()))
        create_doc(root, [], header_row, 1, default_column_name)


def on_closing(root):
    if messagebox.askokcancel("Выход", "Вы уверены, что хотите выйти?"):
        root.destroy()


if __name__ == '__main__':
    root = Tk()
    root.title("Моя программа")
    root.protocol("WM_DELETE_WINDOW", lambda: on_closing(root))

    excel_read(root, EXCEL_PATH)

    root.mainloop()