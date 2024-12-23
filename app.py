from tkinter import Tk, ttk

from svod_parser_v2 import start_process as st_v2

root = Tk()


from svod_parser_v1 import flat_column
from svod_parser_v1 import start_process as st_v1

root.title("Формирование спец NEW")


def validate_input(new_value):
    # Проверяем, является ли введенное значение пустым или состоит только из цифр
    return new_value == "" or new_value.isdigit()


label_column_flat = ttk.Label(root, text="Номер столбца с квартирой:")
label_column_flat.grid(row=0, column=0)


# Создаем валидацию
vcmd = (root.register(validate_input), "%P")

entry_flat_column = ttk.Entry(
    root,
    width=6,
    justify="right",
    textvariable=flat_column,
    validate="key",
    validatecommand=vcmd,
)
entry_flat_column.grid(row=1, column=0)

parse_file_btn = ttk.Button(text="Сформировать (вариант 1)", command=st_v1)
parse_file_btn.grid(column=0, row=2, padx=10)


parse_file_btn = ttk.Button(text="Сформировать (вариант 2)", command=st_v2)
parse_file_btn.grid(column=0, row=3, padx=10, pady=20)

root.mainloop()
