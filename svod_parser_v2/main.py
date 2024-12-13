import logging
import os
import pathlib
from tkinter import filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string as cifm
from openpyxl.utils.cell import get_column_letter
from openpyxl_image_loader import SheetImageLoader

from common.styles import GREEN_COLOR, thin_border

from .consts import (
    IMAGE_MAX_HEIGHT,
    IMAGE_MAX_WIDTH,
    ROW_MAIN_HEIGHT,
    ROW_MINOR_HEIGHT,
    SVOD_SHEET,
    TEMPLATE_FILENAME,
    columns_relation,
    static_rows,
)
from .enums import SpecColumn, SvodColumn
from .exceptions import ColumnNotFound
from .utils import transform_formula

current_dir = os.path.dirname(os.path.abspath(__file__))
template_path = os.path.join(current_dir, TEMPLATE_FILENAME)


def start_process():
    file_path = pathlib.Path(
        filedialog.askopenfilename(
            title="Выберите файл Свод", filetypes=[("Excel files", "*.xlsx")]
        )
    )
    try:
        if main(file_path):
            messagebox.showinfo("Информация", "Готово")
    except Exception as ex:
        logging.exception(ex)
        messagebox.showerror("Ошибка", "Произошла непредвиденная ошибка")


def main(file_path: pathlib.Path) -> bool:
    wb_source_with_formulas = load_workbook(file_path)
    try:
        ws_svod = wb_source_with_formulas[SVOD_SHEET]
    except KeyError:
        messagebox.showerror(
            "Ошибка",
            f'В загруженном файле нет листа "{SVOD_SHEET}", я умею работать только с таким листом',
        )
        return

    wb_new = load_workbook(template_path)
    ws_spec = wb_new.active

    row_spec = 3

    i = 1

    image_loader = SheetImageLoader(ws_svod)

    # Переносим данные
    for row in range(9, ws_svod.max_row + 1):
        if str(ws_svod.cell(row, 3).value).lower() in static_rows:
            ws_spec.cell(
                row_spec, cifm(SpecColumn.НОМЕР), ws_svod.cell(row, 3).value.upper()
            )
            row_spec += 1
            continue

        cell_address = f"{SvodColumn.ВНЕШНИЙ_ВИД}{row}"

        if image_loader.image_in(cell_address):
            im = Image(image_loader.get(cell_address))

            if im.width < IMAGE_MAX_WIDTH:
                im.height = IMAGE_MAX_WIDTH * im.height / im.width
                im.width = IMAGE_MAX_WIDTH

            if im.height > IMAGE_MAX_HEIGHT:
                im.width = IMAGE_MAX_HEIGHT * im.width / im.height
                im.height = IMAGE_MAX_HEIGHT

            if im.width > IMAGE_MAX_WIDTH:
                im.height = IMAGE_MAX_WIDTH * im.height / im.width
                im.width = IMAGE_MAX_WIDTH

            cell_address = f"{SpecColumn.ВНЕШНИЙ_ВИД}{row_spec}"
            ws_spec.add_image(im, cell_address)

        ws_spec.cell(row_spec, cifm(SpecColumn.НОМЕР), i)

        for col_from, col_to in columns_relation.items():
            if col_from in [cifm("FE")]:
                continue

            value = ws_svod.cell(row, col_from).value
            if str(value).startswith("="):
                try:
                    new_formula = transform_formula(value).format(row=row_spec)
                except ColumnNotFound:
                    address = f"{get_column_letter(col_from)}{row}"
                    messagebox.showerror(
                        "Ошибка",
                        f"В своде в ячейке {address} в формуле используется колонка, которой нет в итоговом файле",
                    )
                    return

                ws_spec.cell(row_spec, col_to, new_formula)
            else:
                ws_spec.cell(row_spec, col_to, ws_svod.cell(row, col_from).value)

        row_spec += 1
        i += 1

    # Делаем оформление
    for row in range(3, ws_spec.max_row + 1):
        if not ws_spec.cell(row, 1).value:
            break

        if str(ws_spec.cell(row, cifm(SpecColumn.НОМЕР)).value).lower() in static_rows:
            ws_spec.row_dimensions[row].height = ROW_MINOR_HEIGHT
            ws_spec.merge_cells(
                start_row=row,
                start_column=cifm(SpecColumn.НОМЕР),
                end_row=row,
                end_column=cifm(SpecColumn.ПОЛНАЯ_КОМПЛЕКТАЦИЯ) + 1,
            )
            cell = ws_spec.cell(row, cifm(SpecColumn.НОМЕР))
            cell.fill = PatternFill("solid", GREEN_COLOR)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(bold=True)
            ws_spec.border = thin_border
            continue

        if row >= 3:
            ws_spec.row_dimensions[row].height = ROW_MAIN_HEIGHT
            ws_spec.cell(row, cifm(SpecColumn.ЗАПАС_ПРОЦ)).number_format = "0%"

        for col in range(1, cifm(SpecColumn.ПОЛНАЯ_КОМПЛЕКТАЦИЯ) + 2):
            cell = ws_spec.cell(row, col)

            if row >= 3:
                cell.border = thin_border

            if get_column_letter(col) in [
                SpecColumn.ЦЕНА_ОТ_ЗАКУПОК_С_НДС,
                SpecColumn.ЦЕНА_ОТ_ЗАКУПОК_С_НДС_РОЗН,
                SpecColumn.ОЖИДАЕМАЯ_СКИДКА_К_РОЗН_ЦЕНЕ,
                SpecColumn.ИТОГОВАЯ_ЦЕНА_В_РАСЧЕТЕ,
                SpecColumn.ОБЪЕМ_ИЗ_РП,
                SpecColumn.ЧИСТОВАЯ_ОТДЕЛКА,
                SpecColumn.ИТОГО_С_ЗАПАСОМ,
            ]:
                ws_spec.cell(row, col).number_format = "0.0"

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

    try:
        wb_new.save(file_path.with_name("Excel казахстан.xlsx"))
    except PermissionError:
        messagebox.showerror(
            "Ошибка", "Закройте файл Excel казахстан.xlsx и попробуйте снова"
        )
        return False

    return True
