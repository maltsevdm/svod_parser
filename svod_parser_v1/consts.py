from tkinter import IntVar

from common.enums import SvodHeaders

from .enums import SpecColumn

TEMPLATE_FILENAME = "template.xlsx"

ROW_COLUMNS = 2
IMAGE_MAX_WIDTH = 100
IMAGE_MAX_HEIGHT = 150

ROW_MAIN_HEIGHT = 150
ROW_MINOR_HEIGHT = 15

ROW_NAMES = 5
ROW_FLATS = 7

COL_L_FORMULA = "=M{row}*(1+N{row})"
COL_M_FORMULA = "=R{row}"

SVOD_SHEET = "Свод"

static_rows = [
    "отделка пола",
    "отделка стен",
    "отделка потолка",
    "светильники, розетки и выключатели",
    "мебель и декор",
    "бытовая техника",
    "сантехника",
    "двери и комплектующие",
]

columns_relation = {
    SvodHeaders.НАИМЕНОВАНИЕ_ПО_ПРОЕКТУ: SpecColumn.НАИМЕНОВАНИЕ_ПО_ПРОЕКТУ,
    SvodHeaders.ВНЕШНИЙ_ВИД: SpecColumn.ВНЕШНИЙ_ВИД,
    SvodHeaders.НАИМЕНОВАНИЕ_ПОЛНОЕ: SpecColumn.НАИМЕНОВАНИЕ_ПО_РП,
    SvodHeaders.ЦЕНА_ОТ_ЗАКУПОК_ОПТ: SpecColumn.ЦЕНА_С_НДС,
    SvodHeaders.ПОСТАВЩИК: SpecColumn.ПОСТАВЩИК,
    SvodHeaders.ЕД_ИЗМ_1: SpecColumn.ЕД_ИЗМ_1,
    SvodHeaders.КОЛВО_ДЛЯ_ЗАКУПА_С_УЧЕТОМ_ЗАПАСА: SpecColumn.КОЛВО_ОБЩЕЕ_ДЛЯ_ЗАКУПА,
    SvodHeaders.КОЛВО_ДЛЯ_ЗАКУПА_С_РП: SpecColumn.КОЛВО_ДЛЯ_ЗАКУПА_С_РП,
    SvodHeaders.СТРОИТЕЛЬНЫЙ_ЗАПАС: SpecColumn.СТРОИТЕЛЬНЫЙ_ЗАПАС,
    SvodHeaders.ЕД_ИЗМ_2: SpecColumn.ЕД_ИЗМ_2,
    SvodHeaders.МАТЕРИАЛЫ_С_РП: SpecColumn.МАТЕРИАЛЫ_С_РП,
    SvodHeaders.ЕД_ИЗМ_3: SpecColumn.ЕД_ИЗМ_3,
    SvodHeaders.МАТЕРИАЛЫ_ДЛЯ_ЗАКУПА_С_РП: SpecColumn.МАТЕРИАЛЫ_ДЛЯ_ЗАКУПА_С_РП,
}
flat_column = IntVar(value=0)
