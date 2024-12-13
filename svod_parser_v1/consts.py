from tkinter import IntVar
from .enums import SpecColumn, SvodColumn

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
    SvodColumn.NAME_BY_PROJECT: SpecColumn.NAME_BY_PROJECT,
    SvodColumn.APPEARANCE: SpecColumn.APPEARANCE,
    SvodColumn.NAME_BY_RP: SpecColumn.NAME_BY_RP,
    SvodColumn.PRICE_WITH_NDS: SpecColumn.PRICE_WITH_NDS,
    SvodColumn.PROVIDER: SpecColumn.PROVIDER,
    SvodColumn.UNITS_FIRST: SpecColumn.UNITS_FIRST,
    SvodColumn.RP_QUANTITY: SpecColumn.RP_QUANTITY,
    SvodColumn.BUILDING_STOCK: SpecColumn.BUILDING_STOCK,
    SvodColumn.UNITS_SECOND: SpecColumn.UNITS_SECOND,
    SvodColumn.RP_MATERIALS: SpecColumn.RP_MATERIALS,
    SvodColumn.UNITS_THIRD: SpecColumn.UNITS_THIRD,
    SvodColumn.MATERIALS_RP_FOR_BUY: SpecColumn.MATERIALS_RP_FOR_BUY,
}
flat_column = IntVar(value=0)
