from openpyxl.utils import column_index_from_string as cifs

from common.enums import SvodHeaders

from .enums import SpecColumn

TEMPLATE_FILENAME = "template.xlsx"

ROW_COLUMNS = 2
IMAGE_MAX_WIDTH = 100
IMAGE_MAX_HEIGHT = 150

ROW_MAIN_HEIGHT = 150
ROW_MINOR_HEIGHT = 15


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
    SvodHeaders.НАИМЕНОВАНИЕ_ПОЛНОЕ: SpecColumn.НАИМЕНОВАНИЕ_ПОЛНОЕ,
    SvodHeaders.ЕД_ИЗМ_1: SpecColumn.ЕД_ИЗМ_1,
    SvodHeaders.ЦЕНА_ОТ_ЗАКУПОК_ОПТ: SpecColumn.ЦЕНА_ОТ_ЗАКУПОК_ОПТ,
    SvodHeaders.ЦЕНА_ОТ_ЗАКУПОК_РОЗН: SpecColumn.ЦЕНА_ОТ_ЗАКУПОК_РОЗН,
    SvodHeaders.ПОСТАВЩИК: SpecColumn.ПОСТАВЩИК,
    SvodHeaders.ОЖИДАЕМАЯ_СКИДКА_К_ЦЕНЕ: SpecColumn.ОЖИДАЕМАЯ_СКИДКА_К_ЦЕНЕ,
    SvodHeaders.ИТОГО_ЦЕНА: SpecColumn.ИТОГО_ЦЕНА,
    SvodHeaders.ЕД_ИЗМ_4: SpecColumn.ЕД_ИЗМ_2,
    SvodHeaders.ОБЪЕМЫ_С_РП: SpecColumn.ОБЪЕМЫ_С_РП,
    SvodHeaders.ЕД_ИЗМ_5: SpecColumn.ЕД_ИЗМ_3,
    SvodHeaders.ОБЪЕМ_ДЛЯ_ЗАКУПА: SpecColumn.ОБЪЕМ_ДЛЯ_ЗАКУПА,
    SvodHeaders.СТРОИТЕЛЬНЫЙ_ЗАПАС: SpecColumn.СТРОИТЕЛЬНЫЙ_ЗАПАС_ПРОЦ,
    SvodHeaders.ИТОГО_ОБЪЕМ_С_ЗАПАСОМ_ДЛЯ_ЗАКУПА: SpecColumn.ИТОГО_ОБЪЕМ_С_ЗАПАСОМ_ДЛЯ_ЗАКУПА,
    SvodHeaders.ИТОГО_СТОИМОСТЬ_ЧИСТОВОЙ_ОТДЕЛКИ: SpecColumn.ИТОГО_СТОИМОСТЬ_ЧИСТОВОЙ_ОТДЕЛКИ,
    SvodHeaders.ЧИСТОВАЯ_И_КОРПУСНАЯ_МЕБЕЛЬ: SpecColumn.ЧИСТОВАЯ_И_КОРПУСНАЯ_МЕБЕЛЬ,
    SvodHeaders.ПОЛНАЯ_КОМПЛЕКТАЦИЯ: SpecColumn.ПОЛНАЯ_КОМПЛЕКТАЦИЯ,
    SvodHeaders.СТРОИТЕЛЬНЫЙ_ЗАПАС_ПРОЦ: SpecColumn.СТРОИТЕЛЬНЫЙ_ЗАПАС_ПРОЦ,
}
