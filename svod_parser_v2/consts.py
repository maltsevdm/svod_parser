from .enums import SpecColumn, SvodColumn

from openpyxl.utils import column_index_from_string as cifm

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
    cifm(SvodColumn.НАИМЕНОВАНИЕ_ПО_ПРОЕКТУ): cifm(SpecColumn.НАИМЕНОВАНИЕ_ПО_ПРОЕКТУ),
    cifm(SvodColumn.ВНЕШНИЙ_ВИД): cifm(SpecColumn.ВНЕШНИЙ_ВИД),
    cifm(SvodColumn.НАИМЕНОВАНИЕ_В_РП): cifm(SpecColumn.НАИМЕНОВАНИЕ_В_РП),
    cifm(SvodColumn.ЦЕНА_ОТ_ЗАКУПОК_С_НДС): cifm(SpecColumn.ЦЕНА_ОТ_ЗАКУПОК_С_НДС),
    cifm(SvodColumn.ПОСТАВЩИК): cifm(SpecColumn.ПОСТАВЩИК),
    cifm(SvodColumn.ЕД_ИЗМ): cifm(SpecColumn.ЕД_ИЗМ),
    cifm(SvodColumn.ЦЕНА_ОТ_ЗАКУПОК_С_НДС_РОЗН): cifm(
        SpecColumn.ЦЕНА_ОТ_ЗАКУПОК_С_НДС_РОЗН
    ),
    cifm(SvodColumn.ОЖИДАЕМАЯ_СКИДКА_К_РОЗН_ЦЕНЕ): cifm(
        SpecColumn.ОЖИДАЕМАЯ_СКИДКА_К_РОЗН_ЦЕНЕ
    ),
    cifm(SvodColumn.ИТОГОВАЯ_ЦЕНА_В_РАСЧЕТЕ): cifm(SpecColumn.ИТОГОВАЯ_ЦЕНА_В_РАСЧЕТЕ),
    cifm(SvodColumn.ОБЪЕМ_ИЗ_РП): cifm(SpecColumn.ОБЪЕМ_ИЗ_РП),
    cifm(SvodColumn.ЗАПАС_ПРОЦ): cifm(SpecColumn.ЗАПАС_ПРОЦ),
    cifm(SvodColumn.ИТОГО_С_ЗАПАСОМ): cifm(SpecColumn.ИТОГО_С_ЗАПАСОМ),
    cifm(SvodColumn.ЧИСТОВАЯ_ОТДЕЛКА): cifm(SpecColumn.ЧИСТОВАЯ_ОТДЕЛКА),
    cifm(SvodColumn.ЧИСТОВАЯ_И_КОРПУСНАЯ_МЕБЕЛЬ): cifm(
        SpecColumn.ЧИСТОВАЯ_И_КОРПУСНАЯ_МЕБЕЛЬ
    ),
    cifm(SvodColumn.ПОЛНАЯ_КОМПЛЕКТАЦИЯ): cifm(SpecColumn.ПОЛНАЯ_КОМПЛЕКТАЦИЯ),
    cifm("FE"): cifm(SpecColumn.ЗАПАС_ПРОЦ),
}
