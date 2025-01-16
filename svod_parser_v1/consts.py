
from common.enums import SvodHeaders

from .enums import SpecColumn

TEMPLATE_FILENAME = "template.xlsx"

IMAGE_MAX_WIDTH = 100
IMAGE_MAX_HEIGHT = 125

ROW_MAIN_HEIGHT = 125
ROW_MINOR_HEIGHT = 15

COL_L_FORMULA = "=M{row}*(1+N{row})"
COL_M_FORMULA = "=R{row}"

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
