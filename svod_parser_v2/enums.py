from enum import IntEnum
from openpyxl.utils import column_index_from_string as cifs


class SpecColumn(IntEnum):
    НОМЕР = cifs("A")
    НАИМЕНОВАНИЕ_ПО_ПРОЕКТУ = cifs("B")
    ВНЕШНИЙ_ВИД = cifs("C")
    НАИМЕНОВАНИЕ_ПОЛНОЕ = cifs("D")
    ЕД_ИЗМ_1 = cifs("E")
    ЦЕНА_ОТ_ЗАКУПОК_ОПТ = cifs("F")
    ЦЕНА_ОТ_ЗАКУПОК_РОЗН = cifs("G")
    ПОСТАВЩИК = cifs("H")
    ОЖИДАЕМАЯ_СКИДКА_К_ЦЕНЕ = cifs("J")
    ИТОГО_ЦЕНА = cifs("K")
    ЕД_ИЗМ_2 = cifs("M")
    ОБЪЕМЫ_С_РП = cifs("N")
    ЕД_ИЗМ_3 = cifs("P")
    ОБЪЕМ_ДЛЯ_ЗАКУПА = cifs("Q")
    СТРОИТЕЛЬНЫЙ_ЗАПАС_ПРОЦ = cifs("R")
    ИТОГО_ОБЪЕМ_С_ЗАПАСОМ_ДЛЯ_ЗАКУПА = cifs("S")
    ИТОГО_СТОИМОСТЬ_ЧИСТОВОЙ_ОТДЕЛКИ = cifs("U")
    ЧИСТОВАЯ_И_КОРПУСНАЯ_МЕБЕЛЬ = cifs("V")
    ПОЛНАЯ_КОМПЛЕКТАЦИЯ = cifs("W")
