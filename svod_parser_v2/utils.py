import re

from openpyxl.utils.cell import column_index_from_string, get_column_letter

from .consts import columns_relation
from .exceptions import ColumnNotFound


def transform_formula(input_string):
    # Используем регулярное выражение для поиска букв + числа
    def replacer(match):
        letter = match.group(1)

        col_index_from = column_index_from_string(letter)
        col_index_to = columns_relation.get(col_index_from)

        if not col_index_to:
            raise ColumnNotFound

        return f"{get_column_letter(col_index_to)}{{row}}"

    # Применяем замену через re.sub с функцией замены
    return re.sub(r"([A-Z]+)(\d+)", replacer, input_string)
