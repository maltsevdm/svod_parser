from openpyxl.worksheet.worksheet import Worksheet

from common.consts import ROW_HEADERS
from common.enums import SvodHeaders


def get_svod_columns(ws: Worksheet) -> dict[str, int]:
    headers = []
    for col in range(1, ws.max_column + 1):
        value = str(ws.cell(ROW_HEADERS, col).value).strip()

        headers.append(value)

    columns = {}
    for header in SvodHeaders:
        h_value = header.value
        if header.name.startswith("ЕД_ИЗМ"):
            h_value = h_value[:-1]

        if h_value not in headers:
            raise Exception(f"Не удалось найти колонку {header}")

        h_index = headers.index(h_value)
        columns[header] = h_index + 1
        if header.name.startswith("ЕД_ИЗМ"):
            headers[h_index] = "Ед. изм. temp."

    return columns
