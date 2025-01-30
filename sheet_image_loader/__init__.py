from openpyxl.utils.cell import get_column_letter
from openpyxl_image_loader import SheetImageLoader


class CustomSheetImageLoader(SheetImageLoader):
    def __init__(self, sheet):
        """Loads all sheet images"""
        sheet_images = sheet._images
        for image in sheet_images:
            row = image.anchor._from.row + 1
            col = get_column_letter(image.anchor._from.col + 1)
            self._images[f"{col}{row}"] = image._data
