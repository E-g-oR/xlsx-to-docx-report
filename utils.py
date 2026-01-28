import openpyxl
from typing import Iterable

def get_sheet_from_file(file: str):
    wb = openpyxl.load_workbook(file, data_only=True, read_only=True)
    sheet = wb.active
    return sheet

def get_anchors_coordinates(anchors: Iterable[str], sheet):
    need = set(anchors)
    found = {}
    for col in range(1, sheet.max_column + 1):
        for row in range(1, sheet.max_row + 1):
            value = sheet.cell(row=row, column=col).value
            if value is not None and value in need:
                found[value] = (row, col)
                need.remove(value)
                if not need:
                    return found
    return found

def get_next_right_value(coord, sheet):
    (row, col) = coord
    for c in range(col + 1, sheet.max_column + 1):
        value = sheet.cell(row=row, column=c).value
        if value is not None:
            return value

def get_last_value_in_row(coord, sheet):
    (row, col) = coord
    last_value = None
    for c in range(col+1, sheet.max_column + 1):
        value = sheet.cell(row=row, column=c).value
        if value is not None:
            last_value = value
    return last_value