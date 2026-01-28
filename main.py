from pathlib import Path
from typing import Iterable
import openpyxl

ANCHORS = {
    "Счет-фактура №": "NEXT_RIGHT",
    "Грузополучатель и его адрес:": "NEXT_RIGHT",
    "Всего к оплате (9)": "LAST_RIGHT",
}

def get_sheet_from_file(file: str):
    wb = openpyxl.load_workbook(file)
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

def get_next_row_value(coord, sheet):
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

rules = {
    "NEXT_RIGHT": get_next_row_value,
    "LAST_RIGHT": get_last_value_in_row
}

files = sorted(p.name for p in Path(".").glob("УПД*.xlsx"))

numbers = set()
consignees = set()
payment_sum = 0

required = set(ANCHORS.keys())

for file in files:
    print("Читаю файл:", file)
    sheet = get_sheet_from_file(file)

    found = get_anchors_coordinates(ANCHORS.keys(), sheet)
    missing_anchors = required - found.keys()
    if missing_anchors:
        print("Не найдены якоря:", ", ".join(sorted(missing_anchors)))
        continue


    values = {}
    for label, rule_name in ANCHORS.items():
        if label not in found:
            continue
        coords = found[label]
        values[label] = rules[rule_name](coords, sheet)

    missing_values = [k for k in required if values.get(k) is None]
    if missing_values:
        print("Не найдены значения для:", ", ".join(sorted(missing_values)))
        continue


    number = values["Счет-фактура №"]
    consignee = values["Грузополучатель и его адрес:"]
    payment = values["Всего к оплате (9)"]

    numbers.add(number)
    consignees.add(consignee)
    payment_sum += payment

print("\nКоличество файлов:", len(files),
      "\nКоличество номеров:", len(numbers),
      "\nКоличество грузополучателей:", len(consignees))

print("\nНомера:", numbers)
print("\nГрузополучатели:", consignees)
print("\n\nОбщая сумма:", payment_sum)




