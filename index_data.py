import pandas as pd
import numpy as np
from pathlib import Path


data ={"Счет-фактура №": 0, 
       "Грузополучатель и его адрес:": 0, 
       "Всего к оплате (9)": -1,
       "Покупатель:": 0,
       "ИНН/КПП покупателя:": 0
       }

def find_anchor_coords(anchor: str, data_array: np.ndarray):
    # Проверяем совпадение. Получаем булеву матрицу
    mask = (data_array == anchor)
    
    # Ищем индексы True
    coords = np.argwhere(mask) # argwhere надежнее np.where для поиска координат

    if coords.size > 0:
        # coords — это список найденных точек [[row, col], [row, col]...]
        return coords[0][0], coords[0][1]
    return None

def get_value_right(df: pd.DataFrame, row: int, col: int, index = 0):
    row_slice = df.iloc[row, col + 1:]
    value = row_slice.dropna().iloc[index]
    return value

def parse_file(file_name: str, folder_path: Path):
    print(f"Reading ({file_name}) file...")

    file_path = folder_path / file_name
    df = pd.read_excel(file_path, engine="calamine", header=None)
    data_array = df.values.astype(str)
    result = {}
    for key, index in data.items():
            coords = find_anchor_coords(key, data_array)
            if not coords:
                continue
            row, col = coords
            value = get_value_right(df, row, col, index)
            result[key] = value
    return result

def index_data(folder_path: Path):
    # folder_path = Path("./assets")
    files = sorted(p.name for p in folder_path.glob("УПД*.xlsx"))
    print("Количество:", len(files))
    all_data = [{"file": f, **parse_file(f, folder_path)} for f in files]
    df = pd.DataFrame(all_data)
    return df
    # 1. Список уникальных номеров УПД
    # unique_numbers = df["Счет-фактура №"].dropna().unique().tolist()

    # # 2. Список уникальных адресов
    # unique_addresses = df["Грузополучатель и его адрес:"].dropna().unique().tolist()

    # # 3. Общая сумма по всем документам
    # # errors='coerce' превратит мусор в NaN, чтобы sum() не сломался
    # total_sum = pd.to_numeric(df["Всего к оплате (9)"], errors='coerce').sum()

    # print("Done:")
    # # Вывод для проверки
    # print(f"Найдено номеров: {len(unique_numbers)}")
    # print(f"Найдено адресов: {len(unique_addresses)}")
    # print(f"Общая сумма: {total_sum}")

