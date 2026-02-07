import pandas as pd
import numpy as np
from pathlib import Path
from typing import TypedDict, Optional, Literal
from pathlib import Path
import logging
from tqdm import tqdm


class DocumentData(TypedDict):
    file_path: Path
    doc_type: Literal["UPD", "INVOICE"]  # Ограничиваем только этими значениями
    total_sum: float
    client_name: Optional[str]           # Может быть str, а может быть None
    client_address: Optional[str]
    doc_number: Optional[str]
    raw_text: Optional[str]              # Техническое поле для "грязных" данных
    raw_date: Optional[str]              # Техническое поле для "грязных" данных


client_names = []
invoices_queue = []


data ={"Счет-фактура №": 0, 
       "Грузополучатель и его адрес:": 0, 
       "Всего к оплате (9)": -1,
       "Покупатель:": 0,
       "ИНН/КПП покупателя:": 0,
       "Документ об отгрузке": 0
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

def parse_file(file_path: Path, folder_path: Path) -> DocumentData:

    is_upd = "УПД" in file_path.name

    # file_path = folder_path / file_name
    
    doc: DocumentData = {
        "file_path": file_path,
        "doc_type": "UPD" if is_upd else "INVOICE",
        "total_sum": 0.0,
        "client_name": None,
        "client_address": None,
        "doc_number": None,
        "raw_text": None,
        "raw_date": None,
    }
    
    df = pd.read_excel(file_path, engine="calamine", header=None)
    data_array = df.values.astype(str)

    if is_upd:
        result = {}
        for key, index in data.items():
                coords = find_anchor_coords(key, data_array)
                if not coords:
                    continue
                row, col = coords
                value = get_value_right(df, row, col, index)
                result[key] = value
        doc["client_name"] = result["Покупатель:"]    
        doc["client_address"] = result["Грузополучатель и его адрес:"]    
        doc["doc_number"] = result["Счет-фактура №"]    
        doc["total_sum"] = result["Всего к оплате (9)"]    
        doc["raw_date"] = result["Документ об отгрузке"]    
    else:
        coords = find_anchor_coords('''Покупатель
(Заказчик):''', data_array)
        
        row, col = coords
        value = get_value_right(df, row, col)
        doc["raw_text"] = value
    return doc

def index_data(folder_path: Path):
    # folder_path = Path("./assets")
    # Собираем список из двух поисков
    files = list(folder_path.glob("УПД*.xlsx")) + list(folder_path.glob("Счет*.xlsx"))

    # Сортируем уже финальный список путей
    files = sorted(files, key=lambda p: p.name)

    
    all_data = []
    total_files = len(files)
    if total_files == 0:
        print("🛑 Ни одного нужного файла не найдено.")
        print("   Завершаю выполнение.")
        logging.info("🛑 Ни одного нужного файла не найдено.")
        logging.info("Завершаю выполнение.")
        
        return
    
    print("Количество:", total_files)
    logging.info(f"Всего файлов: {total_files}")


    # Вместо простого цикла используем tqdm как обертку над списком
    # desc — это описание, которое будет висеть слева от полоски
    pbar = tqdm(files, desc="Индексация файлов", unit="file")


    for index, f in enumerate(pbar, start=1):
        # Выводим прогресс: индекс и имя файла
        # pbar.write(f"[{index}/{total_files}] Обработка: {f.name}...")
        logging.info(f"[{index}/{total_files}] Обработка: {f.name}...")
        
        parsed_results = parse_file(f, folder_path)
        
        # Собираем словарь (аналог твоей записи с **)
        all_data.append(parsed_results)

    logging.info(f"Все файлы прочитаны и обработаны")
    # Собираем все "чистые" имена из УПД в один справочник
    known_clients = {d["client_name"] for d in all_data if d["doc_type"] == "UPD" and d["client_name"]}
    
    print(f"🔍 Найдено уникальных клиентов в УПД: {len(known_clients)}")
    logging.info(f"Справочник клиентов сформирован: {known_clients}")

    # "Лечим" счета
    print("🩹 Сопоставляю счета с найденными клиентами...")

    for d in all_data:
        if d["doc_type"] == "INVOICE" and d["raw_text"]:
            for name in known_clients:
                if name in d["raw_text"]:
                    d["client_name"] = name
                    break
    return all_data
 


