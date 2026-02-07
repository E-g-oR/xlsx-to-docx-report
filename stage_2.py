import logging
import pandas as pd
from pathlib import Path


def stage_2(df: pd.DataFrame, target_dir: Path):
    print("   Подготовка данных и сопоставление адресов...")
    
    # 1. Загружаем список эталонных адресов
    addresses_source = target_dir / "адреса.xlsx"
    if not addresses_source.exists():
        print(f"🛑 Файл не найден: {addresses_source}")
        print("   Пропускаю шаг 2: статистика по адресам")
        print("   Перехожу к шагу 3: группировка по папкам")
        logging.info(f"!!! Файл не найден: {addresses_source}")
        logging.info(f"Пропускаю шаг 2: статистика по адресам")
        logging.info(f"Перехожу к шагу 3: группировка по папкам")
        return
    addresses_df = pd.read_excel(addresses_source, engine="calamine")
    valid_addresses = addresses_df["Уникальные адреса"].unique()

    # 2. Очистка данных перед расчетами
    df = df.copy()
    df["total_sum"] = pd.to_numeric(df["total_sum"], errors='coerce').fillna(0)
    df["doc_number"] = df["doc_number"].astype(str)
    df["date"] = df["date"].astype(str)

    # 3. Фильтруем df, оставляя только те адреса, что есть в списке
    filtered_df = df[df["client_address"].isin(valid_addresses)]

    # 4. Группируем данные (агрегация заменяет цикл)
    final_lookup = filtered_df.groupby("client_address").agg({
        "doc_number": lambda x: ", ".join(x.unique()),
        "date": lambda x: ", ".join(x.unique()),
        "total_sum": "sum"
    }).reset_index()

    # Переименовываем столбцы для красоты
    final_lookup.columns = ["Адрес", "Номера УПД", "Даты", "Сумма"]

    # 5. Сохранение
    
    # target_dir.mkdir(parents=True, exist_ok=True)
    print(f"   Создаю папку \"output\"")
    logging.info(f"Создаю папку \"output\"")
    
    output_folder = target_dir / "output" 
    output_folder.mkdir(parents=True, exist_ok=True)

    output_path = output_folder / "данные-по-адресам.xlsx"

    print(f"   Сохраняю файл...")
    logging.info(f"Сохраняю файл {output_path}...")

    final_lookup.to_excel(output_path, index=False)
    
    print(f"   Файл сохранен: {output_path}")