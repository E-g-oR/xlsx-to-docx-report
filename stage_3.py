import shutil
import pandas as pd
from pathlib import Path

def stage_3(df: pd.DataFrame, target_dir: Path):
    print("\n--- Запуск Этапа 3: Распределение файлов ---")
    
    # Группируем пути к файлам по именам клиентов
    # Мы используем .dropna(subset=['client_name']), чтобы не упасть на пустых значениях
    # Или заменяем None на "Неопознано" перед группировкой:
    df_clean = df.copy()
    df_clean['client_name'] = df_clean['client_name'].fillna("НЕОПОЗНАННО")
    
    grouped_files = df_clean.groupby("client_name")["file_path"].apply(list).to_dict()

    for client_name, file_paths in grouped_files.items():
        # 1. Формируем путь к папке клиента
        client_folder = target_dir / client_name.strip()
        
        # 2. Создаем папку (если её нет)
        client_folder.mkdir(exist_ok=True)
        
        print(f"Группа [{client_name}]: перемещаю {len(file_paths)} файл(ов)")
        
        for source_path in file_paths:
            # 3. Формируем финальный путь файла (папка клиента + имя файла)
            destination = client_folder / source_path.name
            
            try:
                # Перемещаем файл
                if source_path.exists():
                    shutil.move(str(source_path), str(destination))
                else:
                    print(f"⚠️ Файл не найден: {source_path}")
            except Exception as e:
                print(f"❌ Ошибка при перемещении {source_path.name}: {e}")

    print("\n✅ Все файлы успешно распределены по папкам.")