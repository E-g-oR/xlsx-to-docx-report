import logging
from pathlib import Path
from setup_logs import setup_logs
import tkinter as tk
from tkinter import filedialog
import sys
from index_data import index_data
from stage_1 import get_docx_report_for_all_UPD
from stage_2 import stage_2
from stage_3 import stage_3
import pandas as pd

def get_working_directory():
    """Открывает окно выбора папки."""
    root = tk.Tk()
    root.withdraw()  # Скрываем основное маленькое окно tkinter
    root.attributes('-topmost', True)  # Поверх всех окон
    
    selected_dir = filedialog.askdirectory(title="Выберите папку с файлами Excel")
    
    if not selected_dir:
        print("Папка не выбрана. Выход...")
        sys.exit(0)
        
    return Path(selected_dir)


def main():
    target_dir = get_working_directory()
    # target_dir = Path("./assets")

    setup_logs(target_dir)
    # logging.info("Старт")
    try:
        logging.info("Processing started")


        logging.info("🔎 Этап 0: Подготовга данных")
        all_data = index_data(target_dir)

        df = pd.DataFrame(all_data)
        df["date"] = df["raw_date"].str.extract(r'(\d{2}\.\d{2}\.\d{4})')
        logging.info("Этап 0: Готово.")
        logging.info("----------------------------")

        print("\n📋 Этап 1: Отчет для страховки")
        logging.info("📋 Этап 1: Отчет для страховки")
        get_docx_report_for_all_UPD(df, target_dir)
        logging.info("Этап 1: Готово.")
        print("✅ Этап 1: Готово.\n")

        print("🗺️  Этап 2: Данные по адресам")
        logging.info("🗺️ Этап 2: Данные по адресам")
        stage_2(df)
        logging.info("Этап 2: Готово.")
        print("✅ Этап 2: Готово.\n")

        print("📂 Этап 3: Группировка по папкам")
        logging.info("📂 Этап 3: Группировка по папкам")
        stage_3(df, target_dir)
        print("✅ Этап 3: Готово.\n")
        logging.info("✨ Процесс полностью завершен!")
        print("✨ Процесс полностью завершен!\n")
    
    except Exception as e:
            logging.exception("Fatal error: %s", e)  # автоматически пишет traceback
            raise

if __name__ == "__main__":
    main()
    

