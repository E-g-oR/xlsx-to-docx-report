import logging
from pathlib import Path
from setup_logs import setup_logs
import tkinter as tk
from tkinter import filedialog
import sys
import cProfile
import pstats
from index_data import index_data
from stage_1 import get_docx_report_for_all_UPD
from stage_2 import stage_2
from stage_3 import stage_3

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
    # target_dir = get_working_directory()
    target_dir = Path("./assets")

    setup_logs(target_dir)
    logging.info("Start")
    try:
        logging.info("Processing started")
        logging.info("Выбираю файлы .xlsx, начинающиеся с \"УПД\":")

        df = index_data(target_dir)

        get_docx_report_for_all_UPD(df, target_dir)

        stage_2(df)

        stage_3()

        logging.info("Готово.")
    
    except Exception as e:
            logging.exception("Fatal error: %s", e)  # автоматически пишет traceback
            raise

if __name__ == "__main__":
    with cProfile.Profile() as pr:
        main()
    
    # Обрабатываем результаты
stats = pstats.Stats(pr)
stats.sort_stats('cumtime') # Сортируем по кумулятивному времени

# ОГРАНИЧИВАЕМ ВЫВОД:
stats.print_stats(10) # Выведет только ТОП-10 самых "тяжелых" функций



