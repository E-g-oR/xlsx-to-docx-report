from pathlib import Path
from typing import Iterable
import openpyxl
import logging
from pathlib import Path
from datetime import datetime
import traceback

ANCHORS = {
    "Счет-фактура №": "NEXT_RIGHT",
    "Грузополучатель и его адрес:": "NEXT_RIGHT",
    "Всего к оплате (9)": "LAST_RIGHT",
}

def setup_logging(app_dir: Path) -> Path:
    logs_dir = app_dir / "logs"
    runs_dir = logs_dir / "runs"
    runs_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    run_log = runs_dir / f"{ts}.log"
    latest_log = logs_dir / "latest.log"

    # logging в файл + в latest (одной настройкой проще сделать два handler-а)
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    h1 = logging.FileHandler(run_log, encoding="utf-8")
    h1.setFormatter(fmt)
    logger.addHandler(h1)

    h2 = logging.FileHandler(latest_log, mode="w", encoding="utf-8")  # перезапись
    h2.setFormatter(fmt)
    logger.addHandler(h2)

    logging.info("Start. Run log: %s", run_log.name)
    return run_log


def rotate_logs(app_dir: Path, keep: int = 30):
    runs_dir = app_dir / "logs" / "runs"
    logs = sorted(runs_dir.glob("*.log"), key=lambda p: p.stat().st_mtime, reverse=True)
    for p in logs[keep:]:
        try:
            p.unlink()
        except Exception:
            pass

app_dir = Path(__file__).resolve().parent  # рядом с exe (в onefile может быть иначе, см. ниже)
run_log = setup_logging(app_dir)
rotate_logs(app_dir, keep=30)

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

try:
    # твоя логика
    logging.info("Processing started")
    logging.info("Выбираю файлы .xlsx, начинающиеся с \"УПД\":")
    files = sorted(p.name for p in Path(".").glob("УПД*.xlsx"))

    numbers = set()
    consignees = set()
    payment_sum = 0

    required = set(ANCHORS.keys())

    for file in files:
        logging.info(f"Читаю файл: {file}")
        sheet = get_sheet_from_file(file)

        found = get_anchors_coordinates(ANCHORS.keys(), sheet)
        missing_anchors = required - found.keys()
        if missing_anchors:
            msg = ", ".join(sorted(missing_anchors))
            logging.warning(f"Не найдены якоря: {msg}" )
            continue


        values = {}
        for label, rule_name in ANCHORS.items():
            if label not in found:
                continue
            coords = found[label]
            values[label] = rules[rule_name](coords, sheet)

        missing_values = [k for k in required if values.get(k) is None]
        if missing_values:
            msg = ", ".join(sorted(missing_values))
            logging.warning(f"Не найдены значения для: {msg}")
            continue


        number = values["Счет-фактура №"]
        consignee = values["Грузополучатель и его адрес:"]
        payment = values["Всего к оплате (9)"]

        numbers.add(number)
        consignees.add(consignee)
        payment_sum += payment
except Exception as e:
    logging.exception("Fatal error: %s", e)  # автоматически пишет traceback
    raise

logging.info(f"Количество файлов: {len(files)}")
logging.info(f"Количество номеров: {len(numbers)}")
logging.info(f"Количество грузополучателей: {len(consignees)}")

logging.info(f"Номера: {numbers}")
logging.info(f"Грузополучатели: {consignees}")
logging.info(f"Общая сумма: {payment_sum}")




