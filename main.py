import logging
from pathlib import Path
from utils import get_next_right_value, get_last_value_in_row, get_sheet_from_file, get_anchors_coordinates
from setup_logs import setup_logs
from docx import Document


ANCHORS = {
    "Счет-фактура №": "NEXT_RIGHT",
    "Грузополучатель и его адрес:": "NEXT_RIGHT",
    "Всего к оплате (9)": "LAST_RIGHT",
}

rules = {
    "NEXT_RIGHT": get_next_right_value,
    "LAST_RIGHT": get_last_value_in_row
}

def main():
    setup_logs()
    logging.info("Start")
    try:
        # твоя логика
        logging.info("Processing started")
        logging.info("Выбираю файлы .xlsx, начинающиеся с \"УПД\":")
        files = sorted(p.name for p in Path("./assets").glob("УПД*.xlsx"))

        numbers = set()
        consignees = set()
        payment_sum = 0

        required = set(ANCHORS.keys())

        for file in files:
            logging.info(f"Читаю файл: {file}")
            sheet = get_sheet_from_file(f"./assets/{file}")

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
    

        logging.info(f"Количество файлов: {len(files)}")
        logging.info(f"Количество номеров: {len(numbers)}")
        logging.info(f"Количество грузополучателей: {len(consignees)}")

        numbers_formatted = ", ".join(sorted(numbers))
        # consignees_formatted = "\n– ".join(sorted(consignees))
        
        logging.info(f"Номера: {numbers_formatted}")
        logging.info(f"Грузополучатели: {consignees}")
        logging.info(f"Общая сумма: {payment_sum}")

        # write to docx file
        logging.info("Создаю документ docx...")
        document = Document()

        logging.info("Пишу параграф с номерами...")
        document.add_paragraph(f"Номера: {numbers_formatted}")

        logging.info("Пишу параграф с грузополучателями...")
        document.add_paragraph("Грузополучатели:")
        for c in consignees:
            document.add_paragraph(f"– {c}")

        logging.info("Пишу параграф с общей суммой...")
        document.add_paragraph(f"Общая сумма: {payment_sum}")
        
        logging.info("Сохраняю документ...")
        Path("./output").mkdir(exist_ok=True)
        document.save("./output/report.docx")

        logging.info("Готово.")
    
    except Exception as e:
            logging.exception("Fatal error: %s", e)  # автоматически пишет traceback
            raise

if __name__ == "__main__":
    main()



