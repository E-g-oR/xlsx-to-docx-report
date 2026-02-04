import logging
from pathlib import Path
from typing import Iterable
from datetime import datetime
import sys

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

    logging.info("Старт. Лог запуска: %s", run_log.name)
    return run_log


def rotate_logs(app_dir: Path, keep: int = 30):
    runs_dir = app_dir / "logs" / "runs"
    logs = sorted(runs_dir.glob("*.log"), key=lambda p: p.stat().st_mtime, reverse=True)
    for p in logs[keep:]:
        try:
            p.unlink()
        except Exception:
            pass



def setup_logs(target_dir: Path):
    # Теперь app_dir — это та папка, которую выбрал пользователь
    run_log = setup_logging(target_dir)
    rotate_logs(target_dir, keep=30)
    return run_log