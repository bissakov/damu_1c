import logging
from datetime import datetime
from pathlib import Path

import pytz
from rich.console import Console
from rich.logging import RichHandler


def setup_logger(project_root: Path) -> Path:
    today = datetime.now(pytz.timezone("Asia/Almaty"))

    logging.Formatter.converter = lambda *args: today.timetuple()

    log_folder = project_root / "logs"

    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    log_format = (
        "%(asctime).19s %(levelname)s %(name)s %(filename)s %(funcName)s : %(message)s"
    )
    formatter = logging.Formatter(log_format)

    today_str = today.strftime("%d.%m.%y")
    year_month_folder = log_folder / today.strftime("%Y_%m")
    year_month_folder.mkdir(parents=True, exist_ok=True)

    logger_file = year_month_folder / f"{today_str}.log"

    file_handler = logging.FileHandler(logger_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    logging.getLogger("httpcore").setLevel(logging.WARNING)
    logging.getLogger("requests").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)

    logger.addHandler(file_handler)

    logger.addHandler(
        RichHandler(
            console=Console(width=255),
            omit_repeated_times=False,
            rich_tracebacks=False,
        )
    )

    return logger_file
