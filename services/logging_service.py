import logging
from pathlib import Path


LOG_PATH = Path("logs/app.log")


def get_app_logger(name="etl.app"):
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.propagate = False

    if not any(
        isinstance(handler, logging.FileHandler)
        and Path(handler.baseFilename) == LOG_PATH.resolve()
        for handler in logger.handlers
    ):
        handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
        logger.addHandler(handler)

    return logger


def log_event(message, **fields):
    logger = get_app_logger()
    if fields:
        detail = " ".join(f"{key}={value}" for key, value in fields.items())
        logger.info("%s %s", message, detail)
    else:
        logger.info(message)
