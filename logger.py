import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

LOG_FILE_NAME = "OutlookDBBuilder.log"
LOG_MAX_BYTES = 5 * 1024 * 1024
LOG_BACKUP_COUNT = 3
_logger = None

def setup_logger(base_dir: Path):
    global _logger
    if _logger is not None:
        return _logger
    logger = logging.getLogger("OutlookDBBuilder")
    logger.setLevel(logging.INFO); logger.propagate = False; logger.handlers.clear()
    try:
        h = RotatingFileHandler(Path(base_dir)/LOG_FILE_NAME, maxBytes=LOG_MAX_BYTES,
                                backupCount=LOG_BACKUP_COUNT, encoding="utf-8")
        h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s", "%Y-%m-%d %H:%M:%S"))
        logger.addHandler(h)
    except Exception:
        logger.addHandler(logging.NullHandler())
    _logger = logger
    return logger

def get_logger():
    return _logger or logging.getLogger("OutlookDBBuilder")
def log_info(m):
    try: get_logger().info(str(m))
    except Exception: pass
def log_warning(m):
    try: get_logger().warning(str(m))
    except Exception: pass
def log_error(m):
    try: get_logger().error(str(m))
    except Exception: pass
def log_exception(m):
    try: get_logger().exception(str(m))
    except Exception: pass
