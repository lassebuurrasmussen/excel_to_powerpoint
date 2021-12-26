import logging


def get_logger(level=logging.INFO):
    logger = logging.getLogger("readers.py")
    logger.setLevel(level)
    fh = logging.StreamHandler()
    logger.addHandler(fh)
    return logger
