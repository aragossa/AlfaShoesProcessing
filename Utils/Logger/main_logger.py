import datetime
import logging
import sys

from Utils.ConfigReader.ConfigReader import read_config

FORMATTER = logging.Formatter("%(asctime)s — %(name)s — %(levelname)s — %(funcName)s:%(lineno)d — %(message)s")


def get_console_handler():
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(FORMATTER)
    return console_handler


def get_file_handler():
    file_name = "output"
    open(f"logs/{file_name}.log", 'w').close()
    file_handler = logging.FileHandler(f"logs/{file_name}.log")
    file_handler.setFormatter(FORMATTER)
    return file_handler


def get_logger(logger_name):
    logs_level = read_config("logs.level").get("level")
    logger = logging.getLogger(logger_name)
    if logs_level == "DEBUG":
        logger.setLevel(logging.DEBUG)
    else:
        logger.setLevel(logging.INFO)
    logger.addHandler(get_console_handler())
    logger.addHandler(get_file_handler())
    logger.propagate = False
    return logger
