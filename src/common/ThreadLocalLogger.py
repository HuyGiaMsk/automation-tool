import os.path
import logging
import sys
import threading

from logging import Logger, FileHandler, StreamHandler, Formatter
from logging.handlers import RotatingFileHandler
from typing import TextIO
from src.common.Constants import LOG_FOLDER

thread_local_logger = threading.local()


def create_thread_local_logger(class_name: str, thread_uuid: str, logging_console_level: int = logging.INFO) -> Logger:
    thread_local_logger.logger = create_logger(class_name=class_name,
                                               thread_uuid=thread_uuid,
                                               logging_console_level=logging_console_level)
    return thread_local_logger.logger


def get_current_logger() -> Logger:
    if not hasattr(thread_local_logger, 'logger'):
        default_logger: Logger = logging.getLogger('DefaultLogger')

        if default_logger.level == logging.NOTSET:
            thread_local_logger.logger = create_thread_local_logger(class_name='DefaultLogger',
                                                                    thread_uuid='DefaultUUID',
                                                                    logging_console_level=logging.INFO)
        else:
            thread_local_logger.logger = default_logger

    return thread_local_logger.logger


def create_logger(class_name: str, thread_uuid: str, logging_console_level: int = logging.INFO) -> Logger:
    class_name: str = os.path.splitext(os.path.basename(class_name))[0]
    created_logger: Logger = logging.getLogger(class_name)

    if not os.path.exists(LOG_FOLDER):
        os.mkdir(LOG_FOLDER)

    file_handler: FileHandler = RotatingFileHandler(filename=os.path.join(LOG_FOLDER, '{}.log'.format(class_name)),
                                                    maxBytes=1024 * 1000 * 10,
                                                    backupCount=3)
    file_handler.setLevel(logging.INFO)

    console_handler: StreamHandler[TextIO] = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging_console_level)

    formatter: Formatter = logging.Formatter(
        '{} - %(asctime)s - %(levelname)s - %(filename)s %(funcName)s#%(lineno)d: %(message)s'.format(thread_uuid))
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    created_logger.addHandler(file_handler)
    created_logger.addHandler(console_handler)
    created_logger.setLevel(logging.INFO)

    return created_logger
