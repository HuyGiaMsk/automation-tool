import os
import sys


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


ASB_CURRENT_FILE_PATH = resource_path(os.path.abspath(__file__))
COMMON_DIR = os.path.dirname(ASB_CURRENT_FILE_PATH)
SOURCE_DIR = os.path.dirname(COMMON_DIR)
ROOT_DIR = os.path.dirname(SOURCE_DIR)
LOG_DIR = os.path.join(ROOT_DIR, 'log')
TASKS_DIR = os.path.join(SOURCE_DIR, 'task')

ZIP_EXTENSION = '.zip'
