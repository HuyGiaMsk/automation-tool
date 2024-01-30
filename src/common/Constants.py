import os

SOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(os.path.dirname(SOURCE_DIR))
LOG_FOLDER = os.path.join(ROOT_DIR, 'log')

PATH_TO_DRIVER = os.path.join(ROOT_DIR, 'chrome_driver')
PREFIX_DRIVER_NAME = 'chromedriver-'
DRIVER_EXTENSION = '.exe'

ZIP_EXTENSION = '.zip'
