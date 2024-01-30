import os

import requests
import zipfile
import wget
import re

from logging import Logger
from src.common.FileUtil import remove_all_in_folder
from src.common.Constants import PATH_TO_DRIVER, PREFIX_DRIVER_NAME, DRIVER_EXTENSION
from src.common.ThreadLocalLogger import get_current_logger


def get_latest_version_from_google(base_number_version: str) -> str:
    logger: Logger = get_current_logger()
    url: str
    if int(base_number_version) < 115:
        url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_' + str(base_number_version)
    else:
        url = 'https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_' + str(base_number_version)
    response = requests.get(url)
    specific_version: str = response.text
    logger.info('Specific chrome driver version {} is suitable for our local machine chrome version{}'.format(
        specific_version, base_number_version))
    return specific_version


def get_current_local_chrome_base_version() -> str:
    logger: Logger = get_current_logger()
    base_number_version: str = '119'
    chrome_registry = os.popen(r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version')
    replies = chrome_registry.read()
    replies = replies.split('\n')
    chrome_registry.close()

    for reply in replies:
        if 'version' in reply:
            reply = reply.strip()
            tokens = re.split(r"\s+", reply)
            full_version = tokens[len(tokens) - 1]
            tokens = full_version.split('.')
            base_number_version = tokens[0]
            break
    logger.info('Local machine chrome version used is {}'.format(base_number_version))
    return base_number_version


def place_suitable_chromedriver():
    logger: Logger = get_current_logger()
    base_driver_version: str = get_current_local_chrome_base_version()
    destination_path: str = os.path.join(PATH_TO_DRIVER, PREFIX_DRIVER_NAME + base_driver_version + DRIVER_EXTENSION)

    if not os.path.exists(destination_path):
        specific_version: str = get_latest_version_from_google(base_driver_version)

        download_url: str
        extracted_folder: str
        if int(base_driver_version) < 115:
            download_url = "https://chromedriver.storage.googleapis.com/" + specific_version + "/chromedriver_win32.zip"
            extracted_folder = ''
        else:
            download_url = ("https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/" + specific_version +
                            "/win64/chromedriver-win64.zip")
            extracted_folder = 'chromedriver-win64'

        logger.info('Downloading chrome driver {} is complete'.format(specific_version))
        if not os.path.exists(PATH_TO_DRIVER):
            os.makedirs(PATH_TO_DRIVER)

        latest_driver_zip = wget.download(url=download_url, out=os.path.join(PATH_TO_DRIVER, 'chromedriver.zip'))
        with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
            zip_ref.extractall(path=PATH_TO_DRIVER)

        full_path_extracted_folder: str = os.path.join(PATH_TO_DRIVER, extracted_folder)

        os.remove(latest_driver_zip)
        os.rename(os.path.join(full_path_extracted_folder, 'chromedriver.exe'), destination_path)
        os.chmod(destination_path, 777)

        remove_all_in_folder(full_path_extracted_folder)
        os.rmdir(full_path_extracted_folder)
        logger.info('Chrome driver will be placed at {} for further operations'.format(destination_path))


def get_full_browser_driver_path() -> str:
    base_version: str = get_current_local_chrome_base_version()
    return os.path.join(PATH_TO_DRIVER, PREFIX_DRIVER_NAME + base_version + DRIVER_EXTENSION)
