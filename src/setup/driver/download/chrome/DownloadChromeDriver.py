import os
import zipfile
from abc import ABC, abstractmethod
from logging import Logger

import wget

from src.common.FileUtil import remove_all_in_folder
from src.common.ThreadLocalLogger import get_current_logger
from src.setup.driver.download.DownloadDriver import DownloadDriver
from src.setup.driver.query.DriverInfoQuery import DriverInfoQuery
from src.setup.packaging.path.PathResolvingService import PathResolvingService

PATH_TO_DRIVER = PathResolvingService.get_instance().resolve('chrome_driver')
PREFIX_DRIVER_NAME = 'chromedriver-'


class DownloadChromeDriver(DownloadDriver, ABC):
    def __init__(self, driver_info_query: DriverInfoQuery):
        super().__init__(driver_info_query)

    def download_and_place_suitable_version_driver(self):
        logger: Logger = get_current_logger()

        destination_path: str = self.get_expected_driver_abs_path()

        platform_acronym = self.get_platform_acronym()

        if not os.path.exists(destination_path):
            base_driver_version: str = self.driver_info_query.get_base_version_from_local()
            specific_version: str = self.driver_info_query.get_specific_version_from_origin(base_driver_version)

            download_url: str
            extracted_folder: str
            if int(base_driver_version) < 115:
                download_url = f"https://chromedriver.storage.googleapis.com/{specific_version}/chromedriver_{platform_acronym}32.zip"
                extracted_folder = ''
            else:
                download_url = (
                    f"https://storage.googleapis.com/chrome-for-testing-public/{specific_version}/{platform_acronym}64/chromedriver-{platform_acronym}64.zip")
                extracted_folder = f'chromedriver-{platform_acronym}64'

            logger.info('Downloading chrome driver {} is complete'.format(specific_version))
            if not os.path.exists(PATH_TO_DRIVER):
                os.makedirs(PATH_TO_DRIVER)

            latest_driver_zip = wget.download(url=download_url, out=os.path.join(PATH_TO_DRIVER, 'chromedriver.zip'))
            with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
                zip_ref.extractall(path=PATH_TO_DRIVER)

            full_path_extracted_folder: str = os.path.join(PATH_TO_DRIVER, extracted_folder)

            os.remove(latest_driver_zip)
            os.rename(os.path.join(full_path_extracted_folder, 'chromedriver' + self.get_driver_extension()),
                      destination_path)
            os.chmod(destination_path, 0o777)

            remove_all_in_folder(full_path_extracted_folder)
            os.rmdir(full_path_extracted_folder)
            logger.info('Chrome driver will be placed at {} for further operations'.format(destination_path))

    def get_expected_driver_abs_path(self) -> str:
        base_version: str = self.driver_info_query.get_base_version_from_local()
        return str(os.path.join(PATH_TO_DRIVER, PREFIX_DRIVER_NAME + base_version + self.get_driver_extension()))

    @abstractmethod
    def get_driver_extension(self) -> str:
        pass

    @abstractmethod
    def get_platform_acronym(self) -> str:
        pass
