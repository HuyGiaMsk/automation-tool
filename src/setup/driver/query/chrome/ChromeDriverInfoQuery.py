from abc import ABC
from logging import Logger

import requests

from src.common.ThreadLocalLogger import get_current_logger
from src.setup.driver.query.DriverInfoQuery import DriverInfoQuery


class ChromeDriverInfoQuery(DriverInfoQuery, ABC):

    def get_specific_version_from_origin(self, base_version: str) -> str:
        logger: Logger = get_current_logger()
        url: str
        if int(base_version) < 115:
            url = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_' + str(base_version)
        else:
            url = 'https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_' + str(base_version)
        response = requests.get(url)
        specific_version: str = response.text
        logger.info('Specific chrome driver version {} is suitable for our local machine chrome version{}'.format(
            specific_version, base_version))

        return specific_version
