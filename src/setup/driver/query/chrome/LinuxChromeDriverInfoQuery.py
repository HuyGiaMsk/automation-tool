import os
import re
from logging import Logger

from src.common.ThreadLocalLogger import get_current_logger
from src.setup.driver.query.chrome.ChromeDriverInfoQuery import ChromeDriverInfoQuery


class LinuxChromeDriverInfoQuery(ChromeDriverInfoQuery):

    def get_base_version_from_local(self) -> str:
        logger: Logger = get_current_logger()
        base_number_version: str = ''

        try:
            chrome_version_output = os.popen('google-chrome --version').read().strip()
            logger.info(f'Chrome version output: {chrome_version_output}')
            version_match = re.search(r'\d+\.\d+\.\d+\.\d+', chrome_version_output)
            if version_match:
                full_version = version_match.group(0)
                tokens = full_version.split('.')
                base_number_version = tokens[0]

            logger.info('Local machine chrome version used is {}'.format(base_number_version))
        except Exception as e:
            message: str = f'Error retrieving Chrome version: {e}'
            logger.error(message)
            raise Exception(message)

        return base_number_version
