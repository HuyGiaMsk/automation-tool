import os
import re
from logging import Logger

from src.common.ThreadLocalLogger import get_current_logger
from src.setup.driver.query.chrome.ChromeDriverInfoQuery import ChromeDriverInfoQuery


class WindowsChromeDriverInfoQuery(ChromeDriverInfoQuery):

    def get_base_version_from_local(self) -> str:
        logger: Logger = get_current_logger()
        base_number_version: str = ''
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

        if base_number_version == '':
            message: str = f'Error retrieving Chrome version'
            logger.error(message)
            raise Exception(message)

        logger.info('Local machine chrome version used is {}'.format(base_number_version))
        return base_number_version
