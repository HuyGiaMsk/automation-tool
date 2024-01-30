import time
from logging import Logger
from typing import Callable

from selenium.webdriver import Keys
from selenium.webdriver.common.by import By

from src.common.ThreadLocalLogger import get_current_logger
from src.task.AutomatedTask import AutomatedTask


class Download_bill_maersk(AutomatedTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['username', 'password', 'download.folder', 'excel.path', 'excel.sheet',
                                    'excel.column.booking',
                                    'excel.column.so', 'excel.column.becode']
        return mandatory_keys

    def automate(self) -> None:
        logger: Logger = get_current_logger()
        logger.info(
            "---------------------------------------------------------------------------------------------------------")
        logger.info("Start processing")

        self._driver.get('https://accounts.maersk.com/ocean-maeu/auth/login')

        logger.info('Try to login')
        self.__login()
        logger.info("Login successfully")




    def __login(self):
        user_name: str = self._settings['username']
        pass_word: str = self._settings['password']

        self._click_when_element_present(by=By.CSS_SELECTOR, value='div.coi-banner__page-footer button:nth-child(3)')
        self._type_when_element_present(by=By.ID, value='username-input', content=user_name)
        self._type_when_element_present(by=By.ID, value='username-input', content= Keys.TAB)
        self._click_when_element_present(by=By.ID, value='login-submit-button')
        time.sleep(1)

        self._driver.switch_to.window(self._driver.window_handles[-1])