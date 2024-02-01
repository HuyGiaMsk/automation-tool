import time
from logging import Logger
from typing import Callable

from selenium.webdriver import Keys
from selenium import webdriver
from selenium.webdriver.common.by import By

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ThreadLocalLogger import get_current_logger
from src.task.AutomatedTask import AutomatedTask


class Download_bill_maersk(AutomatedTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['username', 'password', 'download.folder', 'excel.path', 'excel.sheet',
                                     'excel.column.bill', 'country', 'office', 'customer_code', 'excel.column.doc.type',
                                     'excel.column.status']
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

        bills: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                 self._settings['excel.sheet'],
                                                                 self._settings['excel.column.bill'])
        if len(bills) == 0:
            logger.error('Input booking id list is empty ! Please check again')

        self.perform_mainloop_on_collection(bills, Download_bill_maersk.operation_on_each_element)

    def __login(self):

        logger: Logger = get_current_logger()

        user_name: str = self._settings['username']
        country: str = self._settings['country']
        office: str = self._settings['office']
        customer_code: str = self._settings['customer_code']

        self._click_when_element_present(by=By.CSS_SELECTOR, value='div.coi-banner__page-footer button:nth-child(3)')
        logger.info('Accepted all cookies')

        try_login_count: int = 1
        try:
            if try_login_count > 10:
                raise Exception
            self._type_when_element_present(by=By.CSS_SELECTOR, value='mc-input#username input', content=user_name)
            logger.info('inputted user name')

            self._type_when_element_present(by=By.CSS_SELECTOR, value='mc-input#username input', content=Keys.TAB)

            time.sleep(1)

            self._click_and_wait_navigate_to_other_page(by=By.ID, value='login-submit-button')

            time.sleep(1)
            try_login_count += 1

        except:
            logger.info('cannot get elements')

        self._type_when_element_present(by=By.CSS_SELECTOR,
                                        value='fieldset div:nth-child(2) mc-c-typeahead[data-cy=country-input]',
                                        content=country)
        self._click_when_element_present(by=By.CSS_SELECTOR, value='mc-c-list.suggestions-list')

        self._type_when_element_present(by=By.CSS_SELECTOR,
                                        value='fieldset div:nth-child(3) mc-c-typeahead[data-cy=office-input]',
                                        content=office)

        self._type_when_element_present(by=By.ID, value='customer-code', content=customer_code)

        self._click_and_wait_navigate_to_other_page(by=By.CSS_SELECTOR, value='mc-button#customer-code-submit')

    def operation_on_each_element(self, bill):

        logger: Logger = get_current_logger()
        if self.terminated is True:
            return

        with self.pause_condition:

            while self.paused:
                self.pause_condition.wait()

            if self.terminated is True:
                return

        logger.info("Processing bill : " + bill)
