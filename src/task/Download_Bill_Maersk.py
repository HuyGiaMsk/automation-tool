import time
from logging import Logger
from typing import Callable

from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ThreadLocalLogger import get_current_logger
from src.task.AutomatedTask import AutomatedTask


class Download_Bill_Maersk(AutomatedTask):

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

        self._driver.get('https://www.maersk.com/portaluser/login')

        logger.info('Try to login')
        self.__login()
        logger.info("Login successfully")

        bills: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                 self._settings['excel.sheet'],
                                                                 self._settings['excel.column.bill'])
        if len(bills) == 0:
            logger.error('Input booking id list is empty ! Please check again')

        self.perform_mainloop_on_collection(bills, Download_Bill_Maersk.operation_on_each_element)

    def __login(self):

        logger: Logger = get_current_logger()

        user_name: str = self._settings['username']
        country: str = self._settings['country']
        office: str = self._settings['office']
        customer_code: str = self._settings['customer_code']

        self._click_when_element_present(by=By.CSS_SELECTOR, value='div.coi-banner__page-footer button:nth-child(3)')
        logger.info('Accepted all cookies')

        self._get_when_element_present(by=By.CSS_SELECTOR, value='#maersk-app main form mc-input#username')

        username_element: WebElement = self._driver.execute_script(
            'return document.querySelector(\'#maersk-app main form mc-input#username\').shadowRoot.querySelector(\'input\')')
        username_element.send_keys(user_name)

        username_element.send_keys(Keys.TAB)
        logger.info('inputted user name')

        time.sleep(1)

        mc_button = self._driver.execute_script('return document.querySelector(\'mc-button#login-submit-button\')')
        shadow_root = self._driver.execute_script('return arguments[0].shadowRoot', mc_button)
        button_inside_shadow_dom = self._driver.execute_script(
            'return arguments[0].querySelector(\'button#button.items-center\')', shadow_root)

        button_inside_shadow_dom.click()

        time.sleep(1)

        logger.info('clicked button - need your access')

        self._wait_navigating_to_other_page_complete(previous_url='https://www.maersk.com/portaluser/login',
                                                     expected_end_with='https://www.maersk.com/portaluser/select-customer')

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
