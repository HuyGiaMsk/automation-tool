import time
from datetime import datetime
from enum import Enum
from logging import Logger
from typing import Callable

from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ThreadLocalLogger import get_current_logger
from src.task.AutomatedTask import AutomatedTask


class BookingToInfoIndex(Enum):
    BILL_INDEX_IN_TUPLE = 0
    TYPE_INDEX_IN_TUPLE = 1


class Download_Bill_Maersk(AutomatedTask):
    bill_to_info = {}

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

        type_bill: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                     self._settings['excel.sheet'],
                                                                     self._settings['excel.column.doc.type'])
        if len(bills) == 0:
            logger.error('Input booking id list is empty ! Please check again')

        if len(bills) != len(type_bill):
            raise Exception("Please check your input data length of bills, type_bill are not equal")

        index: int = 0
        for bill in bills:
            self.bill_to_info[bill] = (type_bill[index])
            index += 1

        last_bill: str = ''
        for bill in bills:
            logger.info("Processing booking : " + bill)
            self.__navigate_and_download(bill)
            last_bill = bill

        self._driver.close()
        logger.info(
            "---------------------------------------------------------------------------------------------------------")
        logger.info("End processing")
        logger.info("It ends at {}. Press any key to end program...".format(datetime.now()))
        # self.perform_mainloop_on_collection(bills, Download_Bill_Maersk.operation_on_each_element)

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

        # _______________________________________________
        # time.sleep(5)
        # button_element: WebElement = self._driver.execute_script(
        #     'return document.querySelector(\'#maersk-app main form div.mt-40 mc-button#login-submit-button\').shadowRoot.querySelector(\'button\')')
        # button_element.click()
        # ___________________________________________________

        self._click_and_wait_navigate_to_other_page(by=By.ID, value='login-submit-button')
        logger.info('clicked button - need your access')

        time.sleep(1)
        self._click_when_element_present(by=By.CSS_SELECTOR, value='div.coi-banner__page-footer button:nth-child(3)')

        try_country_count: int = 0
        try:
            if try_country_count > 50:
                raise Exception
            country_element: WebElement = self._driver.execute_script(
                'return document.querySelector(\'#maersk-app form mc-c-typeahead[data-cy="country-input"]\')'
                '.shadowRoot.querySelector(\'input\')')
            country_element.send_keys(country)
            country_element.send_keys(Keys.ARROW_DOWN)
            country_element.send_keys(Keys.ENTER)
            country_element.send_keys(Keys.TAB)

            try_country_count += 1
            time.sleep(1)
        except:
            logger.info('Cannot loggin in Customer Country')

        try_office_count: int = 0
        try:
            if try_office_count > 50:
                raise Exception
            office_element: WebElement = self._driver.execute_script(
                'return document.querySelector(\'#maersk-app form mc-c-typeahead[data-cy="office-input"]\')'
                '.shadowRoot.querySelector(\'input\')')
            office_element.send_keys(office)
            office_element.send_keys(Keys.ARROW_DOWN)
            office_element.send_keys(Keys.ENTER)
            office_element.send_keys(Keys.TAB)

            try_office_count += 1
            time.sleep(1)
        except:
            logger.info('Cannot loggin in Customer Office')

        try_customer_count: int = 0
        try:
            if try_customer_count > 50:
                raise Exception
            customer_code_element: WebElement = self._driver.execute_script(
                'return document.querySelector(\'#maersk-app form mc-input[data-cy="customer-code-input"]\')'
                '.shadowRoot.querySelector(\'input\')')
            customer_code_element.send_keys(customer_code)
            customer_code_element.send_keys(Keys.ARROW_DOWN)
            customer_code_element.send_keys(Keys.ENTER)
            customer_code_element.send_keys(Keys.TAB)

            try_customer_count += 1
            time.sleep(1)
        except:
            logger.info('Cannot loggin in Customer Code')
        self._click_and_wait_navigate_to_other_page(by=By.ID, value='customer-code-submit')
        time.sleep(100000)

    def __navigate_and_download(self, bill):

        logger: Logger = get_current_logger()

        logger.info("Processing bill : " + bill)
        self._driver.get('https://www.maersk.com/shipment-details/{}/documents'.format(bill))

        list_element: list[WebElement] = self._driver.execute_script(
            'return document.querySelector(\'#main #maersk-app mc-tab-bar \').shadowRoot.querySelector(\'.documents-list.mb-5\')')

        self.find_matched_option_shadow(by=By.CSS_SELECTOR, list_options_selector=list_element, search_keyword='a')
