import os
import time
from datetime import datetime
from enum import Enum
from logging import Logger
from typing import Callable

from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.StringUtil import get_row_index_from_excel_cell_format
from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.AutomatedTask import AutomatedTask


class BookingToInfoIndex(Enum):
    BILL_INDEX_IN_TUPLE = 0
    TYPE_INDEX_IN_TUPLE = 1


class Download_Bill_Maersk(AutomatedTask):
    bill_to_info = {}

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

        self._excel_provider = XlwingProvider()
        self.bill_type_to_download_code: dict[str, str] = {
            'certifiedTrueCopy': 'CertifiedTrueCopy',
            'waybill': 'Waybill',
            'verifyCopy': 'VerifyCopy'
        }

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
        self._driver.get('https://www.maersk.com/hub/')

        current_url: str = self._driver.current_url

        if not current_url.endswith('hub/'):

            find_element_login: WebElement = self._get_when_element_present(by=By.CSS_SELECTOR,
                                                                            value='#maersk-app main form mc-input#username')
            if find_element_login is not None:
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

        excel_row_index: int = get_row_index_from_excel_cell_format(self._settings['excel.column.bill'])
        for bill in bills:

            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

            logger.info("Processing booking : " + bill)
            self.__navigate_and_download(bill, excel_row_index)
            excel_row_index += 1

        self._driver.close()

        workbook_path = self._settings['excel.path']
        workbook = self._excel_provider.get_workbook(workbook_path)
        self._excel_provider.close(workbook=workbook)
        self._excel_provider.quit_session()

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

        try_accept_cookies: int = 0
        try:
            if try_accept_cookies > 10:
                raise Exception

            self._click_when_element_present(by=By.CSS_SELECTOR,
                                             value='div.coi-banner__page-footer button:nth-child(3)')
            logger.info('Accepted all cookies')
            try_accept_cookies += 1
        except:
            logger.info('All cookies accepted before')

        self._get_when_element_present(by=By.CSS_SELECTOR, value='#maersk-app main form mc-input#username')

        username_element: WebElement = self._driver.execute_script(
            'return document.querySelector(\'#maersk-app main form mc-input#username\').shadowRoot.querySelector(\'input\')')
        username_element.send_keys(user_name)

        username_element.send_keys(Keys.TAB)
        logger.info('inputted user name')

        self._click_and_wait_navigate_to_other_page(by=By.ID, value='login-submit-button')
        logger.info('clicked button - need your access')

        time.sleep(2)
        self._click_when_element_present(by=By.CSS_SELECTOR, value='div.coi-banner__page-footer button:nth-child(3)')

        # try to input customer data
        # input country
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
        except:
            logger.info('Cannot loggin in Customer Country')

        # input office
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
        except:
            logger.info('Cannot loggin in Customer Office')

        # input customer code
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
        except:
            logger.info('Cannot loggin in Customer Code')

        # try to submit button
        try_click_count: int = 0
        try:
            if try_click_count > 50:
                raise Exception

            current_url: str = self._driver.current_url
            if current_url.endswith('hub/'):
                pass

            self._click_when_element_present(by=By.ID, value='customer-code-submit')

            try_click_count += 1
            time.sleep(1)
        except:
            logger.info('Cannot seting account')

        logger.info('Done setting account, progressing to navigate and download Bill')

    def __navigate_and_download(self, bill: str, excel_row_index: int):

        logger: Logger = get_current_logger()

        logger.info("Processing bill : " + bill)

        self._driver.get('https://www.maersk.com/shipment-details/{}/documents'.format(bill))

        shipment_content = self._try_to_get_if_element_present(by=By.CSS_SELECTOR,
                                                               value='header.shipment-summary section.shipment-summary__content',
                                                               waiting_time=4)

        path_to_excel_contain_pdfs_content = self._settings['excel.path']
        sheet_name = self._settings['excel.sheet']

        if shipment_content is None:
            self.filling_value(workbook_path=path_to_excel_contain_pdfs_content, sheet_name=sheet_name,
                               row_index=excel_row_index, status='Not yet')
            return

        option_documents: WebElement = self.find_matched_option_shadow_bill_msk(by=By.CSS_SELECTOR,
                                                                                list_options_selector='#main #maersk-app mc-tab-bar div:nth-child(1) div.documents-list__group div.tasks-documents-card',
                                                                                search_keyword=self.bill_to_info.get(
                                                                                    bill))

        if option_documents is None:
            self.filling_value(workbook_path=path_to_excel_contain_pdfs_content, sheet_name=sheet_name,
                               row_index=excel_row_index, status='Missing')
            return

        # click button download
        # Dậu đổ bìm leo ver1
        button_inside_shadow_root = self._driver.execute_script('''
            return arguments[0].shadowRoot.querySelector('button');
        ''', option_documents.find_element(By.CSS_SELECTOR, 'mc-button'))

        # Dậu đổ bìm leo ver2
        self._driver.execute_script("arguments[0].click();", button_inside_shadow_root)

        bill_type = option_documents.find_element(by=By.CSS_SELECTOR, value='h3').get_attribute('data-test')
        download_code = self.bill_type_to_download_code.get(bill_type)
        file_path: str = '{}\{}_{}.pdf'.format(self._settings['download.folder'], bill, download_code)

        self._wait_download_file_complete(file_path=file_path)

        self.filling_value(workbook_path=path_to_excel_contain_pdfs_content, sheet_name=sheet_name,
                           row_index=excel_row_index, status='Done')

        root_dir = os.path.dirname(file_path)
        source = file_path
        des = os.path.join(root_dir, '{}.pdf'.format(bill))

        if os.path.exists(des):
            os.remove(des)

        os.rename(source, des)

    def find_matched_option_shadow_bill_msk(self: object, by: str, list_options_selector: str,
                                            search_keyword: str) -> WebElement:
        options: list[WebElement] = self._driver.find_elements(by=by, value=list_options_selector)
        finding_option = None
        for current_option in options:
            current_inner_text = current_option.find_element(by=By.CSS_SELECTOR, value='h3').get_attribute('data-test')
            if current_inner_text.lower() in search_keyword.lower():
                finding_option = current_option
                break
        return finding_option

    def filling_value(self, workbook_path, sheet_name: str, row_index: int, status: str):
        workbook = self._excel_provider.get_workbook(workbook_path)
        sheet = workbook.sheets[sheet_name]
        self._excel_provider.change_value_at(sheet, row_index, 3, status)
        self._excel_provider.save(workbook)
