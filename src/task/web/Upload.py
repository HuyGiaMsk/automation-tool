import os
import time
from logging import Logger
from typing import Callable

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ResourceLock import ResourceLock
from src.common.ThreadLocalLogger import get_current_logger
from src.task.WebTask import WebTask


class Upload(WebTask):
    def getCurrentPercent(self):
        pass

    def get_current_percent(self) -> float:
        pass

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self._document_folder = self._download_folder

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['username', 'password', 'download.folder', 'excel.path', 'excel.sheet',
                                     'excel.column.so', 'excel.column.becode']
        return mandatory_keys

    def automate(self):
        logger: Logger = get_current_logger()
        logger.info(
            "---------------------------------------------------------------------------------------------------------")
        logger.info("Start processing")
        self._driver.get('https://portal.damco.com/Applications/documentmanagement/')
        logger.info('Try to login')
        self.__login()
        logger.info("Login successfully")
        logger.info("Navigate to refresh page the first time")

        while True:
            current_url: str = self._driver.current_url
            if current_url.endswith('documentmanagement/'):
                break
            self._click_when_element_present(by=By.CSS_SELECTOR, value='a.DOCUMENT_MANAGEMENT')
            time.sleep(2)

        # iframe switch before clicking CNEE BECODE, check ID or Classname
        try_attempt_count: int = 0
        while True:
            try:
                if try_attempt_count > 50:
                    raise Exception('Look like we have problems with the web structure changed - '
                                    'we could not click on the option_booking ! Need to our investigation')
                iframe = self._driver.find_element(by=By.ID, value='applicationIframe')
                self._driver.switch_to.frame(iframe)
                logger.info('get iframe by id successfully')
                break

            except Exception as exception:
                # try to get iframe - demacia web :)
                logger.error(str(exception))
                while True:
                    self._driver.get('https://portal.damco.com/Applications/documentmanagement/')
                    current_url: str = self._driver.current_url
                    if current_url.endswith('documentmanagement/'):
                        break
                    self._click_when_element_present(by=By.CSS_SELECTOR, value='a.DOCUMENT_MANAGEMENT')
                    time.sleep(2)

                self._click_when_element_present(by=By.ID, value='applicationIframe')
                time.sleep(1)
                logger.info("trying to get iframe")
                try_attempt_count += 1
                continue

        # get cneebecode
        becodes: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                   self._settings['excel.sheet'],
                                                                   self._settings[
                                                                       'excel.column.becode'])
        so_numbers: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                      self._settings['excel.sheet'],
                                                                      self._settings['excel.column.so'])
        becode_to_sonumber: dict[str, str] = {}

        if not len(so_numbers) == len(becodes):
            raise Exception('be_codes and so_numbers do not have thhe same length')

        # Processign to upload
        index = 0
        for so_number in so_numbers:
            becode_to_sonumber[so_number] = becodes[index]
            index += 1

        for so_number, becode in becode_to_sonumber.items():

            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

            self._upload(so_number, becode)

        logger.info("Complete Upload")

        self._driver.close()
        logger.info(
            "---------------------------------------------------------------------------------------------------------")
        logger.info("End processing")

    def __login(self) -> None:
        username: str = self._settings['username']
        password: str = self._settings['password']

        self._type_when_element_present(by=By.ID, value='ctl00_ContentPlaceHolder1_UsernameTextBox', content=username)
        self._type_when_element_present(by=By.ID, value='ctl00_ContentPlaceHolder1_PasswordTextBox', content=password)
        self._click_and_wait_navigate_to_other_page(by=By.CSS_SELECTOR, value='input[type=submit]')

    def _upload(self, so_number: str, becode: str):
        logger: Logger = get_current_logger()

        self._click_when_element_present(by=By.ID, value='selectedClientId_chosen')

        # delete all elements not expect are present
        self._type_when_element_present(by=By.CSS_SELECTOR, value='li.search-field input[type=text]',
                                        content=Keys.SPACE)
        self._type_when_element_present(by=By.CSS_SELECTOR, value='li.search-field input[type=text]',
                                        content=Keys.BACKSPACE)
        self._type_when_element_present(by=By.CSS_SELECTOR, value='li.search-field input[type=text]',
                                        content=Keys.BACKSPACE)

        # continue working - add becode
        self._type_when_element_present(by=By.CSS_SELECTOR, value='li.search-field input[type=text]', content=becode)

        # Only accept becode already exist
        self._click_when_element_present(by=By.CSS_SELECTOR,
                                         value='div.chosen-drop ul.chosen-results li.active-result')

        # click SO Number
        self._click_when_element_present(by=By.CSS_SELECTOR, value='div.styled-select option[value=SO]')

        # input SO Number
        self._type_when_element_present(by=By.ID, value='searchForm_objectValue1', content=Keys.CONTROL + 'a')
        self._type_when_element_present(by=By.ID, value='searchForm_objectValue1', content=Keys.DELETE)
        self._type_when_element_present(by=By.ID, value='searchForm_objectValue1', content=so_number)

        # go to next page to upload
        self._click_when_element_present(by=By.ID, value='searchButton')

        logger.info('Go to next page successfully')

        self._click_when_element_present(by=By.CSS_SELECTOR, value='#template #row0 td:nth-child(6) input')
        logger.info('clicked so box')

        # self._type_when_element_present(by=By.CSS_SELECTOR, value='#moreOptionUpload button', value='#moreOptionUpload button')
        # this code is use like formal code below:
        upload_button: WebElement = self._get_element_satisfy_predicate(by=By.CSS_SELECTOR,
                                                                        element_selector='#moreOptionUpload button',
                                                                        method=expected_conditions.presence_of_element_located(
                                                                            (By.CSS_SELECTOR,
                                                                             '#moreOptionUpload button')))
        upload_button.send_keys(Keys.CONTROL + Keys.ENTER)

        logger.info('Go to next window')

        number_of_tabs = len(self._driver.window_handles)
        try_count = 0
        while number_of_tabs < 2:
            if try_count > 10:
                raise Exception('Can not invoke a new tab')
            time.sleep(1)
        self._driver.switch_to.window(self._driver.window_handles[-1])
        time.sleep(2)
        logger.info('switched to new tab')

        # switched tab 2
        folder_upload = os.path.join(self._download_folder, so_number)
        file_name_inv: str = '_INV'
        file_name_pkl: str = '_PKL'
        file_name_hbl: str = '_HBL'
        file_name_cbl: str = '_CBL'
        file_name_fcr: str = '_FCR'
        file_name_clr: str = '_CLR'
        with ResourceLock(file_path=folder_upload):
            row_index: int = 0
            for file_name in os.listdir(folder_upload):
                file_path: str = os.path.join(folder_upload, file_name)

                # INV upload
                if file_name_inv in file_name:
                    suitable_option_inv: WebElement = self.find_matched_option(by=By.CSS_SELECTOR,
                                                                               list_options_selector='#row{} td:nth-child(1) option'.format(
                                                                                   row_index),
                                                                               search_keyword='Commercial Invoice')
                    suitable_option_inv.click()
                    self._type_when_element_present(by=By.CSS_SELECTOR,
                                                    value='#row{} input[type=file]'.format(row_index),
                                                    content=file_path)
                    row_index = row_index + 1

                # PL upload
                if file_name_pkl in file_name:
                    suitable_option_pkl: WebElement = self.find_matched_option(by=By.CSS_SELECTOR,
                                                                               list_options_selector='#row{} td:nth-child(1) option'.format(
                                                                                   row_index),
                                                                               search_keyword='Packing List')
                    suitable_option_pkl.click()
                    self._type_when_element_present(by=By.CSS_SELECTOR,
                                                    value='#row{} input[type=file]'.format(row_index),
                                                    content=file_path)
                    row_index = row_index + 1

                # House Bill upload
                if file_name_hbl in file_name:
                    suitable_option_house_bill: WebElement = self.find_matched_option(by=By.CSS_SELECTOR,
                                                                                      list_options_selector='#row{} td:nth-child(1) option'.format(
                                                                                          row_index),
                                                                                      search_keyword='House Bill of Lading')
                    suitable_option_house_bill.click()
                    self._type_when_element_present(by=By.CSS_SELECTOR,
                                                    value='#row{} input[type=file]'.format(row_index),
                                                    content=file_path)
                    row_index = row_index + 1

                # Carrier Bill upload
                if file_name_cbl in file_name:
                    suitable_option_bill: WebElement = self.find_matched_option(by=By.CSS_SELECTOR,
                                                                                list_options_selector='#row{} td:nth-child(1) option'.format(
                                                                                    row_index),
                                                                                search_keyword='Bill of Lading')
                    suitable_option_bill.click()
                    self._type_when_element_present(by=By.CSS_SELECTOR,
                                                    value='#row{} input[type=file]'.format(row_index),
                                                    content=file_path)
                    row_index = row_index + 1

                # FCR upload
                if file_name_fcr in file_name:
                    suitable_option_fcr: WebElement = self.find_matched_option(by=By.CSS_SELECTOR,
                                                                               list_options_selector='#row{} td:nth-child(1) option'.format(
                                                                                   row_index),
                                                                               search_keyword='Forwarders Cargo Receipt')
                    suitable_option_fcr.click()
                    self._type_when_element_present(by=By.CSS_SELECTOR,
                                                    value='#row{} input[type=file]'.format(row_index),
                                                    content=file_path)
                    row_index = row_index + 1

                # CLR
                if file_name_fcr in file_name:
                    suitable_option_clr: WebElement = self.find_matched_option(by=By.CSS_SELECTOR,
                                                                               list_options_selector='#row{} td:nth-child(1) option'.format(
                                                                                   row_index),
                                                                               search_keyword='Container Load Result')
                    suitable_option_clr.click()
                    self._type_when_element_present(by=By.CSS_SELECTOR,
                                                    value='#row{} input[type=file]'.format(row_index),
                                                    content=file_path)
                    row_index = row_index + 1

        # click submit button
        self._type_when_element_present(by=By.CSS_SELECTOR, value='input.buttongap1.update',
                                        content=Keys.CONTROL + Keys.ENTER)
        logger.info('Confirmed upload')

        # switched to tab 3
        time.sleep(1)
        self._driver.switch_to.window(self._driver.window_handles[-1])
        logger.info('switched to 3rd tab')
        self._click_when_element_present(by=By.CSS_SELECTOR, value='input.button.upload')

        number_of_tabs = len(self._driver.window_handles)
        while number_of_tabs == 3:
            time.sleep(2)
            number_of_tabs = len(self._driver.window_handles)

        # re-switch tab2
        self._driver.switch_to.window(self._driver.window_handles[-1])
        logger.info('re-switched to 2nd tab')

        self._click_when_element_present(by=By.CSS_SELECTOR, value='input.button.upload')
        while number_of_tabs == 2:
            time.sleep(2)
            number_of_tabs = len(self._driver.window_handles)

        # re-switch tab1 and re-get iframe
        time.sleep(1)

        self._driver.switch_to.window(self._driver.window_handles[-1])
        iframe = self._driver.find_element(by=By.ID, value='applicationIframe')
        self._driver.switch_to.frame(iframe)
        logger.info('re-switched to 1st tab')

        self._click_when_element_present(by=By.PARTIAL_LINK_TEXT, value='SEARCH')

        logger.info('Back to Homepage')
