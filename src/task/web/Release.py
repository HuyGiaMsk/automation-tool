import time
from logging import Logger
from typing import Callable

from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ThreadLocalLogger import get_current_logger
from src.task.WebTask import WebTask


class Release(WebTask):
    def getCurrentPercent(self):
        pass

    def get_current_percent(self) -> float:
        pass

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self._document_folder = self._download_folder

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['username', 'password', 'excel.path', 'excel.sheet', 'download.folder',
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
                if try_attempt_count > 20:
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

        if not len(becodes) == len(so_numbers):
            raise Exception('be_codes and so_numbers do not have thhe same length')

        # Processing to upload
        index = 0
        for becode in becodes:
            becode_to_sonumber[becode] = so_numbers[index]
            index += 1

        for becode, so_number in becode_to_sonumber.items():

            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

            self._release(becode, so_number)

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

    def _release(self, becode: str, so_number: str):
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

        self._click_when_element_present(by=By.CSS_SELECTOR, value='#template #row0 td:nth-child(2) input',
                                         time_sleep=2)
        logger.info('clicked CBL box')

        self._click_when_element_present(by=By.CSS_SELECTOR, value='#moreOptionSo button')
        element = self._driver.find_element(by=By.ID, value="PLMenuId")
        driver_release = self._driver

        actions = ActionChains(driver_release)
        actions.key_down(Keys.CONTROL).click(element).key_up(Keys.CONTROL).perform()
        logger.info('Go to next window')

        number_of_tabs = len(self._driver.window_handles)
        try_count = 0
        while number_of_tabs < 2:
            if try_count > 10:
                raise Exception('Can not invoke a new tab')
            time.sleep(1)
        self._driver.switch_to.window(self._driver.window_handles[-1])
        time.sleep(2)
        logger.info('switched to tab2')

        # switched tab 2
        self._click_when_element_present(by=By.NAME, value='search2')
        self._click_when_element_present(by=By.NAME, value='ok')
        logger.info('releasing in tab2 - going to switch to tab1')
        time.sleep(1)

        self._driver.switch_to.window(self._driver.window_handles[-1])
        iframe = self._driver.find_element(by=By.ID, value='applicationIframe')
        self._driver.switch_to.frame(iframe)
        logger.info('re-switched to 1st tab')

        self._click_when_element_present(by=By.PARTIAL_LINK_TEXT, value='SEARCH')
        logger.info('Back to Homepage')
