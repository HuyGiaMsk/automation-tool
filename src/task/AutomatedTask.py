import logging
import os
import time
import uuid
from abc import abstractmethod, ABC
from datetime import datetime
from logging import Logger
from typing import Callable

from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.remote.webdriver import WebDriver as AnyDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait

from src.common.Percentage import Percentage
from src.common.ResumableThread import ResumableThread
from src.common.StringUtil import validate_keys_of_dictionary
from src.common.ThreadLocalLogger import get_current_logger, create_thread_local_logger
from src.setup.DownloadDriver import place_suitable_chromedriver, get_full_browser_driver_path


class AutomatedTask(Percentage, ResumableThread, ABC):
    @abstractmethod
    def mandatory_settings(self) -> list[str]:
        pass

    @abstractmethod
    def automate(self):
        pass

    @property
    def settings(self) -> dict[str, str]:
        return self._settings

    @settings.setter
    def settings(self, settings: dict[str, str]):
        if settings is None:
            raise ValueError("Provided settings is None")
        self._settings = settings

    def __init__(self,
                 settings: dict[str, str],
                 callback_before_run_task: Callable[[], None]):

        super().__init__()

        logger: Logger = get_current_logger()
        self._settings: dict[str, str] = settings
        self.callback_before_run_task = callback_before_run_task

        self._download_folder = self._settings.get('download.folder')
        if self._download_folder is not None:
            if os.path.isfile(self._download_folder):
                logger.info(f"Provided download folder '{self._download_folder}'is not valid. It is a file, "
                            f"not folder")
                raise Exception(f"Provided download folder '{self._download_folder}'is not valid. It is a file, "
                                f"not folder")

            if not os.path.exists(self._download_folder):
                os.makedirs(self._download_folder)
                logger.info(f"Create folder '{self._download_folder}' because it is not existed by default")

        if self._settings.get('time.unit.factor') is None:
            self._timingFactor = 1.0
        else:
            self._timingFactor = float(self._settings.get('time.unit.factor'))

        if self._settings.get('use.GUI') is None:
            self.use_gui = False
        else:
            self.use_gui = 'True'.lower() == str(self._settings.get('use.GUI')).lower()

        if not self.use_gui:
            logger.info('Run in headless mode')
        self._driver = None

        if settings.get('time.unit.factor') is None:
            self._time_sleep: float = 1
        else:
            self._time_sleep = float(settings.get('time.unit.factor'))

    def perform(self) -> None:
        mandatory_settings: list[str] = self.mandatory_settings()
        mandatory_settings.append('invoked_class')
        validate_keys_of_dictionary(self._settings, set(mandatory_settings))

        browser_driver: str = get_full_browser_driver_path()
        self._driver = self._setup_driver(browser_driver)
        logger: Logger = create_thread_local_logger(class_name=self._settings['invoked_class'],
                                                    thread_uuid=str(uuid.uuid4()))

        if self.callback_before_run_task is not None:
            self.callback_before_run_task()

        try:
            self.automate()
        except Exception as exception:
            logger.exception(str(exception))
        logger.info("Done task. It ends at {}".format(datetime.now()))
        del logging.Logger.manager.loggerDict[self._settings['invoked_class']]

    def perform_mainloop_on_collection(self,
                                       collection,
                                       critical_operation_on_each_element: Callable[[object], None]):
        self.current_element_count = 0
        self.total_element_size = len(collection)
        logger: Logger = get_current_logger()

        for each_element in collection:

            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    logger.info("Currently pause")
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

            critical_operation_on_each_element(each_element)
            self.current_element_count = self.current_element_count + 1

    def _setup_driver(self, browser_driver: str) -> WebDriver:

        options: webdriver.ChromeOptions = webdriver.ChromeOptions()

        if not self.use_gui:
            options.add_argument("--headless")
            options.add_argument('--disable-gpu')
            options.add_argument("--window-size=%s" % "1920,1080")
            options.add_argument("--use-fake-ui-for-media-stream")

        else:
            options.add_argument("--start-maximized")

        options.add_argument('--disable-extensions')
        options.add_argument('--disable-infobars')
        options.add_argument('--disable-notifications')

        # prefs = {
        #     "download.prompt_for_download": False,
        #     "download.directory_upgrade": True,
        #     "safebrowsing.enabled": False
        # }
        #
        # options.add_experimental_option("prefs", prefs)
        # # username = os.getlogin()
        # #
        # # options.add_argument(r'--user-data-dir=C:\Users\{}\AppData\Local\Google\Chrome\User Data'.format(username))
        # options.add_argument('--disable-extensions')
        # options.add_argument('--disable-infobars')
        # options.add_argument('--disable-notifications')
        # options.add_argument("--disable-blink-features=AutomationControlled")
        # options.add_experimental_option("excludeSwitches", ["enable-automation"])
        # options.add_experimental_option("useAutomationExtension", False)

        download_path: str = self._download_folder
        prefs: dict = {
            "profile.default_content_settings.popups": 0,
            "download.default_directory": r'{}'.format(download_path),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "excludeSwitches": ['enable-logging'],
            "safebrowsing.enable": False
        }
        if not self.use_gui:
            prefs['plugins.always_open_pdf_externally'] = True

        options.add_experimental_option("prefs", prefs)

        if not os.path.exists(browser_driver):
            place_suitable_chromedriver()
        service: webdriver.ChromeService = webdriver.ChromeService(executable_path=r'{}'.format(browser_driver))

        driver: webdriver.Chrome = webdriver.Chrome(service=service, options=options)
        return driver

    def sleep(self) -> None:
        self._time_sleep = float(self.settings.get('time.unit.factor'))
        time.sleep(self._time_sleep)
        return

    def _wait_download_file_complete(self, file_path: str) -> None:
        logger: Logger = get_current_logger()
        logger.info(r'Waiting for downloading {} complete'.format(file_path))
        attempt_counting: int = 0
        max_attempt: int = 60 * 3
        while True:
            if attempt_counting > max_attempt:
                raise Exception('The webapp waiting too long to download {}. Please check'.format(file_path))

            if not os.path.exists(file_path):
                time.sleep(1 * self._timingFactor)
                attempt_counting += 1
                continue

            break
        logger.info(r'Downloading {} complete'.format(file_path))

    def _wait_navigating_to_other_page_complete(self, previous_url: str, expected_end_with: str = None) -> None:
        logger: Logger = get_current_logger()
        attempt_counting: int = 0
        max_attempt: int = 1200
        while True:
            current_url: str = self._driver.current_url

            if attempt_counting > max_attempt:
                raise Exception('The webapp is not navigating as expected, previous url is{}'.format(previous_url))

            if current_url == previous_url:
                logger.info('Still waiting for {}\'s changing'.format(previous_url))
                time.sleep(1 * self._timingFactor)
                attempt_counting += 1
                continue

            if expected_end_with is not None and not current_url.endswith(expected_end_with):
                logger.warning('It has been navigated to {}'.format(current_url))
                time.sleep(1 * self._timingFactor)
                attempt_counting += 1
                continue

            break

    def _wait_to_close_all_new_tabs_except_the_current(self):
        current_attempt: int = 0
        max_attempt: int = 60 * 3
        while True:
            number_of_current_tabs: int = len(self._driver.window_handles)
            if current_attempt > max_attempt:
                logging.error('Can not load the file')
                raise Exception('Can not load the file')

            if number_of_current_tabs > 1:
                time.sleep(1 * self._timingFactor)
                runner: int = number_of_current_tabs
                while runner > 1:
                    time.sleep(1 * self._timingFactor)
                    self._driver.switch_to.window(self._driver.window_handles[runner - 1])
                    self._driver.close()
                    runner = runner - 1

                self._driver.switch_to.window(self._driver.window_handles[0])
                break
            else:
                time.sleep(1 * self._timingFactor)
                current_attempt = current_attempt + 1

    def _type_when_element_present(self, by: str, value: str, content: str, time_sleep: int = 1) -> WebElement:
        web_element: WebElement = self.__get_element_satisfy_predicate(by,
                                                                       value,
                                                                       expected_conditions.presence_of_element_located(
                                                                           (by, value)),
                                                                       time_sleep)

        web_element.send_keys(content)
        return web_element

    def _click_when_element_present(self, by: str, value: str, time_sleep: int = 1) -> WebElement:
        web_element: WebElement = self.__get_element_satisfy_predicate(by,
                                                                       value,
                                                                       expected_conditions.presence_of_element_located(
                                                                           (by, value)),
                                                                       time_sleep)

        web_element.click()
        return web_element

    def _click_and_wait_navigate_to_other_page(self, by: str, value: str, time_sleep: int = 1) -> WebElement:
        previous_url: str = self._driver.current_url
        web_element: WebElement = self.__get_element_satisfy_predicate(by,
                                                                       value,
                                                                       expected_conditions.presence_of_element_located(
                                                                           (by, value)),
                                                                       time_sleep)
        web_element.click()
        self._wait_navigating_to_other_page_complete(previous_url=previous_url)
        return web_element

    def _get_when_element_present(self, by: str, value: str, time_sleep: int = 1) -> WebElement:
        web_element: WebElement = self.__get_element_satisfy_predicate(by,
                                                                       value,
                                                                       expected_conditions.presence_of_element_located(
                                                                           (by, value)),
                                                                       time_sleep)
        return web_element

    def _try_to_get_if_element_present(self, by: str, value: str, time_sleep: int = 1, waiting_time: int = 30
                                       ) -> WebElement:

        try:
            web_element: WebElement = self.__get_element_satisfy_predicate(by,
                                                                           value,
                                                                           expected_conditions.presence_of_element_located(
                                                                               (by, value)),
                                                                           time_sleep,
                                                                           waiting_time)

            return web_element
        except TimeoutException:
            return None

    def __get_element_satisfy_predicate(self,
                                        by: str,
                                        element_selector: str,
                                        method: Callable[[AnyDriver], WebElement],
                                        first_time_sleep: int = 1,
                                        waiting_time: int = 30) -> WebElement:
        time.sleep(first_time_sleep * self._timingFactor)
        if self.use_gui:
            WebDriverWait(self._driver, waiting_time * self._timingFactor).until(method)
        queried_element: WebElement = self._driver.find_element(by=by, value=element_selector)
        return queried_element

    def find_matched_option(self: object, by: str, list_options_selector: str, search_keyword: str) -> WebElement:
        options: list[WebElement] = self._driver.find_elements(by=by, value=list_options_selector)
        finding_option = None
        for current_option in options:
            current_inner_text = current_option.get_attribute('innerText')
            if current_inner_text == search_keyword:
                finding_option = current_option
                break
        if finding_option is None:
            raise Exception('Can not find out the option whose inner text match your search keyword')
        return finding_option

    def find_matched_option_shadow(self: object, by: str, list_options_selector: str,
                                   search_keyword: str) -> WebElement:
        options: list[WebElement] = self._driver.find_elements(by=by, value=list_options_selector)
        finding_option = None
        for current_option in options:
            current_inner_text = current_option.get_attribute('innerText')
            if current_inner_text == search_keyword:
                finding_option = current_option
                break
        if finding_option is None:
            raise Exception('Can not find out the option whose inner text match your search keyword')
        return finding_option
