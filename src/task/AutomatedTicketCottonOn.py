import os
import time
import threading

from datetime import datetime, timedelta
from logging import Logger
from typing import Callable

from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

from src.task.AutomatedTask import AutomatedTask
from src.common.FileUtil import get_excel_data_in_column_start_at_row, extract_zip, \
    check_parent_folder_contain_all_required_sub_folders, remove_all_in_folder
from src.common.StringUtil import join_set_of_elements
from src.common.ThreadLocalLogger import get_current_logger
from src.common.Constants import ZIP_EXTENSION


class AutomatedTicketCottonOn(AutomatedTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self.last_booking = None

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['username', 'password', 'excel.path', 'excel.sheet',
                                    'excel.column.booking', 'download.folder', 'excel.column.so', 'excel.column.becode']
        return mandatory_keys

    def automate(self) -> None:
        logger: Logger = get_current_logger()
        logger.info(
            "---------------------------------------------------------------------------------------------------------")
        logger.info("Start processing")

        self._driver.get('https://app.shipeezi.com/')

        logger.info('Try to login')
        self.__login()
        logger.info("Login successfully")

        logger.info("Navigate to overview Booking page the first time")
        # click navigating operations on header
        self._click_when_element_present(by=By.CSS_SELECTOR, value='div[data-cy=nav-Operations]')
        # click navigating overview bookings page - on the header
        self._click_and_wait_navigate_to_other_page(by=By.CSS_SELECTOR, value='li[data-cy=bookings]')

        booking_ids: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                       self._settings['excel.sheet'],
                                                                       self._settings['excel.column.booking'])
        if len(booking_ids) == 0:
            logger.error('Input booking id list is empty ! Please check again')

        self.last_booking: str = ''
        self.current_element_count = 0
        self.total_element_size = len(booking_ids)
        self.perform_mainloop_on_collection(booking_ids, self.operation_on_each_element)
        self._driver.close()

        logger.info(
            "---------------------------------------------------------------------------------------------------------")
        logger.info("End processing")
        logger.info("Summary info about list of successful and unsuccessful attempts to download each "
                    "booking's documents during the program")

        # Display summary info to the user
        self.__check_up_all_downloads(set(booking_ids), self.last_booking)

    def operation_on_each_element(self, booking):
        logger: Logger = get_current_logger()
        logger.info("Processing booking : " + booking)
        self.__navigate_and_download(booking)
        self.last_booking = booking

    def __login(self) -> None:
        username: str = self._settings['username']
        password: str = self._settings['password']

        self._type_when_element_present(by=By.ID, value='user-mail', content=username)
        self._type_when_element_present(by=By.ID, value='pwd', content=password)
        self._click_and_wait_navigate_to_other_page(by=By.CSS_SELECTOR, value='button[type=submit]')

    def __check_up_all_downloads(self, booking_ids: set[str], last_booking: str) -> None:
        logger: Logger = get_current_logger()
        last_booking_downloaded_folder: str = os.path.join(self._download_folder, last_booking)
        self._wait_download_file_complete(last_booking_downloaded_folder)

        is_all_contained, successful_bookings, unsuccessful_bookings = check_parent_folder_contain_all_required_sub_folders(
            parent_folder=self._download_folder, required_sub_folders=booking_ids)

        logger.info('{} successful booking folders containing documents has been download'
                    .format(len(successful_bookings)))
        successful_bookings = join_set_of_elements(successful_bookings, " ")
        logger.info(successful_bookings)

        if not is_all_contained:
            logger.error('{} fail attempts for downloading documents in all these bookings'
                         .format(len(unsuccessful_bookings)))
            successful_bookings = join_set_of_elements(unsuccessful_bookings, " ")
            logger.info(successful_bookings)

    def __navigate_and_download(self, booking: str) -> None:
        logger: Logger = get_current_logger()
        search_box: WebElement = self._type_when_element_present(by=By.CSS_SELECTOR,
                                                                 value='div[data-cy=search] input',
                                                                 content=booking)

        # try to click option_booking - which usually out of focus and be removed from the DOM / cause exception
        try_attempt_count: int = 0
        while True:
            try:
                if try_attempt_count > 20:
                    raise Exception('Look like we have problems with the web structure changed - '
                                    'we could not click on the option_booking ! Need to our investigation')

                time.sleep(1 * self._timingFactor)
                self._driver.find_element(by=By.CSS_SELECTOR, value='.MuiAutocomplete-option:nth-child(1)').click()
                logger.info('Clicked option_booking for {} successfully'.format(booking))
                break
            except Exception as exception:
                logger.error(str(exception))

                time.sleep(1 * self._timingFactor)
                self._driver.execute_script("arguments[0].value = '{}';".format(booking), search_box)

                time.sleep(1 * self._timingFactor)
                search_box.click()

                logger.info('The {}th sent new key and click to try revoke the autocomplete board show up '
                            'option_booking for {}'.format(try_attempt_count, booking))
                try_attempt_count += 1
                continue

        # click detail booking
        self._click_when_element_present(by=By.CSS_SELECTOR, value='td[data-cy=table-cell-actions] '
                                                                   'div[data-cy=action-details]'
                                                                   ':nth-child(2)')
        # click tab document
        self._click_when_element_present(by=By.CSS_SELECTOR, value='button[data-cy=documents]')

        # click view file
        self._click_when_element_present(by=By.CSS_SELECTOR, value='div[data-cy=shipment-documents-box] '
                                                                   '.MuiGrid-container '
                                                                   '.MuiGrid-item:nth-child(6) button')

        # wait until the progress bar on view file disappear
        time.sleep(1 * self._timingFactor)
        WebDriverWait(self._driver, 120 * self._timingFactor).until(ec.invisibility_of_element(
            (By.CSS_SELECTOR, 'div[data-cy=shipment-documents-box] .MuiGrid-container '
                              '.MuiGrid-item:nth-child(6) button .progressbar')))

        self._wait_to_close_all_new_tabs_except_the_current()

        # click downLoad all files
        self._click_when_element_present(by=By.CSS_SELECTOR, value='div[data-cy=shipment-documents-box] '
                                                                   'div:nth-child(2) button')

        full_file_path: str = os.path.join(self._download_folder, booking + ZIP_EXTENSION)
        self._wait_download_file_complete(full_file_path)
        extract_zip_task = threading.Thread(target=extract_zip,
                                            args=(full_file_path, self._download_folder,
                                                  self.delete_redundant_opening_pdf_files, None),
                                            daemon=False)
        extract_zip_task.start()

        # click to back to the overview Booking page
        self._click_and_wait_navigate_to_other_page(by=By.CSS_SELECTOR,
                                                    value='.MuiBreadcrumbs-ol .MuiBreadcrumbs-li:nth-child(1)')
        logger.info("Navigating back to overview Booking page")

    @staticmethod
    def delete_redundant_opening_pdf_files(download_folder: str) -> None:

        remove_all_in_folder(folder_path=download_folder,
                             only_files=True,
                             file_extension="pdf",
                             elapsed_time=timedelta(minutes=2))
