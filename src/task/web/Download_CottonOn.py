import os
import threading
import time
from datetime import datetime, timedelta
from enum import Enum
from logging import Logger
from typing import Callable

from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement

from src.common.FileUtil import get_excel_data_in_column_start_at_row, extract_zip, \
    remove_all_in_folder
from src.common.ResourceLock import ResourceLock
from src.common.ThreadLocalLogger import get_current_logger
from src.task.WebTask import WebTask


# Define an enumeration class
class BookingToInfoIndex(Enum):
    SO_INDEX_IN_TUPLE = 0
    BECODE_INDEX_IN_TUPLE = 1


class Download_CottonOn(WebTask):
    booking_to_info = {}

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['username', 'password', 'excel.path', 'excel.sheet',
                                     'excel.column.booking', 'download.folder',
                                     'excel.column.so', 'excel.column.becode']
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
        time.sleep(1)
        self._click_when_element_present(by=By.CSS_SELECTOR, value='div[data-cy=nav-Operations]')
        # click navigating overview bookings page - on the header
        self._click_and_wait_navigate_to_other_page(by=By.CSS_SELECTOR, value='li[data-cy=bookings]')

        booking_ids: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                       self._settings['excel.sheet'],
                                                                       self._settings[
                                                                           'excel.column.booking'])

        becodes: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                   self._settings['excel.sheet'],
                                                                   self._settings[
                                                                       'excel.column.becode'])

        so_numbers: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                      self._settings['excel.sheet'],
                                                                      self._settings['excel.column.so'])
        if len(booking_ids) == 0:
            logger.error('Input booking id list is empty ! Please check again')

        if len(booking_ids) != len(becodes) or len(booking_ids) != len(becodes):
            raise Exception("Please check your input data length of becode, sonumber and booking are not equal")

        # info means becode and so number
        index: int = 0
        for booking in booking_ids:
            self.booking_to_info[booking] = (so_numbers[index], becodes[index])
            index += 1

        last_booking: str = ''
        for booking in booking_ids:

            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

            logger.info("Processing booking : " + booking)
            self.__navigate_and_download(booking)
            last_booking = booking

        self._driver.close()
        logger.info(
            "---------------------------------------------------------------------------------------------------------")
        logger.info("End processing")
        logger.info("Summary info about list of successful and unsuccessful attempts to download each "
                    "booking's documents during the program")

        # Display summary info to the user
        # self.__check_up_all_downloads(set(booking_ids))

        # Pause and wait for the user to press Enter
        logger.info("It ends at {}. Press any key to end program...".format(datetime.now()))

    def __login(self) -> None:
        username: str = self._settings['username']
        password: str = self._settings['password']

        self._type_when_element_present(by=By.ID, value='user-mail', content=username)
        self._type_when_element_present(by=By.ID, value='pwd', content=password)
        time.sleep(2)
        self._click_and_wait_navigate_to_other_page(by=By.CSS_SELECTOR, value='button[type=submit]')

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
                                                                   'div[data-cy=action-details] svg ')
        # click tab document
        self._click_when_element_present(by=By.ID, value='item-documents')

        # wait until the progress bar on view file disappear
        time.sleep(1 * self._timingFactor)

        # click downLoad all files
        self._click_when_element_present(by=By.CSS_SELECTOR, value='div[data-cy=shipment-documents-box] '
                                                                   'div:nth-child(2) button')

        full_file_path: str = os.path.join(self._download_folder, booking + '.zip')
        self._wait_download_file_complete(full_file_path)
        extract_zip_task = threading.Thread(target=extract_zip,
                                            args=(full_file_path, self._download_folder,
                                                  self.delete_redundant_opening_pdf_files,
                                                  self.rename_all_files_in_folder_extracted),
                                            daemon=False)

        extract_zip_task.start()
        # click to back to the overview Booking page
        self._click_when_element_present(by=By.CSS_SELECTOR, value='button[data-cy=iconButtonClose] '
                                                                   'span.MuiIconButton-label svg')
        self._click_when_element_present(by=By.CSS_SELECTOR, value='div[role=button] svg')
        logger.info("Navigating back to overview Booking page")

    @staticmethod
    def delete_redundant_opening_pdf_files(download_folder: str) -> None:
        # aim to perform in root folder which is actually the defined download folder
        remove_all_in_folder(folder_path=download_folder,
                             only_files=True,
                             file_extension="pdf",
                             elapsed_time=timedelta(minutes=2))

    @staticmethod
    def rename_all_files_in_folder_extracted(extracted_dir: str) -> None:
        file_name_inv: str = "Invoice"
        file_name_pkl: str = "Packing"

        booking_id: str = os.path.split(extracted_dir)[1]
        so_number = Download_CottonOn.booking_to_info[booking_id][BookingToInfoIndex.SO_INDEX_IN_TUPLE.value]
        with ResourceLock(file_path=extracted_dir):

            for file_name in os.listdir(extracted_dir):
                file_path = os.path.join(extracted_dir, file_name)

                if os.path.isdir(file_path):
                    continue

                if file_name_inv in file_name:
                    source = file_path
                    des = os.path.join(extracted_dir, "{}_INV.pdf".format(so_number))
                    os.rename(source, des)
                    continue

                if file_name_pkl in file_name:
                    source = file_path
                    des = os.path.join(extracted_dir, "{}_PKL.pdf".format(so_number))
                    os.rename(source, des)
                    continue

                os.remove(file_path)

        # rename the folder
        root_dir = os.path.dirname(extracted_dir)
        source = extracted_dir
        des = os.path.join(root_dir, so_number)
        os.rename(source, des)
