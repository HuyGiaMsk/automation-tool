import os
import threading
import time
from datetime import datetime, timedelta
from logging import Logger
from typing import Callable

from selenium.webdriver.common.by import By

from src.common.FileUtil import get_excel_data_in_column_start_at_row, extract_zip, \
    check_parent_folder_contain_all_required_sub_folders, remove_all_in_folder
from src.common.StringUtil import join_set_of_elements
from src.common.ThreadLocalLogger import get_current_logger
from src.task.WebTask import WebTask


class Download_Blx(WebTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['username', 'password', 'download.folder', 'excel.path', 'excel.sheet',
                                     'excel.column.bill']
        return mandatory_keys

    def automate(self) -> None:

        logger: Logger = get_current_logger()
        logger.info(
            "---------------------------------------------------------------------------------------------------------")
        logger.info("Start processing")

        self._driver.get('https://apll.get-traction.com/')

        logger.info('Try to login')
        self.__login()
        logger.info("Login successfully")

        bills: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                 self._settings['excel.sheet'],
                                                                 self._settings['excel.column.bill'])

        if len(bills) == 0:
            logger.error('Input booking id list is empty ! Please check again')

        last_bill: str = ''
        self.perform_mainloop_on_collection(bills, self.operation_on_each_element)

        self._driver.close()
        logger.info(
            "---------------------------------------------------------------------------------------------------------")
        logger.info("End processing")
        logger.info("Summary info about list of successful and unsuccessful attempts to download each "
                    "booking's documents during the program")

        # Display summary info to the user
        self.__check_up_all_downloads(set(bills))

        # Pause and wait for the user to press Enter
        logger.info("It ends at {}. Press any key to end program...".format(datetime.now()))

    def __login(self) -> None:
        username: str = self._settings['username']
        password: str = self._settings['password']

        self._type_when_element_present(by=By.ID, value='username', content=username)
        self._type_when_element_present(by=By.ID, value='password', content=password)
        self._click_and_wait_navigate_to_other_page(by=By.CSS_SELECTOR, value='input[type=button]')

    def operation_on_each_element(self, bill):
        logger: Logger = get_current_logger()

        if self.terminated is True:
            return

        with self.pause_condition:

            while self.paused:
                self.pause_condition.wait()

            if self.terminated is True:
                return

        logger.info("Processing booking : " + bill)
        self.__navigate_and_download(bill)
        last_bill = bill

    def __check_up_all_downloads(self, booking_ids: set[str]) -> None:
        logger: Logger = get_current_logger()
        time.sleep(10 * self._timingFactor)
        is_all_contained, successful_bills, unsuccessful_bills = check_parent_folder_contain_all_required_sub_folders(
            parent_folder=self._download_folder, required_sub_folders=booking_ids)

        logger.info('{} successful booking folders containing documents has been download'
                    .format(len(successful_bills)))
        successful_bills = join_set_of_elements(successful_bills, " ")
        logger.info(successful_bills)

        if not is_all_contained:
            logger.error('{} fail attempts for downloading documents in all these bookings'
                         .format(len(unsuccessful_bills)))
            successful_bills = join_set_of_elements(unsuccessful_bills, " ")
            logger.info(successful_bills)

    def __navigate_and_download(self, bill: str) -> None:
        logger: Logger = get_current_logger()
        self._type_when_element_present(by=By.CSS_SELECTOR, value='div.fm.fm-html input[type=text]', content=bill)
        # click find button
        self._click_when_element_present(by=By.CSS_SELECTOR, value='button.gwt-Button')

        # click download
        try:
            self._click_when_element_present(by=By.LINK_TEXT, value='{}_REVISED.zip'.format(bill))
            full_file_path: str = os.path.join(self._download_folder, "{}_REVISED.zip".format(bill))
        except:
            self._click_when_element_present(by=By.LINK_TEXT, value='{}.zip'.format(bill))
            logger.info('get bill not revised')
            full_file_path: str = os.path.join(self._download_folder, bill + '.zip')

        self._wait_download_file_complete(full_file_path)
        extract_zip_task = threading.Thread(target=extract_zip,
                                            args=(full_file_path, self._download_folder,
                                                  self.delete_redundant_opening_pdf_files,
                                                  None),
                                            daemon=False)
        extract_zip_task.start()

        # click to back to the overview Booking page
        self._driver.get('https://apll.get-traction.com/')
        logger.info("Navigating back to overview Booking page")

    @staticmethod
    def delete_redundant_opening_pdf_files(download_folder: str) -> None:
        # aim to perform in root folder which is actually the defined download folder
        remove_all_in_folder(folder_path=download_folder,
                             only_files=True,
                             file_extension="pdf",
                             elapsed_time=timedelta(minutes=2))
