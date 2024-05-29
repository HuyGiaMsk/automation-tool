import re
import time
from logging import Logger
from typing import Callable, Tuple

import autoit
import pyautogui
import pygetwindow as gw
from pywinauto import Application, WindowSpecification, ElementNotFoundError
from pywinauto.controls.common_controls import ListViewWrapper, _listview_item
from pywinauto.controls.win32_controls import ComboBoxWrapper, ButtonWrapper

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.AutomatedTask import AutomatedTask


class GCSS_Automate(AutomatedTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self.excel_provider: ExcelReaderProvider = None
        self.window_title_stack: list[str] = None
        self.destination_worksheet = None
        self.current_status_excel_col_index: int = 0
        self.current_status_excel_row_index: int = 0

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['excel.path', 'excel.sheet', 'excel.shipment', 'excel.status.cell']
        return mandatory_keys

    def automate(self):
        logger: Logger = get_current_logger()
        self.excel_provider: ExcelReaderProvider = XlwingProvider()
        path_to_excel = self._settings['excel.path']
        workbook = self.excel_provider.get_workbook(path=path_to_excel)
        logger.info('Loading excel files')

        sheet_name: str = self._settings['excel.sheet']
        self.destination_worksheet = self.excel_provider.get_worksheet(workbook, sheet_name)

        shipments: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                     self._settings['excel.sheet'],
                                                                     self._settings['excel.shipment'])

        col, row = self.extract_row_col_from_cell_pos_format(self._settings['excel.status.cell'])
        self.current_status_excel_col_index: int = int(self.get_letter_position(col))
        self.current_status_excel_row_index: int = int(row)

        self.wait_for_window('Pending Tray')
        self.current_element_count = 0
        self.total_element_size = len(shipments)

        for shipment in shipments:
            logger.info("Start process shipment " + shipment)
            if self.terminated is True:
                return
            with self.pause_condition:
                while self.paused:
                    logger.info("Currently pause")
                    self.pause_condition.wait()
                if self.terminated is True:
                    return

            pyautogui.hotkey('ctrl', 'o')
            pyautogui.typewrite(shipment)
            pyautogui.hotkey('tab')
            pyautogui.hotkey('enter')

            self.process_on_each_shipment(shipment)

            self.current_element_count += 1
            self.current_status_excel_row_index += 1
            self.excel_provider.save(workbook)
            logger.info("Done with shipment " + shipment)

        self.excel_provider.close(workbook)

    def extract_row_col_from_cell_pos_format(self, start_cell: str) -> Tuple[str, int]:
        result = re.search(r'([a-zA-Z]+)(\d+)', start_cell)
        if result:
            column = result.group(1)
            row = int(result.group(2))
            return column, row
        else:
            raise ValueError("Input does not match Excel cell position format")

    def get_letter_position(self, letter: str):
        if letter.__len__() != 1:
            raise Exception('Please provide the single char, not a word ! the previous input was ' + letter)
        letter_to_position = {chr(i + 96): i for i in range(1, 27)}

        return letter_to_position.get(letter.lower(), "Invalid input")

    def input_status_into_excel(self, message):
        self.excel_provider.change_value_at(self.destination_worksheet, self.current_status_excel_row_index,
                                            self.current_status_excel_col_index, message)

    def process_on_each_shipment(self, shipment):

        """Loop for each shipment in this method.
        Interface with 2nd window (Ctrl R - to get CNEE name)
        3rd window (Ctrl G - Maintain Pricing),
        4th window (Maintain Invoice Details).
        Return Payment Term, Invoice party index with CNEE name and Shipment in file Excel"""

        logger: Logger = get_current_logger()

        GCSS_Shipment_MSL_Active_Title: str = self.wait_for_window(shipment)
        autoit.win_activate(GCSS_Shipment_MSL_Active_Title)

        self.window_title_stack = ['Pending Tray', GCSS_Shipment_MSL_Active_Title, 'Maintain Pricing and Invoicing',
                                   'Maintain Invoice Details']

        app: Application = Application().connect(title=self.window_title_stack[1])
        window: WindowSpecification = app.window(title=self.window_title_stack[1])

        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'r')

        'Get Cnee Name in 2nd window'
        cnee_name: str = self.get_consignee(window, shipment)
        if cnee_name is None:
            self.close_windows_util_reach_first_gscc()
            return

        pyautogui.hotkey('ctrl', 'g')
        time.sleep(0.5)

        'Get the tab Invoice, then click the Modify button in 2nd window'
        click_result: int = self.click_btn_modify(window)

        if click_result == 0:
            self.close_windows_util_reach_first_gscc()
            return

        'Open and interface with 3rd window - Maintain Pricing and Invoicing Window'

        window_title_third: str = 'Maintain Pricing and Invoicing'
        window_third: WindowSpecification = app.window(title=window_title_third)

        self.get_maintain_pricing_invoice_tab(window_third)

        pyautogui.keyDown('ctrl')

        list_views: ListViewWrapper = window_third.children(class_name="SysListView32")[1]
        count_item: int = 0
        for item in list_views.items():
            item: _listview_item
            if str(item.text()).__contains__('Collect'):
                item.select()
                count_item += 1

        if count_item == 0:
            status_collect = 'No term Collect'

            logger.info(status_collect)
            pyautogui.keyUp('ctrl')

            self.input_status_into_excel(status_collect)

            path_to_excel = self._settings['excel.path']

            self.excel_provider.save(workbook=self.excel_provider.get_workbook(path=path_to_excel))

            # self.excel_provider.close(workbook=self.settings['excel.path'])

            self.close_windows_util_reach_first_gscc()

            return

        pyautogui.keyUp('ctrl')

        try_press_i_to_open_maintain_invoice: int = 0

        try:
            if try_press_i_to_open_maintain_invoice > 10:
                raise Exception
            pyautogui.hotkey('i')
            time.sleep(0.5)
            try_press_i_to_open_maintain_invoice += 1

        except:

            message_choose_row_collect = 'Cannot Open window Maintain Invoice Details'
            logger.info(message_choose_row_collect)
            self.input_status_into_excel(message_choose_row_collect)

            self.close_windows_util_reach_first_gscc()

            return None

        'Open 4th Window - Maintain Invoice Details'
        self.wait_for_window('Maintain Invoice Details')

        window_title_maintain: str = 'Maintain Invoice Details'
        window_maintain: WindowSpecification = app.window(title=window_title_maintain)

        ComboBox_maintain_payment: ComboBoxWrapper = window_maintain.children(class_name="ComboBox")[0]
        ComboBox_maintain_payment.select('Collect')

        'when choosing this payment term, it will automate show a dialog Information'
        dialog_title_infor = 'Information'
        dialog_window = app.window(title=dialog_title_infor)
        dialog_window.wait(timeout=10, wait_for='visible')
        dialog_window.type_keys('{ENTER}')

        ComboBox_maintain_invoice_party: ComboBoxWrapper = window_maintain.children(class_name="ComboBox")[1]
        item_invoices = ComboBox_maintain_invoice_party.item_texts()

        runner: int = 0
        for invoice in item_invoices:

            if invoice.__contains__(cnee_name):
                ComboBox_maintain_invoice_party.select(runner)
                dialog_title_qs = 'Question'
                dialog_window = app.window(title=dialog_title_qs)
                dialog_window.wait(timeout=10, wait_for='visible')
                dialog_window.type_keys('{ENTER}')
                break

            runner += 1

        if runner >= item_invoices.__len__():
            # raise ValueError('No item containing {} found in ComboBox'.format(cnee_name))
            message_change_collect = 'Not found {} in Maintain Invoice Details '.format(cnee_name)
            logger.info(message_change_collect)

            self.input_status_into_excel(message_change_collect)

            try_click_cancle_button = 0
            try:
                if try_click_cancle_button > 30:
                    raise Exception
                self.wait_for_window('Maintain Invoice Details')

                Modify_button_in_4th_window: list[ButtonWrapper] = window_maintain.children(class_name="Button")

                for button in Modify_button_in_4th_window:
                    if button.texts()[0] == 'Cancel':
                        button.click()
                        break
                try_click_cancle_button += 1

            except:
                raise Exception

            self.close_windows_util_reach_first_gscc()

            return

        ComboBox_maintain_collect_business: ComboBoxWrapper = window_maintain.children(class_name="ComboBox")[3]
        ComboBox_maintain_collect_business.select('Maersk Bangkok (Bangkok)')

        ComboBox_maintain_printable_freight_line: ComboBoxWrapper = window_maintain.children(class_name="ComboBox")[4]
        ComboBox_maintain_printable_freight_line.select('Yes')

        'Dont have permission to change/edit these infor -> need to recheck'
        # pyautogui.hotkey('ctrl', 'k')
        time.sleep(0.5)

        'Assume this code can be automate close, then we will close all window until found main window of GCSS'

        status_shipment: str = 'Done'
        self.input_status_into_excel(status_shipment)

        self.close_windows_util_reach_first_gscc()

    def wait_for_window(self, title):
        max_attempt: int = 30
        current_attempt: int = 0

        while current_attempt < max_attempt:

            # Get a list of all open window titles
            window_titles: list[str] = gw.getAllTitles()

            for window_title in window_titles:
                if window_title.__contains__(title):
                    autoit.win_activate(window_title)
                    return window_title

            current_attempt += 1
            time.sleep(1)

        raise Exception('Can not find out the asked window {}'.format(title))

    def get_consignee(self, window: WindowSpecification, shipment) -> str | None:

        logger: Logger = get_current_logger()
        try_to_get_cnee: int = 0
        try:
            list_views: ListViewWrapper = window.children(class_name="SysListView32")[0]
            items = list_views.items()

            if try_to_get_cnee > 30:
                raise Exception
            count = 0

            for item in items:
                item: _listview_item

                if str(item.text()).__contains__('Consignee'):
                    cnee_name = (items[count - 1].text())
                    logger.info('Get CNEE {}'.format(cnee_name))
                    return cnee_name

                count += 1

            try_to_get_cnee += 1

            time.sleep(0.5)
        except:
            message_cnee = 'Cannot find CNEE name of the shipment {} in 1st window - Ctrl R'.format(shipment)
            logger.info(message_cnee)
            self.input_status_into_excel(message_cnee)
            return None

    def click_btn_modify(self, window: WindowSpecification) -> int:
        """
        Purpose : click the btn Modify
        Return
               0: Can not click
               1: Clicked ok
        """
        logger: Logger = get_current_logger()

        logger.info("Try to click btn modify")

        Tab_controls = window.children(class_name="SysTabControl32")[0]
        Tab_controls.select(tab=1)

        Modify_button_in_invoices: list[ButtonWrapper] = window.children(class_name="Button")

        for button in Modify_button_in_invoices:

            if button.texts()[0] == 'Modif&y':
                button.click()
                logger.info("Click btn modify ok")
                return 1

        return 0

    def get_maintain_pricing_invoice_tab(self, window: WindowSpecification):
        logger: Logger = get_current_logger()
        try_to_get_systab: int = 0
        try:

            if try_to_get_systab > 30:
                raise Exception

            Tab_controls_maintain = window.children(class_name="SysTabControl32")[0]
            Tab_controls_maintain.select(tab=1)
            time.sleep(0.5)
            try_to_get_systab += 1

        except:
            message_pricing = 'Cannot get tab Invoice in window Maintain pricing and invoicing'
            logger.info(message_pricing)

            self.input_status_into_excel(message_pricing)
            return

    def close_windows_util_reach_first_gscc(self):
        return self.close_windows_with_window_title_stack(self.window_title_stack)

    def close_windows_with_window_title_stack(self, window_title_stack: list[str]):
        logger: Logger = get_current_logger()

        for i in range(window_title_stack.__len__() - 1, 0, -1):
            window_title = window_title_stack[i]

            try:
                app = Application().connect(title=window_title, timeout=10)

                window = app.window(title=window_title)

                if window.exists(timeout=1):
                    window.close()
                    logger.info(f"Window with title '{window_title}' has been closed.")
                else:
                    logger.info(f"Window with title '{window_title}' does not exist.")

            except ElementNotFoundError:
                logger.info(f"No window with title '{window_title}' was found.")
            except Exception as e:
                logger.info(f"An error occurred: {e}")

        self.wait_for_window(window_title_stack[0])
