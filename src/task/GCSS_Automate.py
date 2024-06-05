import re
import time
from logging import Logger
from typing import Callable, Tuple

import autoit
import pyautogui
import pygetwindow as gw
from pygetwindow import Win32Window
from pywinauto import Application, WindowSpecification, ElementNotFoundError
from pywinauto.controls.common_controls import ListViewWrapper, _listview_item
from pywinauto.controls.win32_controls import ComboBoxWrapper, ButtonWrapper, EditWrapper

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
        self.time_sleep: float = float(settings.get('time.unit.factor'))
        self.app: Application = None

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

        self.app: Application = Application().connect(title=self.window_title_stack[1])
        window: WindowSpecification = self.app.window(title=self.window_title_stack[1])

        self.into_parties_tab(window)

        'Get Cnee Name in 2nd window'
        cnee_name: str = self.get_consignee(window, shipment)
        if cnee_name is None:
            self.close_windows_util_reach_first_gscc()
            return

        # 'Get Invoice Tab'
        window: WindowSpecification = self.app.window(title=self.window_title_stack[1])
        self.into_freight_and_pricing_tab(window)

        # 'Get the tab Invoice, then click the Modify button in 2nd window'
        third_window_title: str = 'Maintain Pricing and Invoicing'
        self.into_invoices_tab_and_click_btn_modify(window, next_window_title=third_window_title)

        # 'Open and interface with 3rd window - Maintain Pricing and Invoicing Window'
        pricing_n_invoice_window: WindowSpecification = self.app.window(title=third_window_title)
        Tab_controls_maintain_pricing_and_inv = pricing_n_invoice_window.children(class_name="SysTabControl32")[0]
        Tab_controls_maintain_pricing_and_inv.select(tab=1)

        self.click_all_item_payment_term_collect(logger, pricing_n_invoice_window)

        while not self.is_having_window_with_title('Maintain Invoice Details'):
            pyautogui.hotkey('alt', 'i')
        self.window_title_stack.append('Maintain Invoice Details')

        'Open 4th Window - Maintain Invoice Details'
        window_title_maintain: str = 'Maintain Invoice Details'
        window_maintain: WindowSpecification = self.app.window(title=window_title_maintain)

        ComboBox_maintain_payment: ComboBoxWrapper = window_maintain.children(class_name="ComboBox")[0]
        ComboBox_maintain_payment.select('Collect')

        'when choosing this payment term, it will automate show a dialog Information'
        dialog_title_infor = 'Information'
        dialog_window = self.app.window(title=dialog_title_infor)
        dialog_window.wait(timeout=10, wait_for='visible')
        dialog_window.type_keys('{ENTER}')

        ComboBox_maintain_invoice_party: ComboBoxWrapper = window_maintain.children(class_name="ComboBox")[1]
        item_invoices = ComboBox_maintain_invoice_party.item_texts()

        runner: int = 0
        for invoice in item_invoices:

            if invoice.__contains__(cnee_name):
                ComboBox_maintain_invoice_party.select(runner)
                dialog_title_qs = 'Question'
                dialog_window = self.app.window(title=dialog_title_qs)
                dialog_window.wait(timeout=10, wait_for='visible')
                dialog_window.type_keys('{ENTER}')
                break

            runner += 1

        if runner >= item_invoices.__len__():
            # raise ValueError('No item containing {} found in ComboBox'.format(cnee_name))
            message_change_collect = 'Not found {} in Maintain Invoice Details '.format(cnee_name)
            logger.info(message_change_collect)

            self.input_status_into_excel(message_change_collect)
            self.close_windows_util_reach_first_gscc()

            return

        ComboBox_maintain_collect_business: ComboBoxWrapper = window_maintain.children(class_name="ComboBox")[3]
        ComboBox_maintain_collect_business.select('Maersk Bangkok (Bangkok)')

        ComboBox_maintain_printable_freight_line: ComboBoxWrapper = window_maintain.children(class_name="ComboBox")[4]
        ComboBox_maintain_printable_freight_line.select('Yes')

        # Click OK button in 4th window and window will be auto closed
        while not self.is_having_window_with_title(self.window_title_stack[self.window_title_stack.__len__() - 1]):
            pyautogui.hotkey('alt', 'k')
        self.window_title_stack.append('Maintain Pricing and Invoicing')

        # Click complete collect button _ in Maintain Pricing and Invoicing window
        collect_details_collect_status: EditWrapper = window.children(class_name="Edit")[3]
        while True:
            pyautogui.hotkey('alt', 't')
            logger.info()
            if collect_details_collect_status == 'Yes':
                break

            time.sleep(self.time_sleep)

        # Input Excel
        status_shipment: str = 'Done'
        self.input_status_into_excel(status_shipment)

        self.close_windows_util_reach_first_gscc()

    def click_all_item_payment_term_collect(self, logger, pricing_n_invoice_window) -> int:
        """"
            return the number of item payment term collect have been clicked
        """

        list_views: ListViewWrapper = pricing_n_invoice_window.children(class_name="SysListView32")[1]
        count_item: int = 0

        pyautogui.keyDown('ctrl')
        for item in list_views.items():
            item: _listview_item
            if str(item.text()).__contains__('Collect'):
                item.select()
                count_item += 1

        pyautogui.keyUp('ctrl')
        if count_item > 0:
            return count_item

        status_collect = 'No term Collect'
        logger.info(status_collect)
        pyautogui.keyUp('ctrl')
        self.input_status_into_excel(status_collect)
        path_to_excel = self._settings['excel.path']
        self.excel_provider.save(workbook=self.excel_provider.get_workbook(path=path_to_excel))
        self.close_windows_util_reach_first_gscc()
        return 0

    def into_freight_and_pricing_tab(self, window: WindowSpecification):

        while True:
            pyautogui.hotkey('ctrl', 'g')

            list_views: list[ListViewWrapper] = window.children(class_name="SysListView32")
            if list_views.__len__() == 6:
                break

            time.sleep(0.5)

    def into_parties_tab(self, window: WindowSpecification):

        while True:
            pyautogui.hotkey('ctrl', 'r')
            list_views: list[ListViewWrapper] = window.children(class_name="SysListView32")
            if list_views.__len__() == 2:
                break
            time.sleep(0.5)

    def wait_for_window(self, title):
        max_attempt: int = 30
        current_attempt: int = 0

        while current_attempt < max_attempt:

            window_titles: list[str] = gw.getAllTitles()

            for window_title in window_titles:
                if window_title.__contains__(title):
                    autoit.win_activate(window_title)
                    return window_title

            current_attempt += 1
            time.sleep(self.time_sleep)

        raise Exception('Can not find out the asked window {}'.format(title))

    def get_consignee(self, window: WindowSpecification, shipment) -> str | None:

        logger: Logger = get_current_logger()
        logger.info('Checking parties')
        list_views: ListViewWrapper = window.children(class_name="SysListView32")[0]

        items = list_views.items()
        count_row = 0

        cnee_name = None
        inv_party = None
        cre_party = None

        for item in items:
            item: _listview_item

            if str(item.text()).__contains__('Consignee'):
                cnee_name = (items[count_row - 1].text())
                logger.debug('Get CNEE {}'.format(cnee_name))
                item.select()
                count_row += 1
                continue

            if str(item.text()).__contains__('Invoice Party'):
                inv_party = (items[count_row - 1].text())
                logger.debug('Get INV Party {}'.format(inv_party))
                count_row += 1
                continue

            if str(item.text()).__contains__('Credit Party'):
                cre_party = (items[count_row - 1].text())
                logger.debug('Get Credit Party {}'.format(cre_party))
                count_row += 1
                continue

            count_row += 1

        if cnee_name is None:
            message_cnee = 'Cannot find CNEE name of the shipment {} in 1st window - Ctrl R'.format(shipment)
            logger.info(message_cnee)
            self.input_status_into_excel(message_cnee)
            return None

        if cnee_name == cre_party and cnee_name == inv_party:
            logger.info('All parties are correct, modifying payment term')
            return cnee_name

        cnee_edit_element: EditWrapper = window.children(class_name="Edit")[14]
        cnee_scv_no: str = cnee_edit_element.texts()

        logger.info('Invoice and Credit parties not correct, re-updating these parties')

        self.adding_invoice_and_credit_parties(window, cnee_scv_no)

        logger.info('Updated Invoice and Credit Parties - SCV {}'.format(cnee_scv_no))

        return cnee_name

    def adding_invoice_and_credit_parties(self, window: WindowSpecification, cnee_scv_no: str):
        logger: Logger = get_current_logger()

        while not self.is_having_window_with_title('Party Details'):
            pyautogui.hotkey('alt', 'a')
        self.window_title_stack.append('Party Details')

        while not self.is_having_window_with_title('Customer Search'):
            pyautogui.hotkey('alt', 'c')
        self.window_title_stack.append('Customer Search')

        window = self.app.window(title=self.window_title_stack[self.window_title_stack.__len__() - 1])

        pyautogui.typewrite(cnee_scv_no[0])
        ComboBox_Customer_ID: ComboBoxWrapper = window.children(class_name="ComboBox")[0]
        ComboBox_Customer_ID.select('Customer ID')

        self.window_title_stack.pop()
        while not self.is_current_window_having_title(self.window_title_stack[self.window_title_stack.__len__() - 1]):
            pyautogui.hotkey('alt', 's')
            time.sleep(self.time_sleep)

        window = self.app.window(title=self.window_title_stack[self.window_title_stack.__len__() - 1])
        list_views: ListViewWrapper = window.children(class_name="SysListView32")[0]

        try:
            pyautogui.keyDown('ctrl')
            for item in list_views.items():
                if item.text() == 'Invoice Party':
                    item.select()
                    continue
                if item.text() == 'Credit Party':
                    item.select()
                    continue
        finally:
            pyautogui.keyUp('ctrl')

        list_btn: list[ButtonWrapper] = window.children(class_name="Button")
        button_right: ButtonWrapper = list_btn[3]
        button_right.click()

        # Click OK button _ in Maintain Invoice Details
        self.window_title_stack.pop()
        while not self.is_current_window_having_title(self.window_title_stack[self.window_title_stack.__len__() - 1]):
            pyautogui.hotkey('alt', 'k')
            time.sleep(self.time_sleep)
        window = self.app.window(title=self.window_title_stack[self.window_title_stack.__len__() - 1])

    def into_invoices_tab_and_click_btn_modify(self, window: WindowSpecification, next_window_title: str):
        Tab_controls = window.children(class_name="SysTabControl32")[0]
        Tab_controls.select(tab=1)

        while self.is_having_window_with_title(next_window_title) is False:
            Modify_button_in_invoices: list[ButtonWrapper] = window.children(class_name="Button")

            for button in Modify_button_in_invoices:
                if button.texts()[0] == 'Modif&y':
                    button.click()
                    return
            # pyautogui.hotkey('alt', 'y')

    def is_having_window_with_title(self, title: str) -> bool:
        window_titles: list[str] = gw.getAllTitles()

        for window_title in window_titles:
            if window_title.__contains__(title):
                return True

        return False

    def is_current_window_having_title(self, expected_title: str) -> bool:
        window: Win32Window = gw.getActiveWindow()

        if window.title == expected_title:
            return True

        return False

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
