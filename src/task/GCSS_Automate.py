from logging import Logger
from typing import Callable, Iterator

import autoit
import pyautogui
from pywinauto import Application, WindowSpecification
from pywinauto.controls.common_controls import ListViewWrapper, _listview_item
from pywinauto.controls.win32_controls import ComboBoxWrapper, ButtonWrapper, EditWrapper

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.StringUtil import extract_row_col_from_cell_pos_format
from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.DesktopAppTask import DesktopAppTask


class GCSS_Automate(DesktopAppTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self.excel_provider: ExcelReaderProvider = None
        self.current_worksheet = None
        self.current_status_excel_col_index: int = 0
        self.current_status_excel_row_index: int = 0

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['excel.path', 'excel.sheet', 'excel.shipment', 'excel.status.address',
                                     'excel.status.cell']
        return mandatory_keys

    def automate(self):
        logger: Logger = get_current_logger()
        self.excel_provider: ExcelReaderProvider = XlwingProvider()
        path_to_excel = self._settings['excel.path']
        workbook = self.excel_provider.get_workbook(path=path_to_excel)
        logger.info('Loading excel files')

        sheet_name: str = self._settings['excel.sheet']
        self.current_worksheet = self.excel_provider.get_worksheet(workbook, sheet_name)

        shipments: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                     self._settings['excel.sheet'],
                                                                     self._settings['excel.shipment'])
        status_address: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                          self._settings['excel.sheet'],
                                                                          self._settings['excel.status.address'])

        col, row = extract_row_col_from_cell_pos_format(self._settings['excel.status.cell'])
        self.current_status_excel_col_index: int = int(self.get_letter_position(col))
        self.current_status_excel_row_index: int = int(row)

        self._wait_for_window('Pending Tray')
        self._window_title_stack.append('Pending Tray')

        self.current_element_count = 0
        self.total_element_size = len(shipments)

        for i, shipment in enumerate(shipments):
            logger.info("Start process shipment " + shipment)

            try:
                if status_address[i] != "ADDRESS MATCHED":
                    logger.info(f"Skipping shipment {shipment} due to status: {status_address[i]}")
                    self.input_status_into_excel('Skip')
                    self.excel_provider.save(workbook)
                    continue

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

            except Exception:

                self.input_status_into_excel('An exception error')
                logger.info(f'Cannot handle shipment {shipment}. Moving to next shipment')
                continue

        self.excel_provider.close(workbook)

    def process_on_each_shipment(self, shipment):

        """Loop for each shipment in this method.
        Interface with 2nd window (Ctrl R - to get CNEE name)
        3rd window (Ctrl G - Maintain Pricing),
        4th window (Maintain Invoice Details).
        Return Payment Term, Invoice party index with CNEE name and Shipment in file Excel"""

        logger: Logger = get_current_logger()

        GCSS_Shipment_MSL_Active_Title: str = self._wait_for_window(shipment)
        self._window_title_stack.append(GCSS_Shipment_MSL_Active_Title)
        autoit.win_activate(GCSS_Shipment_MSL_Active_Title)

        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        self.into_parties_tab()

        'Get Cnee Name in 2nd window'
        cnee_name: str = self.get_consignee()
        if cnee_name is None:
            message_cnee = f'Cannot find CNEE name of the shipment {shipment} in 1st window - Ctrl R'
            logger.info(message_cnee)
            self.input_status_into_excel(message_cnee)
            self._close_windows_util_reach_first_gscc()
            return

        invoice_parties: set[str] = self.get_invoice_parties()
        credit_parties: set[str] = self.get_credit_parties()
        if not self.validate_parties_values(cnee_name, invoice_parties, credit_parties):
            logger.info('Invoice and Credit parties not correct, re-updating these parties')
            self.adding_invoice_and_credit_parties()
            logger.info('Updated Invoice and Credit Parties')

        # 'Get Invoice Tab'
        self.into_freight_and_pricing_tab()

        # 'Get the tab Invoice, then click the Modify button in 2nd window'
        Tab_controls = self._window.children(class_name="SysTabControl32")[0]
        Tab_controls.select(tab=1)
        self._hotkey_then_open_new_window('Maintain Pricing and Invoicing', 'alt', 'y')

        # 'Open and interface with 3rd window - Maintain Pricing and Invoicing Window'
        Tab_controls_maintain_pricing_and_inv = self._window.children(class_name="SysTabControl32")[0]
        Tab_controls_maintain_pricing_and_inv.select(tab=1)

        self.click_all_item_payment_term_collect()

        # Into Maintain Invoice Details
        self._hotkey_then_open_new_window('Maintain Invoice Details', 'alt', 'i')

        ComboBox_maintain_payment: ComboBoxWrapper = self._window.children(class_name="ComboBox")[0]
        ComboBox_maintain_payment.select('Collect')

        'when choosing this payment term, it will automate show a dialog Information'
        dialog_title_infor = 'Information'
        dialog_window = self._app.window(title=dialog_title_infor)
        try:
            dialog_window.wait(timeout=5, wait_for='visible')
            dialog_window.type_keys('{ENTER}')
        except Exception:
            pyautogui.hotkey('tab')

        ComboBox_maintain_invoice_party: ComboBoxWrapper = self._window.children(class_name="ComboBox")[1]
        item_invoices: list[str] = ComboBox_maintain_invoice_party.item_texts()

        runner: int = 0
        for invoice in item_invoices:
            half_index: int = int(len(cnee_name) / 2)
            start_cnee_name: str = cnee_name[0:half_index]

            if invoice.startswith(start_cnee_name):
                #
                ComboBox_maintain_invoice_party.select(runner)
                dialog_title_qs = 'Question'
                dialog_window = self._app.window(title=dialog_title_qs)
                dialog_window.wait(timeout=10, wait_for='visible')
                dialog_window.type_keys('{ENTER}')
                break

            runner += 1

        ComboBox_maintain_collect_business: ComboBoxWrapper = self._window.children(class_name="ComboBox")[3]
        ComboBox_maintain_collect_business.select('Maersk Bangkok (Bangkok)')

        ComboBox_maintain_printable_freight_line: ComboBoxWrapper = self._window.children(class_name="ComboBox")[4]
        ComboBox_maintain_printable_freight_line.select('Yes')

        # Click OK button in 4th window and window will be auto closed
        self._hotkey_then_close_current_window('alt', 'k')

        # Click complete collect button _ in Maintain Pricing and Invoicing window
        collect_details_collect_status: EditWrapper = self._window.children(class_name="Edit")[3]
        while True:
            pyautogui.hotkey('alt', 't')
            if collect_details_collect_status.texts()[0] == 'Yes':
                break

            self.sleep()

        # Input Excel
        status_shipment: str = 'Done'
        self.input_status_into_excel(status_shipment)

        self._close_windows_util_reach_first_gscc()

    def get_letter_position(self, letter: str):
        if letter.__len__() != 1:
            raise Exception('Please provide the single char, not a word ! the previous input was ' + letter)
        letter_to_position = {chr(i + 96): i for i in range(1, 27)}

        return letter_to_position.get(letter.lower(), "Invalid input")

    def input_status_into_excel(self, message):
        self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                            self.current_status_excel_col_index, message)

    def click_all_item_payment_term_collect(self) -> int:
        """"
            return the number of item payment term collect have been clicked
        """
        logger: Logger = get_current_logger()
        list_views: ListViewWrapper = self._window.children(class_name="SysListView32")[1]
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
        self._close_windows_util_reach_first_gscc()
        return 0

    def into_freight_and_pricing_tab(self):
        while True:
            pyautogui.hotkey('ctrl', 'g')

            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")
            if list_views.__len__() == 6:
                break

            self.sleep()

    def into_parties_tab(self):

        while True:
            pyautogui.hotkey('ctrl', 'r')
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")
            if list_views.__len__() == 2:
                break

            self.sleep()

    def validate_parties_values(self, consignee: str, invoice_parties: set[str], cre_parties: set[str]) -> bool:

        if not invoice_parties.__contains__(consignee):
            return False

        if not cre_parties.__contains__(consignee):
            return False

        return True

    def get_invoice_parties(self) -> set[str]:
        inv_parties_UI_elements: list[_listview_item] = self.get_party_info("Invoice Party")
        numbers_iter: Iterator[_listview_item] = iter(inv_parties_UI_elements)
        inv_parties_values = set(map(lambda inv_party: inv_party.text(), numbers_iter))
        return inv_parties_values

    def get_credit_parties(self) -> set[str]:
        cre_parties_UI_elements: list[_listview_item] = self.get_party_info("Credit Party")
        numbers_iter: Iterator[_listview_item] = iter(cre_parties_UI_elements)
        cre_parties_values = set(map(lambda inv_party: inv_party.text(), numbers_iter))
        return cre_parties_values

    def get_consignee(self) -> str:
        consignees: list[_listview_item] = self.get_party_info("Consignee")
        if len(consignees) == 0:
            return None

        consignees[0].select()
        return consignees[0].text()

    def get_party_info(self, party_name: str) -> list[_listview_item]:
        logger: Logger = get_current_logger()
        list_views: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

        items = list_views.items()
        count_row = 0

        party_values: list[_listview_item] = []

        for item in items:
            item: _listview_item

            if str(item.text()) == party_name:
                party_values.append(items[count_row - 1])
                count_row += 1
                continue

            count_row += 1

        logger.debug(f'Get {party_name} with value {party_values}')
        return party_values

    def adding_invoice_and_credit_parties(self):
        cnee_edit_element: EditWrapper = self._window.children(class_name="Edit")[14]
        cnee_scv_no: str = cnee_edit_element.texts()

        self._window = self._hotkey_then_open_new_window('Party Details', 'alt', 'a')

        self._window = self._hotkey_then_open_new_window('Customer Search', 'alt', 'c')

        pyautogui.typewrite(cnee_scv_no[0])
        ComboBox_Customer_ID: ComboBoxWrapper = self._window.children(class_name="ComboBox")[0]
        ComboBox_Customer_ID.select('Customer ID')

        # back to Party Details
        self._window = self._hotkey_then_close_current_window('alt', 's')
        list_views: ListViewWrapper = self._window.children(class_name="SysListView32")[0]
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

        # click btn >>
        list_btn: list[ButtonWrapper] = self._window.children(class_name="Button")
        button_right: ButtonWrapper = list_btn[3]
        button_right.click()

        # Click OK button to close Party Details
        self._window = self._hotkey_then_close_current_window('alt', 'k')

        try:
            validation_failed_window = self._window.child_window(title="Validation failed", control_type="Window",
                                                                 found_index=0)
            if validation_failed_window.exists(timeout=1):
                self.input_status_into_excel('Exception error when adding inv & cre parties')
                self._close_windows_util_reach_first_gscc()
                return
        except Exception as e:
            pass
