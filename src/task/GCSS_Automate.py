import time
from logging import Logger
from typing import Callable

import autoit
import pygetwindow as gw
import pywinauto.keyboard
from pywinauto import Application, WindowSpecification
from pywinauto.controls.common_controls import ListViewWrapper, _listview_item
from pywinauto.controls.win32_controls import ButtonWrapper, ComboBoxWrapper
from pywinauto.findwindows import find_element

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.AutomatedTask import AutomatedTask


class GCSS_Automate(AutomatedTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self.excel_provider: ExcelReaderProvider = None
        self.current_row: int = 0
        self.window_title_stack: list[str] = None

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['excel.path', 'excel.sheet',
                                     'excel.shipment', 'excel.status.col', 'excel.status.row']
        return mandatory_keys

    def automate(self):
        logger: Logger = get_current_logger()
        self.excel_provider: ExcelReaderProvider = XlwingProvider()
        path_to_excel = self._settings['excel.path']
        workbook = self.excel_provider.get_workbook(path=path_to_excel)
        logger.info('Loading excel files')

        sheet_name: str = self._settings['excel.sheet']
        worksheet = self.excel_provider.get_worksheet(workbook, sheet_name)

        shipments: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                     self._settings['excel.sheet'],
                                                                     self._settings['excel.shipment'])

        status_starting_column: int = int(self._settings['excel.status.col'])
        status_starting_row: int = int(self._settings['excel.status.row'])
        self.current_row = status_starting_row

        # self.excel_provider.change_value_at(worksheet=worksheet, row=self.current_row, column=status_starting_column,
        #                                     value="hung")
        self._input_excel("hung")
        self.wait_for_window('Pending Tray')

        self.current_element_count = 0
        self.total_element_size = len(shipments)

        for shipment in shipments:

            if self.terminated is True:
                return
            with self.pause_condition:
                while self.paused:
                    logger.info("Currently pause")
                    self.pause_condition.wait()
                if self.terminated is True:
                    return

            pywinauto.keyboard.send_keys('^o')

            pywinauto.keyboard.send_keys(shipment)
            pywinauto.keyboard.send_keys('{TAB}')
            pywinauto.keyboard.send_keys('{ENTER}')

            window_title_second: str = self.wait_for_window(shipment)
            autoit.win_activate(window_title_second)

            self.get_items_in_tables(window_title_second, shipment, worksheet, self.current_row,
                                     status_starting_column)
            self.current_element_count += 1

            status_shipment: str = 'Done shipment {}'.format(shipment)
            logger.info(status_shipment)
            # self.excel_provider.change_value_at(worksheet, self.current_row, status_starting_column,
            #                                     status_shipment)
            self._input_excel(status_shipment)
            self.excel_provider.save(workbook)
            self.excel_provider.close(workbook)
            self.current_row += 1

    def get_items_in_tables(self, window_title_second, shipment, worksheet, status_starting_row,
                            status_starting_column):

        """Loop for each shipment in this method.
        Interface with 2nd window (Ctrl R - to get CNEE name)
        3rd window (Ctrl G - Maintain Pricing),
        4th window (Maintain Invoice Details). Return Payment Term, Invoice party index
        with CNEE name and Shipment in file Excel"""

        self.window_title_stack = ['Pending Tray', window_title_second, 'Maintain Pricing and Invoicing',
                                   'Maintain Invoice Details']

        logger: Logger = get_current_logger()

        app: Application = Application().connect(
            title=window_title_second)
        window: WindowSpecification = app.window(title=window_title_second)

        pywinauto.keyboard.send_keys('^r')

        'Get Cnee Name in 2nd window'
        cnee_name: str = self.get_consignee(window, shipment, worksheet, self.current_row,
                                            status_starting_column)
        if cnee_name is None:
            self.close_windows_util_reach_first_gscc()
            return

        pywinauto.keyboard.send_keys('^g')
        time.sleep(0.5)

        'Get the tab Invoice, then click the Modify button in 2nd window'
        self.get_modify_button(window, shipment, worksheet, status_starting_row, status_starting_column)
        if self.get_modify_button is None:
            # If not found modify button in here, close 2nd window and continue next shipment
            self.close_window_by_title(window_title_second)
            return

        'Open and interface with 3rd window - Maintain Pricing and Invoicing Window'
        window_title_third: str = 'Maintain Pricing and Invoicing'
        window_third: WindowSpecification = app.window(title=window_title_third)

        self.get_maintain_pricing_invoice_tab(window_third, shipment, worksheet, self.current_row,
                                              status_starting_column)
        if self.get_modify_button is None:
            self.close_window_by_title(window_title_third)
            return

        pywinauto.keyboard.send_keys('{VK_CONTROL down}')

        list_views: ListViewWrapper = window_third.children(class_name="SysListView32")[1]
        count_item: int = 0
        for item in list_views.items():
            item: _listview_item
            if str(item.text()).__contains__('Collect'):
                item.select()
                count_item += 1

        if count_item == 0:
            status_collect = 'This shipment have no payment term Collect'

            # self.excel_provider.change_value_at(worksheet, self.current_row, status_starting_column,
            #                                     status_collect)
            self._input_excel(status_collect)
            self.close_windows_util_reach_first_gscc()
            return

        pywinauto.keyboard.send_keys('{VK_CONTROL up}')

        try_press_i_to_open_maintain_invoice: int = 0
        try:
            if try_press_i_to_open_maintain_invoice > 10:
                raise Exception
            pywinauto.keyboard.send_keys('i')
            time.sleep(0.5)
            try_press_i_to_open_maintain_invoice += 1
        except:
            message_choose_row_collect = 'Cannot Open window Maintain Invoice Details'
            logger.info(message_choose_row_collect)
            # self.excel_provider.change_value_at(worksheet, self.current_row, status_starting_column,
            #                                     message_choose_row_collect)

            self._input_excel(message_choose_row_collect)
            self.close_window_by_title(window_title_third)

            return None

        'Open 4th Window - Maintain Invoice Details'
        self.wait_for_window('Maintain Invoice Details')

        window_title_maintain: str = 'Maintain Invoice Details'
        window_maintain: WindowSpecification = app.window(title=window_title_maintain)

        ComboBox_maintain_payment: ComboBoxWrapper = window_maintain.children(class_name="ComboBox")[0]
        ComboBox_maintain_payment.select('Prepaid')

        'when choosing this payment term, it will automate show a dialog Information'
        dialog_title_infor = 'Information'
        dialog_window = app.window(title=dialog_title_infor)
        dialog_window.wait(timeout=10, wait_for='visible')
        dialog_window.type_keys('{ENTER}')

        ComboBox_maintain_invoice_party: ComboBoxWrapper = window_maintain.children(class_name="ComboBox")[1]
        for item in ComboBox_maintain_invoice_party.item_texts():
            if cnee_name in item:
                ComboBox_maintain_invoice_party.select(item)
                break
        else:
            raise ValueError('No item containing {} found in ComboBox'.format(cnee_name))

        dialog_title_qs = 'Question'
        dialog_window = app.window(title=dialog_title_qs)
        dialog_window.wait(timeout=10, wait_for='visible')
        dialog_window.type_keys('{ENTER}')

        ComboBox_maintain_collect_business: ComboBoxWrapper = window_maintain.children(class_name="ComboBox")[3]
        ComboBox_maintain_collect_business.select('Maersk Bangkok (Bangkok)')

        ComboBox_maintain_printable_freight_line: ComboBoxWrapper = window_maintain.children(class_name="ComboBox")[4]
        ComboBox_maintain_printable_freight_line.select('Yes')

        # pyautogui.hotkey('ctrl', 'k')
        # _

    def wait_for_window(self, title):
        max_attempt: int = 30
        current_attempt: int = 0

        while current_attempt < max_attempt:
            print('Still waiting')

            # Get a list of all open window titles
            window_titles: list[str] = gw.getAllTitles()

            for window_title in window_titles:
                if window_title.__contains__(title):
                    autoit.win_activate(window_title)
                    return window_title

            current_attempt += 1
            time.sleep(1)

        raise Exception('Can not find out the asked window {}'.format(title))

    def get_consignee(self, window: WindowSpecification, shipment, worksheet, status_starting_row,
                      status_starting_column) -> str | None:

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
            message = 'Cannot find CNEE name of the shipment {} in 1st window - Ctrl R'.format(shipment)
            logger.info(message)
            # self.excel_provider.change_value_at(worksheet, self.current_row, status_starting_column, message)
            self._input_excel(message)
            return None

    def get_modify_button(self, window: WindowSpecification, shipment, worksheet, status_starting_row,
                          status_starting_column):

        logger: Logger = get_current_logger()

        try_to_get_modify_button: int = 0

        try:
            if try_to_get_modify_button > 30:
                raise Exception

            Tab_controls = window.children(class_name="SysTabControl32")[0]
            Tab_controls.select(tab=1)

            Modify_button_in_invoices: list[ButtonWrapper] = window.children(class_name="Button")

            for button in Modify_button_in_invoices:

                if button.texts()[0] == 'Modif&y':
                    button.click()
                    break

                time.sleep(0.5)
                try_to_get_modify_button += 1

        except:
            message_button = 'Cannot click Modify in 2nd window - Ctrl G'
            logger.info(message_button)
            # self.excel_provider.change_value_at(worksheet, self.current_row, status_starting_column, message_button)
            self._input_excel(message_button)
            return None

    def get_maintain_pricing_invoice_tab(self, window: WindowSpecification, shipment, worksheet,
                                         status_starting_row,
                                         status_starting_column):
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
            # self.excel_provider.change_value_at(worksheet, self.current_row, status_starting_column, message_pricing)
            self._input_excel(message_pricing)
            return None

    def close_windows_util_reach_first_gscc(self):

        # Iterate over the indices 3, 2, and 1 in reverse order
        for i in range(self.window_title_stack.__len__() - 1, 0, -1):
            self.close_window_by_title(self.window_title_stack[i])

        window_titles: list[str] = gw.getAllTitles()
        for abc in window_titles:
            # Get a list of all open window titles
            focused_window = find_element(active_only=True)

            # Get the title of the focused window
            focused_window_title: str = focused_window.name

            if focused_window_title.__contains__(self.window_title_stack[0]):
                break

            self.switch_to_next_window()

    def close_window_by_title(self, title):
        window_titles: list[str] = gw.getAllTitles()
        for abc in window_titles:
            # Get a list of all open window titles
            focused_window = find_element(active_only=True)

            # Get the title of the focused window
            focused_window_title: str = focused_window.name

            if focused_window_title.__contains__(title):
                pywinauto.keyboard.send_keys('%{F4}')
                break

            self.switch_to_next_window()

    def switch_to_next_window(self):
        """
        Switches to the next visible window.
        """
        current_window = find_element(active_only=True).name

        visible_windows: list[str] = gw.getAllTitles()

        if current_window in visible_windows:
            current_index = visible_windows.index(current_window)
            next_index = (current_index + 1) % len(visible_windows)
            next_window = visible_windows[next_index]

            # Set focus to the next window
            autoit.win_activate(next_window)

    def _input_excel(self, status_shipment: str):
        logger: Logger = get_current_logger()

        path_to_excel = self._settings['excel.path']
        workbook = self.excel_provider.get_workbook(path=path_to_excel)
        logger.info('Loading excel files')

        sheet_name: str = self._settings['excel.sheet']
        worksheet = self.excel_provider.get_worksheet(workbook, sheet_name)

        status: str = status_shipment

        for index, cell in enumerate(worksheet['A'], start=1):
            worksheet[f'B{index}'] = status
            workbook.save(path_to_excel)


if __name__ == "__main__":
    settings: dict[str, str] = {}
    abc = GCSS_Automate(settings, None)
    print('Hung')
