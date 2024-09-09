from logging import Logger
from typing import Callable

import pyautogui
import pygetwindow as gw
from pywinauto import Application, WindowSpecification
from pywinauto.controls.common_controls import ListViewWrapper, _listview_item

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.DesktopTask import DesktopTask


class GCSS_SPIR(DesktopTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self.excel_provider: ExcelReaderProvider = None
        self.current_worksheet = None
        self.current_status_excel_row_index: int = 5

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['excel.path', 'excel.sheet', 'excel.shipment', 'excel.status.trucking',
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

        self._wait_for_window('Pending Tray')
        self._window_title_stack.append('Pending Tray')

        self.current_element_count = 0
        self.total_element_size = len(shipments)

        for i, shipment in enumerate(shipments):

            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    logger.info("Currently pause")
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

            logger.info("Start process shipment " + shipment)

            try:
                pyautogui.hotkey('ctrl', 'o')
                pyautogui.typewrite(shipment)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('enter')

                try:
                    popup_window: str = self._wait_for_window('Invalid Booking Number')
                    self._window_title_stack.append(popup_window)
                    gw.getWindowsWithTitle(popup_window)[0].activate()

                    self._app: Application = Application().connect(title=self._window_title_stack.peek())
                    self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

                    if popup_window:
                        pyautogui.hotkey('enter')
                        pyautogui.hotkey('shift', 'tab')
                        pyautogui.hotkey('shift', 'tab')
                        pyautogui.hotkey('down')
                        pyautogui.hotkey('alt', 'k')

                        self.sleep()
                        self.process_on_each_shipment_adhoc(shipment)
                        self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                            3, 'Done')
                        self.excel_provider.save(workbook)

                        logger.info("Done with shipment " + shipment)
                        continue

                except:
                    logger.debug('No popup found to failed handle it')

                self.sleep()
                self.process_on_each_shipment(shipment)
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3, 'Done')
                self.excel_provider.save(workbook)

                logger.info("Done with shipment " + shipment)

            except Exception:

                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3, 'To refresh GCSS (F5) message prompted')
                self.excel_provider.save(workbook)
                logger.info(f'Not handle shipment {shipment}. Moving to next shipment')
                self._close_windows_util_reach_first_gscc()
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                continue

            self.current_status_excel_row_index += 1
            self.current_element_count += 1

        self.excel_provider.save(workbook)
        self.excel_provider.close(workbook)

    def process_on_each_shipment(self, shipment):
        logger: Logger = get_current_logger()

        GCSS_Shipment_MSL_Active_Title: str = self._wait_for_window(shipment)
        self._window_title_stack.append(GCSS_Shipment_MSL_Active_Title)
        gw.getWindowsWithTitle(GCSS_Shipment_MSL_Active_Title)[0].activate()

        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        list_views = self._window.children(class_name="SysListView32")[1]

        runner = 0
        array = [None for _ in range(8)]
        for item in list_views.items():

            array[runner] = item.text()

            if runner != 7:
                runner = runner + 1
                continue

            runner = 0
            if 'LOAD' in array[3]:
                self.into_activity_shipment()
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2, 'Load')
                continue

        # Input Excels
        self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                            3, 'Done')
        self._close_windows_util_reach_first_gscc()

    def process_on_each_shipment_adhoc(self, shipment):
        logger: Logger = get_current_logger()

        GCSS_Shipment_MSL_Active_Title: str = self._wait_for_window(shipment)
        self._window_title_stack.append(GCSS_Shipment_MSL_Active_Title)
        gw.getWindowsWithTitle(GCSS_Shipment_MSL_Active_Title)[0].activate()

        pyautogui.hotkey('ctrl', 'k')

        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        list_views = self._window.children(class_name="SysListView32")[0]

        runner = 0
        array = [None for _ in range(8)]
        for item in list_views.items():

            array[runner] = item.text()

            if runner != 7:
                runner = runner + 1
                continue

            runner = 0
            if 'LOAD' in array[3]:
                self.into_activity_shipment()
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2, 'Load')
                continue

        # Input Excels
        self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                            3, 'Done')
        self._close_windows_util_reach_first_gscc()

    def into_activity_shipment(self):

        while True:
            pyautogui.hotkey('ctrl', 't')
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")

            if list_views.__len__() == 1:
                break
            self.sleep()

        runner = 0
        array = [None for _ in range(6)]
        list_of_activity_plan: list[_listview_item] = []
        list_of_activity_plan_split: list[_listview_item] = []
        listview_activity: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

        for item in listview_activity.items():
            array[runner] = item

            if runner != 5:
                runner = runner + 1
                continue

            runner = 0
            # if array[0].text().startswith('OPS') and array[4].text() == 'Close':
            #     continue

            if array[0].text().startswith('OPS (EQUIPMENT PICKUP)') and array[4].text() == 'Open':
                list_of_activity_plan.append(array[0])

            if array[0].text().startswith('OPS (ICD') and array[4].text() == 'Open':
                list_of_activity_plan.append(array[0])

            if array[0].text().startswith('OPS (EQUIPMENT PICKUP)') and array[4].text() == 'Closed':
                list_of_activity_plan_split.append(array[0])

            if array[0].text().startswith('OPS (ICD') and array[4].text() == 'Closed':
                list_of_activity_plan_split.append(array[0])

        if len(list_of_activity_plan) != 3 and len(list_of_activity_plan_split) != 3:
            self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                3, 'Activity closed, Only update Invoice')
            return

        for activity_plan in list_of_activity_plan:
            activity_plan.select()
            pyautogui.hotkey('alt', 'l')
            activity_plan.deselect()
            self.sleep()

        pyautogui.hotkey('alt', 'i')
        self.sleep()

        pyautogui.hotkey('q')
        pyautogui.hotkey('p')

        self._close_windows_util_reach_first_gscc()

        if len(list_of_activity_plan) != 3 and len(list_of_activity_plan_split) != 3:
            self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                3, 'Activity closed, recheck')
            return
