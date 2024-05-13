import time
from logging import Logger

import autoit
import pygetwindow as gw
from pywinauto import Application
from pywinauto.controls.common_controls import ListViewWrapper

from src.common.ThreadLocalLogger import get_current_logger


class GCSS_Automate():
    def __init__(self):
        self.automate()

    # def mandatory_settings(self) -> list[str]:
    #     mandatory_keys: list[str] = ['username', 'password', 'excel.path', 'excel.sheet',
    #                                  'excel.column.bill']
    #     return mandatory_keys

    # def __init__(self):
    #
    #     with Display():  # Start a virtual display
    #         self.automate()

    def automate(self) -> None:
        logger: Logger = get_current_logger()
        # self.invoke_the_app()
        #
        # # Wait for a moment before sending the key press
        # self.wait_for_window('Select User Profile')
        #
        # # Simulate a mouse click at the specified coordinates
        # pyautogui.click(973, 568)
        #
        # self.wait_for_window('Pending Tray')
        #
        # # Simulate sending the "a" key
        # pyautogui.hotkey('ctrl', 'o')
        # print('Hello')

        window_title = "GCSS - 235734521 - MSL - Active"
        # autoit.win_activate(window_title)

        self.get_items_in_tables(window_title)

    def get_items_in_tables(self, window_title):

        window_title = window_title

        app = Application().connect(
            title=window_title)

        # Get the main window
        window = app.window(
            title=window_title)

        # Activate the window
        window.set_focus()
        for child in window.children():
            print(f"Class Name: {child.class_name()}, Text: {child.window_text()}, Control ID: {child.control_id()}")
        list_views: list[ListViewWrapper] = window.children(class_name="SysListView32")
        # window.send_keys('+')
        # Iterate through the found SysListView32 components and print their details
        for idx, list_view in enumerate(list_views):
            list_view: ListViewWrapper
            print(
                f"Component {idx + 1}: Class Name: {list_view.class_name()}, Control ID: {list_view.control_id()}, Text: {list_view.window_text()}")

            abc = list_view.get_item(1)
            print(abc.text())
            list_view.texts()

            for item in list_view.items():
                # print(item.text())
                if str(item.text()).__contains__('Collect'):
                    item.check()
            print('abc')

    def wait_for_window(self, title):
        max_attempt: int = 30
        current_attempt: int = 0

        while current_attempt < max_attempt:
            print('Still waiting')

            # Get a list of all open window titles
            window_titles = gw.getAllTitles()

            for window_title in window_titles:
                if window_title.__contains__(title):
                    autoit.win_activate(window_title)
                    return window_title

            current_attempt += 1
            time.sleep(1)

        raise Exception('Can not find out the asked window')

    def invoke_the_app(self):
        app_process_name = "GCSSExport.exe"
        if autoit.process_exists(app_process_name):
            print(f"{app_process_name} is already running.")

            # Activate the application window
            # You can use window title, class, handle, etc., to specify the window
            window_title = "GCSS"
            autoit.win_activate(window_title)
            return

        print(f"{app_process_name} is not running.")
        # Path to the .exe file
        exe_path = r"C:\Program Files (x86)\GCSS\PROD_A\{}".format(app_process_name)

        # Command-line arguments
        arguments = "-wsnaddr=//gcssexport1.gls.dk.eur.crb.apmoller.net:15000"

        # Run the .exe file with arguments
        autoit.run(f'"{exe_path}" {arguments}')


if __name__ == '__main__':
    instance = GCSS_Automate()
    instance.automate()
