from logging import Logger
from typing import Callable

import autoit

from src.common.ThreadLocalLogger import get_current_logger
from src.task.AutomatedTask import AutomatedTask


class GCSS_Automate(AutomatedTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['username', 'password', 'excel.path', 'excel.sheet',
                                     'excel.column.bill']
        return mandatory_keys

    def automate(self) -> None:
        logger: Logger = get_current_logger()
        logger.info(
            "---------------------------------------------------------------------------------------------------------")
        logger.info("Start processing")

        # Wait for the main window to appear
        autoit.win_wait_active("Pending Tray - GCSS", 10)

        # Send keys to open file dialog
        autoit.send("^o")

        # Wait for the new window to appear
        autoit.win_wait_active("Open Shipment", 10)

        # Get the handle of the "Edit" field in the "Open Shipment" window
        edit_handle = autoit.control_get_handle("Open Shipment", "", "[CLASS:Edit; INSTANCE:1]")

        # Set text to the right field in the new window using the handle
        autoit.control_send("Open Shipment", "", edit_handle, "235834169")

        # Click the OK button in the new window
        autoit.control_click("Open Shipment", "", "[CLASS:Button; TEXT:OK]")
