import time
from logging import Logger
from typing import Callable, Set

from src.task.AutomatedTask import AutomatedTask
from src.common.ThreadLocalLogger import get_current_logger


class ExampleTask(AutomatedTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['username', 'password', 'excel.path', 'excel.sheet',
                                     'excel.read_column.start_cell', 'hung.path']
        return mandatory_keys

    def automate(self):
        booking_ids = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
        self.perform_mainloop_on_collection(booking_ids, ExampleTask.operation_on_each_element)

    @staticmethod
    def operation_on_each_element(booking):
        logger: Logger = get_current_logger()
        # selenium tasks there
        logger.info("Example automated task - running at booking {}".format(booking))
        time.sleep(2)
