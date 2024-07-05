import logging
import time
import uuid
from abc import abstractmethod, ABC
from datetime import datetime
from logging import Logger
from typing import Callable

from src.common.Percentage import Percentage
from src.common.ResumableThread import ResumableThread
from src.common.StringUtil import validate_keys_of_dictionary
from src.common.ThreadLocalLogger import get_current_logger, create_thread_local_logger


class AutomatedTask(Percentage, ResumableThread, ABC):
    @abstractmethod
    def mandatory_settings(self) -> list[str]:
        pass

    @abstractmethod
    def automate(self):
        pass

    @property
    def settings(self) -> dict[str, str]:
        return self._settings

    @settings.setter
    def settings(self, settings: dict[str, str]):
        if settings is None:
            raise ValueError("Provided settings is None")
        self._settings = settings

    def __init__(self,
                 settings: dict[str, str],
                 callback_before_run_task: Callable[[], None]):
        super().__init__()
        logger: Logger = get_current_logger()
        self._settings: dict[str, str] = settings
        self.callback_before_run_task = callback_before_run_task

        if self._settings.get('time.unit.factor') is None:
            self._timingFactor = 1.0
        else:
            self._timingFactor = float(self._settings.get('time.unit.factor'))

        if self._settings.get('use.GUI') is None:
            self.use_gui = False
        else:
            self.use_gui = 'True'.lower() == str(self._settings.get('use.GUI')).lower()

        if not self.use_gui:
            logger.info('Run in headless mode')

    def perform(self) -> None:
        mandatory_settings: list[str] = self.mandatory_settings()
        mandatory_settings.append('invoked_class')
        validate_keys_of_dictionary(self._settings, set(mandatory_settings))

        logger: Logger = create_thread_local_logger(class_name=self._settings['invoked_class'],
                                                    thread_uuid=str(uuid.uuid4()))

        if self.callback_before_run_task is not None:
            self.callback_before_run_task()

        try:
            self.automate()
        except Exception as exception:
            logger.exception(str(exception))
        logger.info("Done task. It ends at {}".format(datetime.now()))
        del logging.Logger.manager.loggerDict[self._settings['invoked_class']]

    def perform_mainloop_on_collection(self,
                                       collection,
                                       critical_operation_on_each_element: Callable[[object], None]):
        self.current_element_count = 0
        self.total_element_size = len(collection)
        logger: Logger = get_current_logger()

        for each_element in collection:

            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    logger.info("Currently pause")
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

            critical_operation_on_each_element(each_element)
            self.current_element_count = self.current_element_count + 1

    def sleep(self) -> None:
        time.sleep(self._timingFactor)
        return
