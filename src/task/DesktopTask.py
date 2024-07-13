from abc import ABC
from logging import Logger
from typing import Callable, Any

import pyautogui
import pygetwindow as gw
from pygetwindow import Win32Window
from pywinauto import Application, WindowSpecification, ElementNotFoundError

from src.common.Stack import Stack
from src.common.ThreadLocalLogger import get_current_logger
from src.task.AutomatedTask import AutomatedTask


class DesktopTask(AutomatedTask, ABC):

    def __init__(self,
                 settings: dict[str, str],
                 callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self._window_title_stack: Stack[str] = Stack[str]()
        self._app: Application = None
        self._window: WindowSpecification = None

    def _is_current_window_having_title(self, expected_title: str) -> bool:
        window: Win32Window = gw.getActiveWindow()

        if str(window.title).__contains__(expected_title):
            return True

        return False

    def _wait_for_window(self, title):
        max_attempt: int = 30
        current_attempt: int = 0

        while current_attempt < max_attempt:

            window_titles: list[str] = gw.getAllTitles()

            for window_title in window_titles:
                if window_title.__contains__(title):
                    gw.getWindowsWithTitle(window_title)[0].activate()
                    return window_title

            current_attempt += 1
            self.sleep()

        raise Exception('Can not find out the asked window {}'.format(title))

    def _hotkey_then_close_current_window(self, *args: Any) -> WindowSpecification:
        self._window_title_stack.pop()
        current_window_title: str = self._window_title_stack.peek()
        return self.__hotkey_then_activate_window_by_title(current_window_title, args)

    def _hotkey_then_open_new_window(self, window_title: str, *args: Any) -> WindowSpecification:
        if window_title is None or len(window_title) == 0:
            raise Exception('Invalid window title')

        self._window_title_stack.append(window_title)
        new_window_title: str = self._window_title_stack.peek()
        return self.__hotkey_then_activate_window_by_title(new_window_title, args)

    def __hotkey_then_activate_window_by_title(self, window_title, args) -> WindowSpecification:
        counter: int = 0
        while not self._is_current_window_having_title(window_title):

            if counter > 10:
                raise Exception(
                    f"Time is over for waiting {window_title} appear by the hotkey combination {args}")

            pyautogui.hotkey(*args)
            self.sleep()
            counter += 1

        gw.getWindowsWithTitle(window_title)[0].activate()
        self._window = self._app.window(title=self._window_title_stack.peek())
        return self._window

    def _close_windows_util_reach_first_gscc(self):
        return self._close_windows_with_window_title_stack(self._window_title_stack)

    def _close_windows_with_window_title_stack(self, window_title_stack: list[str]):
        logger: Logger = get_current_logger()

        for i in range(window_title_stack.__len__() - 1, 0, -1):
            window_title = window_title_stack[i]

            try:
                app = Application().connect(title=window_title, timeout=10)

                window = app.window(title=window_title)

                if window.exists(timeout=1):
                    window.close()
                    logger.debug(f"Window with title '{window_title}' has been closed.")
                else:
                    logger.debug(f"Window with title '{window_title}' does not exist.")

                window_title_stack.pop()

            except ElementNotFoundError:
                logger.debug(f"No window with title '{window_title}' was found.")
            except Exception as e:
                logger.debug(f"An error occurred: {e}")

        self._wait_for_window(window_title_stack[0])
