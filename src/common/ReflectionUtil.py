import importlib
from types import ModuleType
from typing import Callable

from src.task.AutomatedTask import AutomatedTask


def create_task_instance(setting_states: dict[str, str], task_name: str,
                         callback_before_run_task: Callable[[], None]) -> AutomatedTask:
    clazz_module: ModuleType = importlib.import_module('src.task.' + task_name)
    clazz = getattr(clazz_module, task_name)
    automated_task: AutomatedTask = clazz(setting_states, callback_before_run_task)
    return automated_task
