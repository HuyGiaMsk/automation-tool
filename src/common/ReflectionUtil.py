import importlib
from types import ModuleType
from typing import Callable

from src.common.FileUtil import find_module
from src.setup.packaging.path.PathResolvingService import PathResolvingService
from src.task.AutomatedTask import AutomatedTask

cache: dict[str, ModuleType] = {}


def create_task_instance(setting_states: dict[str, str], task_name: str,
                         callback_before_run_task: Callable[[], None]) -> AutomatedTask:
    if cache.get(task_name):
        clazz = getattr(cache.get(task_name), task_name)
        automated_task: AutomatedTask = clazz(setting_states, callback_before_run_task)
        return automated_task

    module_path = find_module(PathResolvingService.get_instance().get_task_dir(), task_name)
    if module_path is None:
        raise FileNotFoundError(f"Task file {task_name}.py not found in {module_path} and its subdirectories.")

    clazz_module: ModuleType = importlib.import_module(module_path)
    cache[task_name] = clazz_module
    clazz = getattr(clazz_module, task_name)
    automated_task: AutomatedTask = clazz(setting_states, callback_before_run_task)
    return automated_task
