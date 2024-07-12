import importlib
from types import ModuleType
from typing import Callable

from src.common.FileUtil import find_module
from src.setup.packaging.path.PathResolvingService import TASK_DIR
from src.task.AutomatedTask import AutomatedTask

cache: dict[str, ModuleType] = {}


def create_task_instance(setting_states: dict[str, str], task_name: str,
                         callback_before_run_task: Callable[[], None]) -> AutomatedTask:
    if cache.get(task_name):
        clazz = getattr(cache.get(task_name), task_name)
        automated_task: AutomatedTask = clazz(setting_states, callback_before_run_task)
        return automated_task

    module_path = find_module(TASK_DIR, task_name)
    if module_path is None:
        raise FileNotFoundError(f"Task file {task_name}.py not found in {module_path} and its subdirectories.")

    clazz_module: ModuleType = importlib.import_module(module_path)
    cache[task_name] = clazz_module
    clazz = getattr(clazz_module, task_name)
    automated_task: AutomatedTask = clazz(setting_states, callback_before_run_task)
    return automated_task


def prevent_inherit_static_method(func):
    def wrapper(*args, **kwargs):
        import inspect
        current_frame = inspect.currentframe()
        calling_frame = current_frame.f_back
        calling_class = calling_frame.f_locals.get('cls')

        if calling_class:
            method_name = func.__name__
            if method_name in calling_class.__dict__ and isinstance(calling_class.__dict__[method_name], staticmethod):

                base_classes = calling_class.__bases__
                for base_class in base_classes:
                    if method_name in base_class.__dict__ and isinstance(base_class.__dict__[method_name],
                                                                         staticmethod):
                        raise NotImplementedError(
                            f"{method_name} cannot be used in the derived class {calling_class.__name__}")

        return func(*args, **kwargs)

    return wrapper
