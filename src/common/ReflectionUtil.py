import importlib
import os
from types import ModuleType
from typing import Callable

from src.task.AutomatedTask import AutomatedTask


def create_task_instance(setting_states: dict[str, str], task_name: str,
                         callback_before_run_task: Callable[[], None]) -> AutomatedTask:
    clazz_module: ModuleType = importlib.import_module('src.task.' + task_name)
    clazz = getattr(clazz_module, task_name)
    automated_task: AutomatedTask = clazz(setting_states, callback_before_run_task)
    return automated_task

def find_task_module(base_path: str, task_name: str) -> str:
    for root, _, files in os.walk(base_path):
        for file in files:
            if file == f"{task_name}.py":
                relative_path = os.path.relpath(os.path.join(root, file), base_path)
                module_path = relative_path.replace(os.sep, ".").rsplit(".", 1)[0]
                return module_path
    raise FileNotFoundError(f"Task file {task_name}.py not found in {base_path} and its subdirectories.")
