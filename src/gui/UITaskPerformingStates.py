from abc import ABC, abstractmethod

from src.task.AutomatedTask import AutomatedTask


class UITaskPerformingStates(ABC):

    @abstractmethod
    def get_ui_settings(self) -> dict[str, str]:
        pass

    @abstractmethod
    def set_ui_settings(self, new_ui_setting_values: dict[str, str]) -> dict[str, str]:
        pass

    @abstractmethod
    def get_task_name(self) -> str:
        pass

    @abstractmethod
    def get_task_instance(self) -> AutomatedTask:
        pass
