from abc import ABC

from src.observer.EventBroker import EventBroker
from src.observer.PercentChangedEvent import PercentChangedEvent


class Percentage(ABC):

    def __init__(self):
        super().__init__()
        self.total_element_size = -1
        self.__current_element_count = 0

    @property
    def current_element_count(self) -> int:
        return self.__current_element_count

    @current_element_count.setter
    def current_element_count(self, new_value):
        if new_value <= self.__current_element_count:
            return

        self.__current_element_count = new_value
        EventBroker.get_instance().publish(topic=PercentChangedEvent.event_name,
                                           event=PercentChangedEvent(task_name=type(self).__name__,
                                                                     current_percent=self._get_current_percentage()))

    @property
    def __class__(self):
        return super().__class__

    def _get_current_percentage(self) -> float:
        return self.__current_element_count * 100 / self.total_element_size

    def get_percentage_distance(self) -> float:
        return 100 / self.total_element_size
