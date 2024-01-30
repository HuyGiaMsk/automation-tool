from abc import abstractmethod, ABC
from src.observer.Event import Event


class EventHandler(ABC):

    @abstractmethod
    def handle_incoming_event(self, event: Event) -> None:
        pass
