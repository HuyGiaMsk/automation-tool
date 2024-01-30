import threading

from src.observer.EventHandler import EventHandler


class EventBroker:
    _instance = None

    @staticmethod
    def get_instance():
        if EventBroker._instance is None:
            EventBroker._instance = EventBroker()
        return EventBroker._instance

    def __init__(self):
        self.topicToSetOfObserver: dict[str, set[EventHandler]] = {}

    def subscribe(self, topic: str, observer: EventHandler) -> None:
        if topic not in self.topicToSetOfObserver:
            self.topicToSetOfObserver[topic] = set()
        self.topicToSetOfObserver.get(topic).add(observer)

    def publish(self, topic: str, event: threading.Event) -> None:
        observers: set[EventHandler] = self.topicToSetOfObserver.get(topic, set())
        observer: EventHandler
        for observer in observers:
            thread: threading.Thread = threading.Thread(target=observer.handle_incoming_event,
                                                        args=[event],
                                                        daemon=False)
            thread.start()
