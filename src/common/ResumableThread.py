import threading
from abc import abstractmethod


class ResumableThread(threading.Thread):
    def __init__(self, target=None, group=None, name=None,
                 args=(), kwargs=None, *, daemon=None):

        if target is None:
            target = self.perform

        super().__init__(group=group, target=target, name=name,
                         args=args, kwargs=kwargs, daemon=daemon)
        self.paused = False
        self.terminated = False
        self.pause_condition = threading.Condition(threading.Lock())

    @abstractmethod
    def perform(self):
        pass

    def run(self):
        self._target(*self._args, **self._kwargs)

    def pause(self):
        self.paused = True

    def resume(self):
        with self.pause_condition:
            self.paused = False
            self.pause_condition.notify()

    def terminate(self):
        self.terminated = True
        with self.pause_condition:
            self.paused = False
            self.pause_condition.notify()
