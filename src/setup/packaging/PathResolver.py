from abc import ABC, abstractmethod


class PathResolver(ABC):
    @abstractmethod
    def get_executable_directory(self) -> str:
        pass

    @abstractmethod
    def resolve(self, paths: list[str]) -> str:
        pass
