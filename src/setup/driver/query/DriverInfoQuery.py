from abc import abstractmethod, ABC


class DriverInfoQuery(ABC):

    @abstractmethod
    def get_specific_version_from_origin(self, base_version: str) -> str:
        pass

    @abstractmethod
    def get_base_version_from_local(self) -> str:
        pass
