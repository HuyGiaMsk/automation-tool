from abc import ABC, abstractmethod

from src.setup.driver.query.DriverInfoQuery import DriverInfoQuery


class DownloadDriver(ABC):

    def __init__(self, driver_info_query: DriverInfoQuery):
        self.driver_info_query: DriverInfoQuery = driver_info_query

    @abstractmethod
    def download_and_place_suitable_version_driver(self) -> None:
        pass

    @abstractmethod
    def get_expected_driver_abs_path(self) -> str:
        pass
