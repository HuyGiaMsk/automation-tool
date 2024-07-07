import inspect
import platform
from abc import ABC
from types import NoneType

from src.setup.driver.DriverType import DriverType
from src.setup.driver.download.DownloadDriver import DownloadDriver
from src.setup.driver.download.chrome.DownloadLinuxChromeDriver import DownloadLinuxChromeDriver
from src.setup.driver.download.chrome.DownloadWindowsChromeDriver import DownloadWindowsChromeDriver
from src.setup.driver.query.DriverInfoQuery import DriverInfoQuery
from src.setup.driver.query.DriverInfoQueryFactory import DriverInfoQueryFactory

cache: dict[str, DownloadDriver] = {}


class DownloadDriverFactory(ABC):

    @staticmethod
    def get_downloader() -> DownloadDriver:
        driver_type: DriverType = DownloadDriverFactory.__get_driver_type()
        query: DriverInfoQuery = DriverInfoQueryFactory.get_query(driver_type=driver_type)
        platform_name = platform.system()

        cache_key: str = f"{driver_type}-{platform_name}"
        cache_element: DownloadDriver = cache.get(f"{driver_type}-{platform_name}")
        if cache_element is not None:
            return cache_element

        if driver_type is DriverType.SELENIUM:

            instance: DownloadDriver | None = None

            if platform_name == 'Windows':
                instance: DownloadDriver = DownloadWindowsChromeDriver(driver_info_query=query)

            if platform_name == 'Linux':
                instance: DownloadDriver = DownloadLinuxChromeDriver(driver_info_query=query)

            if instance is None:
                raise Exception(f'Still no support downloading {driver_type} for platform {platform_name}')

            cache[cache_key] = instance
            return instance

        raise Exception(f"Invalid driver type, we don't support downloading {driver_type}")

    @staticmethod
    def __get_caller_of_factory() -> type | None:
        current_frame = inspect.currentframe()

        # Traverse the call stack to find the caller
        while current_frame:
            caller_frame = current_frame.f_back
            if caller_frame:
                caller_type: type = type(caller_frame.f_locals.get('self', None))
                if caller_type is not NoneType and caller_type is not DownloadDriverFactory:
                    return caller_type

            current_frame = caller_frame

        return None

    @staticmethod
    def __get_driver_type() -> DriverType:
        from src.task.WebTask import WebTask
        caller_type: type = DownloadDriverFactory.__get_caller_of_factory()
        if caller_type is None:
            return DriverType.NOT_SPECIFIED

        if issubclass(caller_type, WebTask):
            return DriverType.SELENIUM

        return DriverType.NOT_SPECIFIED
