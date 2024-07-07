import platform

from src.setup.driver.DriverType import DriverType
from src.setup.driver.query.DriverInfoQuery import DriverInfoQuery
from src.setup.driver.query.chrome.LinuxChromeDriverInfoQuery import LinuxChromeDriverInfoQuery
from src.setup.driver.query.chrome.WindowsChromeDriverInfoQuery import WindowsChromeDriverInfoQuery

cache: dict[str, DriverInfoQuery] = {}


class DriverInfoQueryFactory:

    @staticmethod
    def get_query(driver_type: DriverType) -> DriverInfoQuery:

        platform_name = platform.system()
        cache_key: str = f"{driver_type}-{platform_name}"
        cache_element: DriverInfoQuery = cache.get(f"{driver_type}-{platform_name}")
        if cache_element is not None:
            return cache_element

        if driver_type is DriverType.SELENIUM:

            instance: DriverInfoQuery | None = None

            if platform_name == 'Windows':
                instance = WindowsChromeDriverInfoQuery()

            if platform_name == 'Linux':
                instance = LinuxChromeDriverInfoQuery()

            if instance is None:
                raise Exception(f'Still no support querying {driver_type} for platform {platform_name}')

            cache[cache_key] = instance
            return instance

        raise Exception(f"Invalid driver type, we don't support querying {driver_type}")
