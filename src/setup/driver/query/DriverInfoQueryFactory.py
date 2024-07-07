import platform

from src.setup.driver.DriverType import DriverType
from src.setup.driver.query.DriverInfoQuery import DriverInfoQuery
from src.setup.driver.query.chrome.LinuxChromeDriverInfoQuery import LinuxChromeDriverInfoQuery
from src.setup.driver.query.chrome.WindowsChromeDriverInfoQuery import WindowsChromeDriverInfoQuery


class DriverInfoQueryFactory:

    @staticmethod
    def get_query(driver_type: DriverType) -> DriverInfoQuery:

        platform_name = platform.system()

        if driver_type is DriverType.SELENIUM:

            if platform_name == 'Windows':
                return WindowsChromeDriverInfoQuery()

            if platform_name == 'Linux':
                return LinuxChromeDriverInfoQuery()

            raise Exception(f'Still no support querying {driver_type} for platform {platform_name}')

        raise Exception(f"Invalid driver type, we don't support querying {driver_type}")
