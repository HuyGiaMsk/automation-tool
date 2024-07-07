from src.setup.driver.download.chrome.DownloadChromeDriver import DownloadChromeDriver
from src.setup.driver.query.DriverInfoQuery import DriverInfoQuery


class DownloadLinuxChromeDriver(DownloadChromeDriver):

    def __init__(self, driver_info_query: DriverInfoQuery):
        super().__init__(driver_info_query)

    def get_driver_extension(self) -> str:
        return ''

    def get_platform_acronym(self) -> str:
        return 'linux'
