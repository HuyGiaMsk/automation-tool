import platform

from src.setup.packaging.admin.AdminPrivilegeProvider import AdminPrivilegeProvider
from src.setup.packaging.admin.LinuxAdmin import LinuxAdmin
from src.setup.packaging.admin.WindowsAdmin import WindowsAdmin


def validate_admin_privilege():
    platform_name = platform.system()

    provider: AdminPrivilegeProvider = None
    if platform_name == 'Windows':
        provider = WindowsAdmin()

    if platform_name == 'Linux':
        provider = LinuxAdmin()

    if provider is None:
        raise Exception('Unsupported os')

    if not provider.is_currently_running_with_admin():
        provider.provide_admin_privilege()
