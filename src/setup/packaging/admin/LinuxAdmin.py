import os
import subprocess
import sys
from logging import Logger

from src.common.ThreadLocalLogger import get_current_logger
from src.setup.packaging.admin.AdminPrivilegeProvider import AdminPrivilegeProvider


class LinuxAdmin(AdminPrivilegeProvider):

    def is_currently_running_with_admin(self) -> bool:
        return os.geteuid() == 0

    def provide_admin_privilege(self) -> None:

        logger: Logger = get_current_logger()

        if not self.is_currently_running_with_admin():
            logger.info("Executable instance is not running with root privileges. Attempting to relaunch with sudo...")
            try:
                script = os.path.abspath(sys.argv[0])
                params = ' '.join([sys.executable] + [script] + sys.argv[1:])
                subprocess.call(['sudo'] + params.split())
                sys.exit(0)
            except Exception as e:
                logger.error(f"Failed to relaunch script with sudo: {e}")
                sys.exit(1)
        else:
            logger.warning("Script is already running with root privileges.")
